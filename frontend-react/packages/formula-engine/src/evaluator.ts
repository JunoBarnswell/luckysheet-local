import type { BinaryOperator, CellAddress, FormulaAst, FormulaReferenceNode, SpillReferenceNode } from './ast';
import { formatFormula } from './ast-format';
import { collectFormulaDependencies, resolveCellReference, resolveRangeReference } from './dependencies';
import { getBuiltinFunction, getFunctionCapability } from './functions';
import { evaluateAdvancedFunction, type AdvancedFunctionArgs } from './functions/advanced';
import { parseFormula } from './parser';
import type { FormulaDependency, RangeDependency } from './range-index';
import { createFormulaError, isArrayValue, isFormulaError, isReferenceValue, type ArrayValue, type FormulaError, type FormulaValue, type ScalarValue } from './values';
import { coerceExcelNumber, normalizeExcelPrecision } from './numeric';
import type { ExcelNumericContext } from './numeric';
import type { WorkbookCollationContext } from './collation';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';
import { createReferenceCursor, type ReferenceFormulaKind, type RowVisibilityResolver } from './reference-cursor';

export interface FormulaEvaluationContext {
  readonly currentCell: CellAddress;
  readCell(address: CellAddress): FormulaValue;
  readRange(range: RangeDependency): Iterable<FormulaValue>;
  readRangeMatrix?(range: RangeDependency): ArrayValue;
  /** Resolve a dynamic-array anchor to its current spill range. */
  readSpillRange?(address: CellAddress): RangeDependency | undefined;
  /** Read a projected value from a dynamic-array spill cell. */
  readSpillValue?(address: CellAddress): FormulaValue | undefined;
  /** 定义名称解析:返回 undefined 视为 #NAME? */
  resolveName?(name: string): FormulaValue | undefined;
  /** 结构化表引用解析 */
  resolveTableReference?(tableName: string, request: {
    specifier?: import('./ast').TableReferenceSpecifier;
    columnName?: string;
    columnEndName?: string;
    thisRow: boolean;
  }): FormulaValue | EvaluationRange | undefined;
  /** Resolve a structured reference without converting it to an opaque string. */
  resolveReference?(reference: FormulaReferenceNode): FormulaEvaluationReference | FormulaError | undefined;
  /** Workbook calendar used by serial/date functions. */
  readonly dateSystem?: ExcelDateSystem;
  /** Host-owned deterministic clock; missing means TODAY/NOW fail-close. */
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Workbook numeric semantics shared by inline and Worker evaluation. */
  readonly numericContext?: ExcelNumericContext;
  readonly collationContext?: WorkbookCollationContext;
  /** Canonical worksheet row visibility used by provenance-aware references. */
  readonly rowVisibility?: RowVisibilityResolver;
  /** Formula identity for a source cell, used to suppress nested aggregates. */
  readonly readFormulaKind?: (address: CellAddress) => ReferenceFormulaKind;
  /** Stable AST identity for the current function occurrence. */
  readonly volatileOccurrence?: string;
  /** Host-provided order-independent random source for volatile functions. */
  readonly random?: (functionName: string, occurrence?: string, elementIndex?: number) => number | FormulaError;
  /** Evaluate a formula AST while overriding one or more input cells. */
  readonly evaluateWithCellOverrides?: (ast: FormulaAst, overrides: readonly FormulaCellOverride[]) => FormulaValue;
  /**
   * Dynamic references are discovered during evaluation (INDIRECT/OFFSET).
   * They are added to the same dependency index as static AST references;
   * callers must not maintain a second dependency graph.
   */
  readonly registerDynamicDependency?: (dependency: FormulaDependency) => void;
}

export interface FormulaCellOverride {
  readonly address: CellAddress;
  readonly value: ScalarValue;
}

/** A data-only evaluation step used by Formula Auditing's Evaluate Formula view. */
export interface FormulaEvaluationTraceStep {
  readonly node: FormulaAst;
  readonly expression: string;
  readonly value: FormulaValue;
}

export interface FormulaEvaluationTrace {
  readonly value: FormulaValue;
  readonly steps: readonly FormulaEvaluationTraceStep[];
}

interface EvaluationRange {
  readonly kind: 'range';
  readonly range: RangeDependency;
}

export interface FormulaEvaluationReference {
  readonly kind: 'reference';
  readonly ranges: readonly RangeDependency[];
}

type EvaluationValue = FormulaValue | EvaluationRange | FormulaEvaluationReference;

export function evaluateFormula(ast: FormulaAst, context: FormulaEvaluationContext): FormulaValue {
  const result = evaluateNode(ast, context);
  return materializeEvaluationValue(result, context);
}

/** Evaluate a formula and capture every AST node's computed value in order. */
export function evaluateFormulaWithTrace(ast: FormulaAst, context: FormulaEvaluationContext): FormulaEvaluationTrace {
  const steps: FormulaEvaluationTraceStep[] = [];
  const result = evaluateNode(ast, context, (node, value) => {
    steps.push({
      node: structuredClone(node),
      expression: formatFormula(node),
      value: structuredClone(materializeEvaluationValue(value, context)),
    });
  });
  return { value: structuredClone(materializeEvaluationValue(result, context)), steps };
}

type EvaluationTraceSink = (node: FormulaAst, value: EvaluationValue) => void;

function materializeEvaluationValue(value: EvaluationValue, context: FormulaEvaluationContext): FormulaValue {
  if (isReferenceValue(value)) return createFormulaError('#REF!', 'Reference value requires a workbook resolver');
  if (!isEvaluationRange(value) && !isEvaluationReference(value)) return value;
  const matrix = isEvaluationRange(value)
    ? readRangeAsMatrix(value.range, context)
    : readReferenceAsMatrix(value.ranges, context);
  return matrix.length === 1 && matrix[0]?.length === 1 ? matrix[0][0]! : matrix;
}

function evaluateNode(node: FormulaAst, context: FormulaEvaluationContext, trace?: EvaluationTraceSink): EvaluationValue {
  let result: EvaluationValue = createFormulaError('#VALUE!', 'Unsupported formula node');
  switch (node.type) {
    case 'number-literal':
      result = node.value;
      break;
    case 'string-literal':
      result = node.value;
      break;
    case 'boolean-literal':
      result = node.value;
      break;
    case 'invalid-reference':
      result = createFormulaError('#REF!', 'Reference was deleted by a structural mutation');
      break;
    case 'cell-reference':
      result = context.readCell(resolveCellReference(node.reference, context.currentCell));
      break;
    case 'range-reference':
      result = { kind: 'range', range: resolveRangeReference(node, context.currentCell) };
      break;
    case 'whole-column-reference':
    case 'whole-row-reference':
    case 'reference-union':
    case 'reference-intersection':
    case 'sheet-range-reference':
    case 'external-reference': {
      const resolved = context.resolveReference?.(node);
      result = resolved ?? createFormulaError('#REF!', 'Reference requires a workbook resolver');
      break;
    }
    case 'spill-reference':
      result = evaluateSpillReference(node, context);
      break;
    case 'unary-expression':
      result = evaluateUnary(node.operator, evaluateNode(node.operand, context, trace), context);
      break;
    case 'binary-expression':
      result = evaluateBinary(
        node.operator,
        evaluateNode(node.left, context, trace),
        evaluateNode(node.right, context, trace),
        context,
      );
      break;
    case 'function-call':
      result = evaluateFunction(node.name, node.arguments, context, trace, `${node.span.start}:${node.span.end}`);
      break;
    case 'name-reference': {
      const resolved = context.resolveName?.(node.name.toUpperCase());
      result = resolved === undefined ? createFormulaError('#NAME?', 'Unknown name: ' + node.name) : resolved;
      break;
    }
    case 'table-reference': {
      const resolved = context.resolveTableReference?.(node.tableName, {
        specifier: node.specifier,
        columnName: node.columnName,
        columnEndName: node.columnEndName,
        thisRow: node.thisRow,
      });
      if (resolved === undefined) {
        const label = node.specifier
          ? `#${node.specifier}`
          : node.columnName ?? '';
        result = createFormulaError('#NAME?', `Unknown table reference: ${node.tableName}[${label}]`);
        break;
      }
      result = resolved;
      break;
    }
  }
  trace?.(node, result);
  return result;
}

function evaluateUnary(operator: '+' | '-' | '%' | '@', operand: EvaluationValue, context: FormulaEvaluationContext): FormulaValue {
  if (isFormulaError(operand)) return operand;
  if (operator === '@' && (isEvaluationRange(operand) || isEvaluationReference(operand))) {
    const range = isEvaluationRange(operand) ? operand.range : operand.ranges[0];
    if (!range) return createFormulaError('#REF!', 'Implicit intersection has no range');
    const { start, end } = range;
    const row = context.currentCell.row >= start.row && context.currentCell.row <= end.row ? context.currentCell.row : start.row;
    const column = context.currentCell.column >= start.column && context.currentCell.column <= end.column ? context.currentCell.column : start.column;
    const address = { sheetId: start.sheetId, row, column };
    return context.readSpillValue?.(address) ?? context.readCell(address);
  }
  if (isEvaluationRange(operand) || isEvaluationReference(operand)) {
    const matrix = isEvaluationRange(operand) ? readRangeAsMatrix(operand.range, context) : readReferenceAsMatrix(operand.ranges, context);
    return matrix.map((row) => row.map((value) => evaluateUnary(operator, value, context)));
  }
  if (isArrayValue(operand)) return operand.map((row) => row.map((value) => evaluateUnary(operator, value, context)));
  const number = toNumber(operand);
  if (isFormulaError(number)) return number;
  if (operator === '-') return normalizeExcelPrecision(-number, context.numericContext);
  if (operator === '%') return normalizeExcelPrecision(number / 100, context.numericContext);
  return normalizeExcelPrecision(number, context.numericContext);
}

function evaluateSpillReference(node: SpillReferenceNode, context: FormulaEvaluationContext): EvaluationValue {
  const operand = node.operand;
  const anchor = operand.type === 'cell-reference'
    ? resolveCellReference(operand.reference, context.currentCell)
    : operand.type === 'range-reference'
      ? resolveRangeReference(operand, context.currentCell).start
      : undefined;
  if (!anchor) return createFormulaError('#REF!', 'Spill operator expects a cell or range reference');
  context.readCell(anchor);
  const range = context.readSpillRange?.(anchor);
  return range ? { kind: 'range', range } : createFormulaError('#REF!', 'Reference does not resolve to a spill range');
}

function evaluateBinary(
  operator: BinaryOperator,
  left: EvaluationValue,
  right: EvaluationValue,
  context: FormulaEvaluationContext,
): FormulaValue {
  if (isFormulaError(left)) return left;
  if (isFormulaError(right)) return right;
  const leftValue = isEvaluationRange(left) || isEvaluationReference(left) ? materializeEvaluationValue(left, context) : left;
  const rightValue = isEvaluationRange(right) || isEvaluationReference(right) ? materializeEvaluationValue(right, context) : right;
  if (isArrayValue(leftValue) || isArrayValue(rightValue)) return liftBinary(operator, leftValue, rightValue, context);
  left = leftValue;
  right = rightValue;

  // String concatenation
  if (operator === '&') {
    const leftStr = left === null ? '' : String(left);
    const rightStr = right === null ? '' : String(right);
    return leftStr + rightStr;
  }

  // Comparisons
  if (
    operator === '=' ||
    operator === '<>' ||
    operator === '<' ||
    operator === '<=' ||
    operator === '>' ||
    operator === '>='
  ) {
    return compareValues(left, right, operator);
  }

  // Arithmetic
  const leftNumber = toNumber(left);
  if (isFormulaError(leftNumber)) return leftNumber;
  const rightNumber = toNumber(right);
  if (isFormulaError(rightNumber)) return rightNumber;

  switch (operator) {
    case '+':
      return normalizeExcelPrecision(leftNumber + rightNumber, context.numericContext);
    case '-':
      return normalizeExcelPrecision(leftNumber - rightNumber, context.numericContext);
    case '*':
      return normalizeExcelPrecision(leftNumber * rightNumber, context.numericContext);
    case '/':
      return rightNumber === 0 ? createFormulaError('#DIV/0!', 'Division by zero') : normalizeExcelPrecision(leftNumber / rightNumber, context.numericContext);
    case '^':
      return normalizeExcelPrecision(Math.pow(leftNumber, rightNumber), context.numericContext);
  }
}

function liftBinary(operator: BinaryOperator, left: FormulaValue, right: FormulaValue, context: FormulaEvaluationContext): FormulaValue {
  const leftMatrix = isArrayValue(left) ? left : [[left]];
  const rightMatrix = isArrayValue(right) ? right : [[right]];
  const rows = Math.max(leftMatrix.length, rightMatrix.length);
  const columns = Math.max(leftMatrix[0]?.length ?? 1, rightMatrix[0]?.length ?? 1);
  const compatible = (matrix: ArrayValue): boolean =>
    (matrix.length === 1 || matrix.length === rows)
    && ((matrix[0]?.length ?? 1) === 1 || (matrix[0]?.length ?? 1) === columns);
  if (!compatible(leftMatrix) || !compatible(rightMatrix)) return createFormulaError('#VALUE!', 'Array shapes are not compatible');
  const valueAt = (matrix: ArrayValue, row: number, column: number): FormulaValue => {
    const sourceRow = matrix.length === 1 ? 0 : row;
    const sourceColumns = matrix[sourceRow]?.length ?? 1;
    return matrix[sourceRow]?.[sourceColumns === 1 ? 0 : column] ?? null;
  };
  return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) =>
    evaluateBinary(operator, valueAt(leftMatrix, row, column), valueAt(rightMatrix, row, column), context)));
}

function compareValues(left: FormulaValue, right: FormulaValue, operator: BinaryOperator): boolean {
  if (left === null && right === null) {
    return operator === '=' || operator === '<=' || operator === '>=';
  }
  if (typeof left === 'number' && typeof right === 'number') {
    switch (operator) {
      case '=': return left === right;
      case '<>': return left !== right;
      case '<': return left < right;
      case '<=': return left <= right;
      case '>': return left > right;
      case '>=': return left >= right;
    }
  }

  const sLeft = String(left ?? '').toLowerCase();
  const sRight = String(right ?? '').toLowerCase();
  switch (operator) {
    case '=': return sLeft === sRight;
    case '<>': return sLeft !== sRight;
    case '<': return sLeft < sRight;
    case '<=': return sLeft <= sRight;
    case '>': return sLeft > sRight;
    case '>=': return sLeft >= sRight;
  }
  return false;
}

function evaluateFunction(
  name: string,
  argumentsList: readonly FormulaAst[],
  context: FormulaEvaluationContext,
  trace?: EvaluationTraceSink,
  volatileOccurrence?: string,
): FormulaValue | EvaluationRange | FormulaEvaluationReference {
  // 需要原始 AST / 返回区间的引用类函数:在求值器内原生实现
  const native = evaluateReferenceFunction(name, argumentsList, context, trace);
  if (native !== undefined) return native;

  const lazy = evaluateLazyFunction(name, argumentsList, context, trace, volatileOccurrence);
  if (lazy !== undefined) return lazy;

  const fn = getBuiltinFunction(name);
  const evaluatedArgs: FormulaValue[] = [];
  const rawRanges: EvaluationValue[] = [];
  for (let index = 0; index < argumentsList.length; index += 1) {
    const argument = argumentsList[index]!;
    const aggregate = aggregateIdentifierArgument(name, index, argument);
    if (aggregate !== undefined) {
      evaluatedArgs.push(aggregate);
      rawRanges.push(aggregate);
      continue;
    }
    const value = evaluateNode(argument, context, trace);
    rawRanges.push(value);
    if (isEvaluationRange(value) || isEvaluationReference(value)) {
      evaluatedArgs.push(materializeEvaluationValue(value, context));
    } else {
      evaluatedArgs.push(value);
    }
  }

  if (fn) {
    try {
      return fn(evaluatedArgs, volatileOccurrence === undefined ? context : { ...context, volatileOccurrence });
    } catch (err) {
      return createFormulaError('#VALUE!', err instanceof Error ? err.message : 'Function evaluation error');
    }
  }

  // 上下文感知函数(SUMIFS 家族 / SUMPRODUCT / SUBTOTAL 等)
  const advanced = evaluateAdvancedFunction(name, { values: evaluatedArgs, ranges: rawRanges } as AdvancedFunctionArgs, {
    toRanges: (value: EvaluationValue) => isEvaluationRange(value)
      ? [value.range]
      : isEvaluationReference(value)
        ? value.ranges
        : undefined,
    readCursor: (range: RangeDependency) => createReferenceCursor(range, context),
  });
  if (advanced !== undefined) return advanced;

  const capability = getFunctionCapability(name);
  return createFormulaError('#NAME?', capability.status === 'unsupported'
    ? `UNSUPPORTED_FUNCTION: ${capability.id}`
    : `Function ${capability.id} is not executable in the current evaluator`);
}

/**
 * Evaluate control-flow functions without materialising branches that cannot
 * affect the result. Besides matching Excel semantics this prevents expensive
 * ranges, volatile calls and invalid references in an unselected branch from
 * polluting the current dependency closure.
 */
function evaluateLazyFunction(
  name: string,
  args: readonly FormulaAst[],
  context: FormulaEvaluationContext,
  trace?: EvaluationTraceSink,
  volatileOccurrence?: string,
): FormulaValue | undefined {
  const normalized = name.trim().toUpperCase();
  const evaluate = (index: number): EvaluationValue => evaluateNode(args[index]!, context, trace);
  const materialize = (value: EvaluationValue): FormulaValue => isEvaluationRange(value) || isEvaluationReference(value)
    ? materializeEvaluationValue(value, context)
    : value;
  const truthy = (value: FormulaValue): boolean => {
    if (isFormulaError(value)) return false;
    if (value === null || value === false || value === '') return false;
    if (typeof value === 'number') return value !== 0;
    if (isArrayValue(value)) return truthy(value[0]?.[0] ?? null);
    return true;
  };

  if (normalized === 'IF') {
    if (args.length < 2 || args.length > 3) return createFormulaError('#VALUE!', 'IF requires two or three arguments');
    const condition = materialize(evaluate(0));
    if (isFormulaError(condition)) return condition;
    const selected = truthy(condition) ? 1 : 2;
    if (selected >= args.length) return false;
    return materialize(evaluate(selected));
  }

  if (normalized === 'IFERROR' || normalized === 'IFNA') {
    if (args.length !== 2) return createFormulaError('#VALUE!', `${normalized} requires two arguments`);
    const first = materialize(evaluate(0));
    const shouldReplace = isFormulaError(first) && (normalized === 'IFERROR' || first.code === '#N/A');
    return shouldReplace ? materialize(evaluate(1)) : first;
  }

  if (normalized === 'IFS') {
    if (args.length === 0 || args.length % 2 !== 0) return createFormulaError('#VALUE!', 'IFS requires pairs of condition and value');
    for (let index = 0; index < args.length; index += 2) {
      const condition = materialize(evaluate(index));
      if (isFormulaError(condition)) return condition;
      if (truthy(condition)) return materialize(evaluate(index + 1));
    }
    return createFormulaError('#N/A', 'No condition matched in IFS');
  }

  if (normalized === 'SWITCH') {
    if (args.length < 3) return createFormulaError('#VALUE!', 'SWITCH requires target, value, result');
    const target = materialize(evaluate(0));
    if (isFormulaError(target)) return target;
    const hasDefault = args.length % 2 === 0;
    const pairEnd = hasDefault ? args.length - 1 : args.length;
    for (let index = 1; index + 1 < pairEnd; index += 2) {
      const candidate = materialize(evaluate(index));
      if (isFormulaError(candidate)) return candidate;
      if (compareValues(target, candidate, '=')) return materialize(evaluate(index + 1));
    }
    return hasDefault ? materialize(evaluate(args.length - 1)) : createFormulaError('#N/A', 'No matching case in SWITCH');
  }

  if (normalized === 'CHOOSE') {
    if (args.length < 2) return createFormulaError('#VALUE!', 'CHOOSE requires an index and a value');
    const indexValue = materialize(evaluate(0));
    const index = toNumber(indexValue);
    if (isFormulaError(index) || !Number.isInteger(index) || index < 1 || index >= args.length) return createFormulaError('#VALUE!', 'CHOOSE index is out of bounds');
    return materialize(evaluate(index));
  }

  if (normalized === 'AND' || normalized === 'OR') {
    if (args.length === 0) return createFormulaError('#VALUE!', `${normalized} requires arguments`);
    const values: FormulaValue[] = [];
    for (let index = 0; index < args.length; index += 1) {
      const value = materialize(evaluate(index));
      if (isFormulaError(value)) return value;
      values.push(value);
      if (normalized === 'AND' && !truthy(value)) return false;
      if (normalized === 'OR' && truthy(value)) return true;
    }
    return normalized === 'AND' ? values.every(truthy) : values.some(truthy);
  }

  if (normalized === 'XLOOKUP') {
    if (args.length < 3 || args.length > 6) return createFormulaError('#VALUE!', 'XLOOKUP requires three to six arguments');
    const values: FormulaValue[] = [];
    for (let index = 0; index < Math.min(3, args.length); index += 1) values.push(materialize(evaluate(index)));
    const fn = getBuiltinFunction(normalized);
    if (!fn) return createFormulaError('#NAME?', `Unknown function: ${name}`);
    let result: FormulaValue;
    try { result = fn(values, volatileOccurrence === undefined ? context : { ...context, volatileOccurrence }); }
    catch (error) { return createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Function evaluation error'); }
    if (isFormulaError(result) && result.code === '#N/A' && args[3]) {
      const fallback = materialize(evaluate(3));
      values.push(fallback);
      try { result = fn(values, volatileOccurrence === undefined ? context : { ...context, volatileOccurrence }); }
      catch (error) { return createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Function evaluation error'); }
    }
    return result;
  }

  return undefined;
}

function aggregateIdentifierArgument(
  functionName: string,
  argumentIndex: number,
  argument: FormulaAst,
): 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX' | undefined {
  const normalizedFunction = functionName.trim().toUpperCase();
  const aggregateIndex = normalizedFunction === 'GROUPBY'
    ? 2
    : normalizedFunction === 'PIVOTBY'
      ? 3
      : -1;
  if (argumentIndex !== aggregateIndex || argument.type !== 'name-reference') return undefined;
  const aggregate = argument.name.trim().toUpperCase();
  return aggregate === 'SUM' || aggregate === 'COUNT' || aggregate === 'AVERAGE' || aggregate === 'MIN' || aggregate === 'MAX'
    ? aggregate
    : undefined;
}

/** ROW / COLUMN / ADDRESS / OFFSET / INDIRECT:需要 AST 或返回区间引用 */
function evaluateReferenceFunction(
  name: string,
  args: readonly FormulaAst[],
  context: FormulaEvaluationContext,
  trace?: EvaluationTraceSink,
): FormulaValue | EvaluationRange | FormulaEvaluationReference | undefined {
  switch (name.toUpperCase()) {
    case 'SJS.TABLE':
      return evaluateSjsTable(args, context, trace);
    case 'ROW': {
      if (args.length === 0) return context.currentCell.row + 1;
      const target = args[0]!;
      if (target.type === 'cell-reference') {
        return resolveCellReference(target.reference, context.currentCell).row + 1;
      }
      if (target.type === 'range-reference') {
        return resolveRangeReference(target, context.currentCell).start.row + 1;
      }
      return createFormulaError('#VALUE!', 'ROW expects a reference');
    }
    case 'COLUMN': {
      if (args.length === 0) return context.currentCell.column + 1;
      const target = args[0]!;
      if (target.type === 'cell-reference') {
        return resolveCellReference(target.reference, context.currentCell).column + 1;
      }
      if (target.type === 'range-reference') {
        return resolveRangeReference(target, context.currentCell).start.column + 1;
      }
      return createFormulaError('#VALUE!', 'COLUMN expects a reference');
    }
    case 'ADDRESS': {
      const values: FormulaValue[] = [];
      for (const argument of args) {
        const value = evaluateNode(argument, context, trace);
        values.push(isEvaluationRange(value) || isEvaluationReference(value) ? createFormulaError('#VALUE!', 'ADDRESS expects scalars') : value);
      }
      const row = toNumber(values[0] ?? 1);
      const column = toNumber(values[1] ?? 1);
      if (isFormulaError(row) || isFormulaError(column)) return createFormulaError('#VALUE!', 'Invalid ADDRESS arguments');
      const absMode = toNumber(values[2] ?? 1);
      if (isFormulaError(absMode)) return createFormulaError('#VALUE!', 'Invalid ADDRESS abs mode');
      let columnLetter = '';
      let remaining = column;
      while (remaining > 0) {
        const modulo = (remaining - 1) % 26;
        columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
        remaining = Math.floor((remaining - 1) / 26);
      }
      const absolute = (mode: number) => (mode === 1 || mode === 2 ? '$' : '');
      const rowPart = absMode === 1 || absMode === 3 ? '$' : '';
      return absolute(absMode as number) + columnLetter + rowPart + String(row);
    }
    case 'OFFSET': {
      const base = args[0];
      if (!base || (base.type !== 'cell-reference' && base.type !== 'range-reference')) {
        return createFormulaError('#VALUE!', 'OFFSET expects a reference');
      }
      const scalar = (node: FormulaAst | undefined, fallback: number): number | FormulaError => {
        if (!node) return fallback;
        const value = evaluateNode(node, context, trace);
        if (isEvaluationRange(value) || isEvaluationReference(value)) return createFormulaError('#VALUE!', 'OFFSET offset must be scalar');
        const numeric = toNumber(value);
        return numeric;
      };
      const rows = scalar(args[1], 0);
      const columns = scalar(args[2], 0);
      const height = scalar(args[3], 1);
      const width = scalar(args[4], 1);
      for (const candidate of [rows, columns, height, width]) {
        if (typeof candidate !== 'number') return candidate;
      }
      const anchorRange = base.type === 'range-reference' ? resolveRangeReference(base, context.currentCell) : undefined;
      const anchorCell = base.type === 'cell-reference' ? resolveCellReference(base.reference, context.currentCell) : anchorRange!.start;
      const startRow = anchorCell.row + (rows as number);
      const startColumn = anchorCell.column + (columns as number);
      const endRow = startRow + Math.max(1, height as number) - 1;
      const endColumn = startColumn + Math.max(1, width as number) - 1;
      if (startRow < 0 || startColumn < 0) return createFormulaError('#REF!', 'OFFSET out of bounds');
      const dynamicRange: RangeDependency = {
        kind: 'range',
        start: { sheetId: anchorCell.sheetId, row: startRow, column: startColumn },
        end: { sheetId: anchorCell.sheetId, row: endRow, column: endColumn },
      };
      context.registerDynamicDependency?.(dynamicRange);
      return {
        kind: 'range',
        range: dynamicRange,
      };
    }
    case 'INDIRECT': {
      const first = args[0];
      if (!first) return createFormulaError('#REF!', 'INDIRECT expects a text reference');
      const value = evaluateNode(first, context, trace);
      if (isEvaluationRange(value) || isEvaluationReference(value)) return createFormulaError('#VALUE!', 'INDIRECT expects text');
      if (typeof value !== 'string') return createFormulaError('#REF!', 'INDIRECT text required');
      try {
        const parsed = parseFormula('=' + value);
        for (const dependency of collectFormulaDependencies(parsed, context.currentCell)) context.registerDynamicDependency?.(dependency);
        const resolved = evaluateNode(parsed, context, trace);
        return resolved;
      } catch {
        return createFormulaError('#REF!', 'INDIRECT cannot parse: ' + value);
      }
    }
    default:
      return undefined;
  }
}

/**
 * SJS.TABLE is the canonical dynamic-array data-table function.  The first
 * argument is the result expression/reference, followed by one or more
 * `(inputs, inputCell)` pairs.  Each input range is sampled in row-major
 * order and the result expression is evaluated with the corresponding input
 * cells overridden for that row.
 */
function evaluateSjsTable(
  args: readonly FormulaAst[],
  context: FormulaEvaluationContext,
  trace?: EvaluationTraceSink,
): FormulaValue {
  if (!context.evaluateWithCellOverrides) return createFormulaError('#BLOCKED!', 'SJS.TABLE requires an override-capable workbook evaluator');
  if (args.length < 3 || (args.length - 1) % 2 !== 0) {
    return createFormulaError('#VALUE!', 'SJS.TABLE expects resultReference and one or more input pairs');
  }

  const inputPairs: Array<{ values: FormulaValue[]; address: CellAddress }> = [];
  for (let index = 1; index < args.length; index += 2) {
    const inputNode = args[index]!;
    const cellNode = args[index + 1]!;
    const address = referenceCellAddress(cellNode, context);
    if (!address) return createFormulaError('#VALUE!', 'SJS.TABLE inputCell must be a cell reference');
    const values = referenceValues(inputNode, context, trace);
    if (isFormulaError(values)) return values;
    inputPairs.push({ values, address });
  }
  const rowCount = inputPairs[0]?.values.length ?? 0;
  if (rowCount === 0 || inputPairs.some((pair) => pair.values.length !== rowCount)) {
    return createFormulaError('#VALUE!', 'SJS.TABLE input ranges must have the same number of values');
  }

  const result: ArrayValue = [];
  for (let row = 0; row < rowCount; row += 1) {
    const overrides = inputPairs.map((pair) => ({ address: pair.address, value: scalarForTable(pair.values[row]!) }));
    const value = context.evaluateWithCellOverrides(args[0]!, overrides);
    trace?.(args[0]!, value as EvaluationValue);
    if (isFormulaError(value)) return value;
    if (isArrayValue(value) || isReferenceValue(value)) return createFormulaError('#CALC!', 'SJS.TABLE resultReference must resolve to a scalar');
    result.push([value]);
  }
  return result;
}

function referenceCellAddress(node: FormulaAst, context: FormulaEvaluationContext): CellAddress | undefined {
  if (node.type !== 'cell-reference') return undefined;
  return resolveCellReference(node.reference, context.currentCell);
}

function referenceValues(node: FormulaAst, context: FormulaEvaluationContext, trace?: EvaluationTraceSink): FormulaValue[] | FormulaError {
  const value = evaluateNode(node, context, trace);
  if (isFormulaError(value)) return value;
  if (isEvaluationRange(value)) return [...context.readRange(value.range)];
  if (isEvaluationReference(value)) return value.ranges.flatMap((range) => [...context.readRange(range)]);
  if (isArrayValue(value)) return value.flat();
  if (isReferenceValue(value)) return createFormulaError('#VALUE!', 'SJS.TABLE inputs must be cell or range references');
  return [value];
}

function scalarForTable(value: FormulaValue): ScalarValue {
  if (isFormulaError(value)) return null;
  if (isArrayValue(value) || isReferenceValue(value)) return null;
  return value;
}

function readRangeAsMatrix(range: RangeDependency, context: FormulaEvaluationContext): ArrayValue {
  if (context.readRangeMatrix) {
    return context.readRangeMatrix(range);
  }
  const matrix: ArrayValue = [];
  for (let row = range.start.row; row <= range.end.row; row++) {
    const rowList: FormulaValue[] = [];
    for (let column = range.start.column; column <= range.end.column; column++) {
      rowList.push(context.readCell({ sheetId: range.start.sheetId, row, column }));
    }
    matrix.push(rowList);
  }
  return matrix;
}

function toNumber(value: FormulaValue): number | ReturnType<typeof createFormulaError> {
  return coerceExcelNumber(value);
}

function isEvaluationRange(value: EvaluationValue): value is EvaluationRange {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'range';
}

function isEvaluationReference(value: EvaluationValue): value is FormulaEvaluationReference {
  return typeof value === 'object'
    && value !== null
    && 'kind' in value
    && (value as { kind: string }).kind === 'reference'
    && 'ranges' in value;
}

function readReferenceAsMatrix(ranges: readonly RangeDependency[], context: FormulaEvaluationContext): ArrayValue {
  const matrix: ArrayValue = [];
  for (const range of ranges) matrix.push(...readRangeAsMatrix(range, context));
  return matrix;
}
