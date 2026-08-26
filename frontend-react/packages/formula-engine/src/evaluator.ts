import type { BinaryOperator, CellAddress, FormulaAst, SpillReferenceNode } from './ast';
import { formatFormula } from './ast-format';
import { resolveCellReference, resolveRangeReference } from './dependencies';
import { getBuiltinFunction } from './functions';
import { evaluateAdvancedFunction, type AdvancedFunctionArgs } from './functions/advanced';
import { parseFormula } from './parser';
import type { RangeDependency } from './range-index';
import { createFormulaError, isArrayValue, isFormulaError, type ArrayValue, type FormulaError, type FormulaValue } from './values';
import { coerceExcelNumber, normalizeExcelPrecision } from './numeric';
import type { ExcelNumericContext } from './numeric';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';

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
    thisRow: boolean;
  }): FormulaValue | EvaluationRange | undefined;
  /** Workbook calendar used by serial/date functions. */
  readonly dateSystem?: ExcelDateSystem;
  /** Host-owned deterministic clock; missing means TODAY/NOW fail-close. */
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Workbook numeric semantics shared by inline and Worker evaluation. */
  readonly numericContext?: ExcelNumericContext;
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

type EvaluationValue = FormulaValue | EvaluationRange;

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
  if (!isEvaluationRange(value)) return value;
  const matrix = readRangeAsMatrix(value.range, context);
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
      result = evaluateFunction(node.name, node.arguments, context, trace);
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
  if (operator === '@' && isEvaluationRange(operand)) {
    const { start, end } = operand.range;
    const row = context.currentCell.row >= start.row && context.currentCell.row <= end.row ? context.currentCell.row : start.row;
    const column = context.currentCell.column >= start.column && context.currentCell.column <= end.column ? context.currentCell.column : start.column;
    const address = { sheetId: start.sheetId, row, column };
    return context.readSpillValue?.(address) ?? context.readCell(address);
  }
  if (isEvaluationRange(operand)) {
    return readRangeAsMatrix(operand.range, context).map((row) => row.map((value) => evaluateUnary(operator, value, context)));
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
  const leftValue = isEvaluationRange(left) ? readRangeAsMatrix(left.range, context) : left;
  const rightValue = isEvaluationRange(right) ? readRangeAsMatrix(right.range, context) : right;
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
): FormulaValue | EvaluationRange {
  // 需要原始 AST / 返回区间的引用类函数:在求值器内原生实现
  const native = evaluateReferenceFunction(name, argumentsList, context, trace);
  if (native !== undefined) return native;

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
    if (isEvaluationRange(value)) {
      evaluatedArgs.push(readRangeAsMatrix(value.range, context));
    } else {
      evaluatedArgs.push(value);
    }
  }

  if (fn) {
    try {
      return fn(evaluatedArgs, context);
    } catch (err) {
      return createFormulaError('#VALUE!', err instanceof Error ? err.message : 'Function evaluation error');
    }
  }

  // 上下文感知函数(SUMIFS 家族 / SUMPRODUCT / SUBTOTAL 等)
  const advanced = evaluateAdvancedFunction(name, { values: evaluatedArgs, ranges: rawRanges } as AdvancedFunctionArgs, {
    currentCell: context.currentCell,
    readMatrix: (range: RangeDependency) => readRangeAsMatrix(range, context),
    toRange: (value: EvaluationValue) => (isEvaluationRange(value) ? value.range : undefined),
  });
  if (advanced !== undefined) return advanced;

  return createFormulaError('#NAME?', `Unknown function: ${name}`);
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
): FormulaValue | EvaluationRange | undefined {
  switch (name.toUpperCase()) {
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
        values.push(isEvaluationRange(value) ? createFormulaError('#VALUE!', 'ADDRESS expects scalars') : value);
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
        if (isEvaluationRange(value)) return createFormulaError('#VALUE!', 'OFFSET offset must be scalar');
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
      return {
        kind: 'range',
        range: {
          kind: 'range',
          start: { sheetId: anchorCell.sheetId, row: startRow, column: startColumn },
          end: { sheetId: anchorCell.sheetId, row: endRow, column: endColumn },
        },
      };
    }
    case 'INDIRECT': {
      const first = args[0];
      if (!first) return createFormulaError('#REF!', 'INDIRECT expects a text reference');
      const value = evaluateNode(first, context, trace);
      if (isEvaluationRange(value)) return createFormulaError('#VALUE!', 'INDIRECT expects text');
      if (typeof value !== 'string') return createFormulaError('#REF!', 'INDIRECT text required');
      try {
        const parsed = parseFormula('=' + value);
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
