import type { BinaryOperator, CellAddress, FormulaAst, FormulaReferenceNode, SpillReferenceNode } from './ast';
import { formatFormula } from './ast-format';
import { resolveCellReference, resolveRangeReference } from './dependencies';
import { getBuiltinFunction, validateFunctionArguments, SCALAR_ARRAY_FUNCTIONS } from './functions';
import { createRangeView, operandRanges } from './range-view';
import { RANGE_CONSUMER_FUNCTIONS, evaluateRangeFunction } from './functions/range-functions';
import { evaluateAggregate, flattenAggregateValue, STREAMING_AGGREGATES, type AggregateArgument } from './functions/aggregate';
import { evaluateAdvancedFunction, type AdvancedFunctionArgs } from './functions/advanced';
import { parseFormula } from './parser';
import type { RangeDependency } from './range-index';
import { createFormulaError, isArrayValue, isFormulaError, isReferenceValue, type ArrayValue, type FormulaError, type FormulaValue, type ScalarValue } from './values';
import { coerceExcelNumber, normalizeExcelPrecision } from './numeric';
import type { ExcelNumericContext } from './numeric';
import type { WorkbookCollationContext } from './collation';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';
import { createReferenceCursor, type ReferenceFormulaKind, type RowVisibilityResolver } from './reference-cursor';
import type { FormulaSheetIdentity } from './sheet-reference';
import { FormulaInputBoundary } from './input-boundary';

export interface FormulaEvaluationContext {
  readonly currentCell: CellAddress;
  readonly sheetOrder: readonly FormulaSheetIdentity[];
  readCell(address: CellAddress): FormulaValue;
  readRange(range: RangeDependency): Iterable<FormulaValue>;
  readRangeMatrix?(range: RangeDependency): ArrayValue;
  /** Occupied inputs and spill values only, for consumers that ignore blanks. */
  readSparseRange?(range: RangeDependency): Iterable<FormulaValue>;
  readSparseRangeCells?(range: RangeDependency): Iterable<{ readonly address: CellAddress; readonly value: FormulaValue }>;
  /** Resolve a dynamic-array anchor to its current spill range. */
  readSpillRange?(address: CellAddress): RangeDependency | undefined;
  /** Read a projected value from a dynamic-array spill cell. */
  readSpillValue?(address: CellAddress): FormulaValue | undefined;
  /** 定义名称解析:返回 undefined 视为 #NAME? */
  resolveName?(name: string): FormulaEvaluationValue | undefined;
  resolveFunction?(name: string): FormulaAst | undefined;
  readonly lexicalEnvironment?: ReadonlyMap<string, EvaluationValue>;
  readonly invocationDepth?: number;
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
  /** Stable workbook reference date for culture-sensitive date-entry parsing. */
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Cycle-scoped clock for volatile TODAY/NOW; does not alter date-entry interpretation. */
  readonly calculationReferenceDate?: CanonicalExcelDateParts;
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
  readonly tableBudget?: { depth: number; remaining: number; deadline: number };
  readonly evaluateWithCellOverrides?: (ast: FormulaAst, overrides: readonly FormulaCellOverride[]) => FormulaValue;
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

interface LambdaEvaluationValue {
  readonly kind: 'lambda';
  readonly parameters: readonly string[];
  readonly body: FormulaAst;
  readonly environment: ReadonlyMap<string, EvaluationValue>;
}

export type FormulaEvaluationValue = FormulaValue | EvaluationRange | FormulaEvaluationReference | LambdaEvaluationValue;
type EvaluationValue = FormulaEvaluationValue;

function isLambda(value: EvaluationValue | undefined): value is LambdaEvaluationValue {
  return typeof value === 'object' && value !== null && !Array.isArray(value) && 'kind' in value && value.kind === 'lambda';
}

export function evaluateFormula(ast: FormulaAst, context: FormulaEvaluationContext): FormulaValue {
  const boundary = new FormulaInputBoundary(context);
  const result = boundary.observe(materializeEvaluationValue(evaluateNode(ast, boundary.context), boundary.context));
  return boundary.failure ?? result;
}

/** Name and callable resolution preserve reference geometry until consumption. */
export function evaluateFormulaOperand(ast: FormulaAst, context: FormulaEvaluationContext): FormulaEvaluationValue {
  const boundary = new FormulaInputBoundary(context);
  const result = boundary.observe(evaluateNode(ast, boundary.context));
  return boundary.failure ?? result;
}

/** Evaluate a formula and capture every AST node's computed value in order. */
export function evaluateFormulaWithTrace(ast: FormulaAst, context: FormulaEvaluationContext): FormulaEvaluationTrace {
  const steps: FormulaEvaluationTraceStep[] = [];
  const boundary = new FormulaInputBoundary(context);
  const result = evaluateNode(ast, boundary.context, (node, value) => {
    const materialized = boundary.observe(materializeEvaluationValue(value, boundary.context));
    steps.push({
      node: structuredClone(node),
      expression: formatFormula(node),
      value: structuredClone(boundary.failure ?? materialized),
    });
  });
  const materialized = boundary.observe(materializeEvaluationValue(result, boundary.context));
  return { value: structuredClone(boundary.failure ?? materialized), steps };
}

type EvaluationTraceSink = (node: FormulaAst, value: EvaluationValue) => void;

function materializeEvaluationValue(value: EvaluationValue, context: FormulaEvaluationContext): FormulaValue {
  if (isLambda(value)) return createFormulaError('#CALC!', 'LAMBDA must be invoked to produce a cell value');
  if (isReferenceValue(value)) return createFormulaError('#REF!', 'Reference value requires a workbook resolver');
  if (!isEvaluationRange(value) && !isEvaluationReference(value)) return value;
  const single = isEvaluationRange(value) ? value.range : value.ranges.length === 1 ? value.ranges[0] : undefined;
  if (single && single.start.sheetId === single.end.sheetId
    && single.start.row === single.end.row && single.start.column === single.end.column) {
    return context.readSpillValue?.(single.start) ?? context.readCell(single.start);
  }
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
    case 'error-literal':
      result = createFormulaError(node.code, `Formula contains ${node.code}`, node.span.start);
      break;
    case 'cell-reference': {
      const address = resolveCellReference(node.reference, context.currentCell, context.sheetOrder);
      result = { kind: 'range', range: { kind: 'range', start: address, end: address } };
      break;
    }
    case 'range-reference':
      if (trace) { evaluateNode(node.start, context, trace); evaluateNode(node.end, context, trace); }
      result = { kind: 'range', range: resolveRangeReference(node, context.currentCell, context.sheetOrder) };
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
      result = node.callee
        ? invokeLambda(evaluateNode(node.callee, context, trace), node.arguments, context, trace)
        : evaluateFunction(node.name, node.arguments, context, trace, `${node.span.start}:${node.span.end}`);
      break;
    case 'name-reference': {
      const local = context.lexicalEnvironment?.get(node.name.toUpperCase());
      if (local !== undefined) { result = local; break; }
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
  if (isLambda(operand)) return createFormulaError('#VALUE!', 'A callable cannot be used as a number');
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
    return evaluateUnary(operator, materializeEvaluationValue(operand, context), context);
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
    ? resolveCellReference(operand.reference, context.currentCell, context.sheetOrder)
    : operand.type === 'range-reference'
      ? resolveRangeReference(operand, context.currentCell, context.sheetOrder).start
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
  if (isLambda(left) || isLambda(right)) return createFormulaError('#VALUE!', 'A callable cannot be used as a scalar');
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
): EvaluationValue {
  const id = name.toUpperCase();
  const argumentError = validateFunctionArguments(id, argumentsList.length);
  if (argumentError) return argumentError;
  if (id === 'LET') {
    if (argumentsList.length < 3 || argumentsList.length % 2 !== 1) return createFormulaError('#VALUE!', 'LET requires name/value pairs and a result');
    const environment = new Map(context.lexicalEnvironment);
    for (let index = 0; index < argumentsList.length - 1; index += 2) {
      const variable = argumentsList[index]!;
      if (!isLexicalName(variable)) return createFormulaError('#VALUE!', 'LET binding requires a valid name');
      const value = evaluateNode(argumentsList[index + 1]!, { ...context, lexicalEnvironment: environment }, trace);
      if (isFormulaError(value)) return value;
      environment.set(variable.name.toUpperCase(), value);
    }
    return evaluateNode(argumentsList[argumentsList.length - 1]!, { ...context, lexicalEnvironment: environment }, trace);
  }
  if (id === 'LAMBDA') {
    if (!argumentsList.length || argumentsList.length > 254) return createFormulaError('#VALUE!', 'LAMBDA requires a body and at most 253 parameters');
    const parameters: string[] = [];
    for (const parameter of argumentsList.slice(0, -1)) {
      if (!isLexicalName(parameter) || parameters.includes(parameter.name.toUpperCase())) return createFormulaError('#VALUE!', 'LAMBDA parameters must be valid unique names');
      parameters.push(parameter.name.toUpperCase());
    }
    return { kind: 'lambda', parameters, body: argumentsList[argumentsList.length - 1]!, environment: new Map(context.lexicalEnvironment) };
  }
  if (id === 'IF' && argumentsList.length >= 2 && argumentsList.length <= 3) {
    const condition = materializeEvaluationValue(evaluateNode(argumentsList[0]!, context, trace), context);
    if (isFormulaError(condition)) return condition;
    if (!Array.isArray(condition)) {
      if (typeof condition === 'string' && condition.toUpperCase() !== 'TRUE' && condition.toUpperCase() !== 'FALSE') return createFormulaError('#VALUE!', 'IF condition is not logical');
      const truth = typeof condition === 'string' ? condition.toUpperCase() === 'TRUE' : Boolean(condition);
      const branch = argumentsList[truth ? 1 : 2];
      return branch ? evaluateNode(branch, context, trace) : false;
    }
    const selected = condition.map((row) => row.map((value) => {
      if (isFormulaError(value)) return value;
      if (typeof value === 'string' && !['TRUE', 'FALSE'].includes(value.toUpperCase())) return createFormulaError('#VALUE!', 'IF condition is not logical');
      return typeof value === 'string' ? value.toUpperCase() === 'TRUE' : Boolean(value);
    }));
    const branches = new Map<boolean, ReturnType<typeof createRangeView>>();
    for (const row of selected) for (const value of row) if (typeof value === 'boolean' && !branches.has(value)) {
      const branch = argumentsList[value ? 1 : 2];
      branches.set(value, createRangeView(branch ? evaluateNode(branch, context, trace) : false, context));
    }
    return selected.map((row, r) => row.map((value, c) => {
      if (isFormulaError(value)) return value;
      const branch = branches.get(value)!;
      if (isFormulaError(branch)) return branch;
      if ((branch.rows !== 1 && branch.rows !== selected.length) || (branch.columns !== 1 && branch.columns !== row.length)) return createFormulaError('#VALUE!', 'IF branch shape does not match its condition');
      return branch.read(branch.rows === 1 ? 0 : r, branch.columns === 1 ? 0 : c);
    }));
  }
  if (id === 'IFERROR' || id === 'IFNA') {
    const value = evaluateNode(argumentsList[0]!, context, trace);
    const materialized = materializeEvaluationValue(value, context);
    const handles = (candidate: FormulaValue) => isFormulaError(candidate) && (id === 'IFERROR' || candidate.code === '#N/A');
    if (Array.isArray(materialized)) {
      if (!materialized.some((row) => row.some(handles))) return value;
      const fallback = createRangeView(evaluateNode(argumentsList[1]!, context, trace), context);
      return materialized.map((row, r) => row.map((candidate, c) => {
        if (!handles(candidate)) return candidate;
        if (isFormulaError(fallback)) return fallback;
        if ((fallback.rows !== 1 && fallback.rows !== materialized.length) || (fallback.columns !== 1 && fallback.columns !== row.length)) return createFormulaError('#VALUE!', 'Error fallback array has the wrong shape');
        return fallback.read(fallback.rows === 1 ? 0 : r, fallback.columns === 1 ? 0 : c);
      }));
    }
    return handles(materialized) ? evaluateNode(argumentsList[1]!, context, trace) : value;
  }
  if (id === 'CHOOSE') {
    const selector = materializeEvaluationValue(evaluateNode(argumentsList[0]!, context, trace), context);
    if (!Array.isArray(selector)) {
      const index = toNumber(selector);
      if (isFormulaError(index)) return index;
      const ordinal = Math.trunc(index);
      return ordinal < 1 || ordinal >= argumentsList.length ? createFormulaError('#VALUE!', 'CHOOSE index is out of bounds') : evaluateNode(argumentsList[ordinal]!, context, trace);
    }
  }
  const localFunction = context.lexicalEnvironment?.get(id);
  if (localFunction !== undefined) return invokeLambda(localFunction, argumentsList, context, trace);
  if (!getBuiltinFunction(id) && !STREAMING_AGGREGATES.has(id)) {
    const defined = context.resolveFunction?.(id);
    if (defined) return invokeLambda(evaluateNode(defined, { ...context, lexicalEnvironment: undefined }, trace), argumentsList, context, trace);
  }
  // 需要原始 AST / 返回区间的引用类函数:在求值器内原生实现
  const native = evaluateReferenceFunction(name, argumentsList, context, trace);
  if (native !== undefined) return native;

  const normalizedName = name.toUpperCase();
  if (RANGE_CONSUMER_FUNCTIONS.has(normalizedName)) {
    return evaluateRangeFunction(normalizedName, argumentsList.map((argument) => evaluateNode(argument, context, trace)), context);
  }
  if (STREAMING_AGGREGATES.has(normalizedName)) {
    function* argumentsForAggregate(): Iterable<AggregateArgument> {
      for (const argument of argumentsList) {
        const value = evaluateNode(argument, context, trace);
        if (isEvaluationRange(value) || isEvaluationReference(value)) {
          const ranges = isEvaluationRange(value) ? [value.range] : value.ranges;
          function* values(): Iterable<FormulaValue> {
            for (const range of ranges) yield* (context.readSparseRange?.(range) ?? context.readRange(range));
          }
          yield { values: values(), reference: true };
        } else {
          yield { values: flattenAggregateValue(materializeEvaluationValue(value, context)), reference: Array.isArray(value) };
        }
      }
    }
    return evaluateAggregate(normalizedName, argumentsForAggregate());
  }

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
    // Visibility aggregates consume geometry through the canonical cursor.
    evaluatedArgs.push((id === 'SUBTOTAL' && index >= 1) || (id === 'AGGREGATE' && index >= 2 && !(argumentsList.length === 4 && index === 3)) ? null : materializeEvaluationValue(value, context));
  }

  if (fn) {
    try {
      const invoke = (args: FormulaValue[]) => {
        if (!['IF', 'IFS', 'SWITCH', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'ISERROR', 'ISERR', 'ISNA'].includes(id)) {
          const error = args.find(isFormulaError);
          if (error) return error;
        }
        const value = fn(args, volatileOccurrence === undefined ? context : { ...context, volatileOccurrence });
        return typeof value === 'number' && !Number.isFinite(value) ? createFormulaError('#NUM!', 'Function result is not finite') : value;
      };
      if (SCALAR_ARRAY_FUNCTIONS.has(id) && evaluatedArgs.some(Array.isArray)) {
        const views = evaluatedArgs.map((arg) => createRangeView(arg, context));
        const error = views.find(isFormulaError);
        if (error) return error;
        const matrices = views as import('./range-view').FormulaRangeView[];
        const rows = matrices.reduce((maximum, view) => Math.max(maximum, view.rows), 1);
        const columns = matrices.reduce((maximum, view) => Math.max(maximum, view.columns), 1);
        if (matrices.some((view) => (view.rows !== 1 && view.rows !== rows) || (view.columns !== 1 && view.columns !== columns))) return createFormulaError('#VALUE!', 'Scalar function arrays cannot be broadcast to the same shape');
        return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => invoke(matrices.map((view) => view.read(view.rows === 1 ? 0 : row, view.columns === 1 ? 0 : column)))));
      }
      return invoke(evaluatedArgs);
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

  return createFormulaError('#NAME?', `Unknown function: ${name}`);
}

function isLexicalName(node: FormulaAst): node is Extract<FormulaAst, { type: 'name-reference' }> {
  return node.type === 'name-reference' && /^[\p{L}_][\p{L}\p{N}_]*$/u.test(node.name) && !['R', 'C'].includes(node.name.toUpperCase());
}

function invokeLambda(value: EvaluationValue, args: readonly FormulaAst[], context: FormulaEvaluationContext, trace?: EvaluationTraceSink): EvaluationValue {
  if (isFormulaError(value)) return value;
  if (!isLambda(value)) return createFormulaError('#VALUE!', 'Expression is not callable');
  if (value.parameters.length !== args.length) return createFormulaError('#VALUE!', 'LAMBDA argument count differs from its parameters');
  const depth = (context.invocationDepth ?? 0) + 1;
  if (depth > 128) return createFormulaError('#NUM!', 'LAMBDA recursion limit exceeded');
  const environment = new Map(value.environment);
  args.forEach((argument, index) => environment.set(value.parameters[index]!, evaluateNode(argument, context, trace)));
  return evaluateNode(value.body, { ...context, lexicalEnvironment: environment, invocationDepth: depth }, trace);
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
    case 'ROW':
    case 'COLUMN': {
      const horizontal = name.toUpperCase() === 'COLUMN';
      if (!args.length) return (horizontal ? context.currentCell.column : context.currentCell.row) + 1;
      const operand = evaluateNode(args[0]!, context, trace);
      const ranges = operandRanges(operand);
      if (!ranges || ranges.length !== 1) return isFormulaError(operand) ? operand : createFormulaError('#VALUE!', `${name} expects one reference`);
      const range = ranges[0]!;
      const count = horizontal ? range.end.column - range.start.column + 1 : range.end.row - range.start.row + 1;
      const start = horizontal ? range.start.column : range.start.row;
      if (count === 1) return start + 1;
      return horizontal ? [Array.from({ length: count }, (_, index) => start + index + 1)] : Array.from({ length: count }, (_, index) => [start + index + 1]);
    }
    case 'ADDRESS': {
      const values = args.map((argument) => materializeEvaluationValue(evaluateNode(argument, context, trace), context));
      const row = toNumber(values[0]!);
      const column = toNumber(values[1]!);
      const absMode = toNumber(values[2] ?? 1);
      if (isFormulaError(row) || isFormulaError(column) || isFormulaError(absMode)) return [row, column, absMode].find(isFormulaError)!;
      if (row < 1 || row > 1_048_576 || column < 1 || column > 16_384 || ![1, 2, 3, 4].includes(Math.trunc(absMode))) return createFormulaError('#VALUE!', 'ADDRESS coordinates or absolute mode are out of bounds');
      const r = Math.trunc(row);
      const c = Math.trunc(column);
      const mode = Math.trunc(absMode);
      const a1 = values[3] === undefined || values[3] === true || values[3] === 1 || String(values[3]).toUpperCase() === 'TRUE';
      let reference: string;
      if (a1) {
        let letters = '';
        for (let remaining = c; remaining > 0; remaining = Math.floor((remaining - 1) / 26)) letters = String.fromCharCode(65 + (remaining - 1) % 26) + letters;
        reference = `${mode === 1 || mode === 3 ? '$' : ''}${letters}${mode === 1 || mode === 2 ? '$' : ''}${r}`;
      } else reference = `R${mode === 1 || mode === 2 ? r : `[${r}]`}C${mode === 1 || mode === 3 ? c : `[${c}]`}`;
      if (values[4] !== undefined) reference = `'${String(values[4]).replace(/'/g, "''")}'!${reference}`;
      return reference;
    }
    case 'OFFSET': {
      const operand = evaluateNode(args[0]!, context, trace);
      const ranges = operandRanges(operand);
      if (!ranges || ranges.length !== 1) return isFormulaError(operand) ? operand : createFormulaError('#VALUE!', 'OFFSET expects one rectangular reference');
      const base = ranges[0]!;
      const scalar = (index: number, fallback: number) => args[index] ? toNumber(materializeEvaluationValue(evaluateNode(args[index]!, context, trace), context)) : fallback;
      const numbers = [scalar(1, 0), scalar(2, 0), scalar(3, base.end.row - base.start.row + 1), scalar(4, base.end.column - base.start.column + 1)];
      const error = numbers.find(isFormulaError);
      if (error) return error;
      const [rows, columns, height, width] = (numbers as number[]).map(Math.trunc) as [number, number, number, number];
      const start = { sheetId: base.start.sheetId, row: base.start.row + rows, column: base.start.column + columns };
      const end = { sheetId: start.sheetId, row: start.row + height - 1, column: start.column + width - 1 };
      if (start.row < 0 || start.column < 0 || height < 1 || width < 1 || end.row >= 1_048_576 || end.column >= 16_384) return createFormulaError('#REF!', 'OFFSET is out of bounds');
      return { kind: 'range', range: { kind: 'range', start, end } };
    }
    case 'INDIRECT': {
      const value = materializeEvaluationValue(evaluateNode(args[0]!, context, trace), context);
      if (isFormulaError(value)) return value;
      if (typeof value !== 'string') return createFormulaError('#REF!', 'INDIRECT text is required');
      const mode = args[1] ? materializeEvaluationValue(evaluateNode(args[1], context, trace), context) : true;
      if (isFormulaError(mode)) return mode;
      try {
        let source = value;
        if (mode === false || mode === 0 || String(mode).toUpperCase() === 'FALSE') {
          const match = /^(.*!)?(R(?:\[[+-]?\d+\]|\d+)?C(?:\[[+-]?\d+\]|\d+)?)(?::(R(?:\[[+-]?\d+\]|\d+)?C(?:\[[+-]?\d+\]|\d+)?))?$/i.exec(value);
          if (!match) return createFormulaError('#REF!', 'INDIRECT expects an R1C1 reference');
          const convert = (cell: string) => {
            const parts = /^R(\[[+-]?\d+\]|\d+)?C(\[[+-]?\d+\]|\d+)?$/i.exec(cell)!;
            const coordinate = (part: string | undefined, base: number) => part === undefined ? base : part.startsWith('[') ? base + Number(part.slice(1, -1)) : Number(part) - 1;
            const row = coordinate(parts[1], context.currentCell.row);
            const column = coordinate(parts[2], context.currentCell.column);
            if (row < 0 || row >= 1_048_576 || column < 0 || column >= 16_384) throw new Error('INDIRECT is out of bounds');
            let letters = '';
            for (let remaining = column + 1; remaining > 0; remaining = Math.floor((remaining - 1) / 26)) letters = String.fromCharCode(65 + (remaining - 1) % 26) + letters;
            return `${letters}${row + 1}`;
          };
          source = `${match[1] ?? ''}${convert(match[2]!)}${match[3] ? `:${convert(match[3])}` : ''}`;
        }
        const parsed = parseFormula('=' + source);
        if (!['cell-reference', 'range-reference', 'whole-column-reference', 'whole-row-reference', 'reference-union', 'reference-intersection', 'sheet-range-reference', 'external-reference', 'table-reference', 'spill-reference', 'name-reference'].includes(parsed.type)) return createFormulaError('#REF!', 'INDIRECT text must be a reference');
        const resolved = evaluateNode(parsed, context, trace);
        return operandRanges(resolved) ? resolved as EvaluationRange | FormulaEvaluationReference : isFormulaError(resolved) ? resolved : createFormulaError('#REF!', 'INDIRECT name must resolve to a reference');
      } catch {
        return createFormulaError('#REF!', 'INDIRECT cannot resolve: ' + value);
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

  const budget = context.tableBudget ?? { depth: 0, remaining: 100_000, deadline: Date.now() + 1000 };
  if (budget.depth >= 8 || budget.remaining <= 0 || Date.now() > budget.deadline) return createFormulaError('#CALC!', 'SJS.TABLE exceeds the shared calculation budget');
  budget.depth += 1;
  try {
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
    if (--budget.remaining < 0 || Date.now() > budget.deadline) return createFormulaError('#CALC!', 'SJS.TABLE exceeds the shared calculation budget');
    const overrides = inputPairs.map((pair) => ({ address: pair.address, value: scalarForTable(pair.values[row]!) }));
    const value = context.evaluateWithCellOverrides(args[0]!, overrides);
    trace?.(args[0]!, value as EvaluationValue);
    if (isFormulaError(value)) return value;
    if (isArrayValue(value) || isReferenceValue(value)) return createFormulaError('#CALC!', 'SJS.TABLE resultReference must resolve to a scalar');
    result.push([value]);
  }
  return result;
  } finally { budget.depth -= 1; }
}

function referenceCellAddress(node: FormulaAst, context: FormulaEvaluationContext): CellAddress | undefined {
  if (node.type !== 'cell-reference') return undefined;
  return resolveCellReference(node.reference, context.currentCell, context.sheetOrder);
}

function referenceValues(node: FormulaAst, context: FormulaEvaluationContext, trace?: EvaluationTraceSink): FormulaValue[] | FormulaError {
  const value = evaluateNode(node, context, trace);
  if (isFormulaError(value)) return value;
  const rangeWithinBudget = (range: Parameters<FormulaEvaluationContext['readRange']>[0]) => (Math.abs(range.end.row - range.start.row) + 1) * (Math.abs(range.end.column - range.start.column) + 1) <= (context.tableBudget?.remaining ?? 100_000);
  if (isEvaluationRange(value)) return rangeWithinBudget(value.range) ? [...context.readRange(value.range)] : createFormulaError('#CALC!', 'SJS.TABLE input range exceeds the calculation budget');
  if (isEvaluationReference(value)) return value.ranges.every(rangeWithinBudget) ? value.ranges.flatMap((range) => [...context.readRange(range)]) : createFormulaError('#CALC!', 'SJS.TABLE input ranges exceed the calculation budget');
  if (isArrayValue(value)) return value.flat();
  if (isReferenceValue(value)) return createFormulaError('#VALUE!', 'SJS.TABLE inputs must be cell or range references');
  return [materializeEvaluationValue(value, context)];
}

function scalarForTable(value: FormulaValue): ScalarValue {
  if (isFormulaError(value)) return null;
  if (isArrayValue(value) || isReferenceValue(value)) return null;
  return value;
}

function readRangeAsMatrix(range: RangeDependency, context: FormulaEvaluationContext): ArrayValue {
  if ((range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1) > 100000) return [[createFormulaError("#CALC!", "Range matrix exceeds 100000 cells")]];
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
  if (ranges.reduce((sum, range) => sum + (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1), 0) > 100000) return [[createFormulaError("#CALC!", "Reference matrix exceeds 100000 cells")]];
  const matrix: ArrayValue = [];
  for (const range of ranges) for (const row of readRangeAsMatrix(range, context)) matrix.push(row);
  return matrix;
}
