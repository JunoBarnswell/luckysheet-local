import type { BinaryOperator, CellAddress, FormulaAst } from './ast';
import { resolveCellReference, resolveRangeReference } from './dependencies';
import { getBuiltinFunction } from './functions';
import { evaluateAdvancedFunction, type AdvancedFunctionArgs } from './functions/advanced';
import { parseFormula } from './parser';
import type { RangeDependency } from './range-index';
import { createFormulaError, isFormulaError, type ArrayValue, type FormulaError, type FormulaValue } from './values';

export interface FormulaEvaluationContext {
  readonly currentCell: CellAddress;
  readCell(address: CellAddress): FormulaValue;
  readRange(range: RangeDependency): Iterable<FormulaValue>;
  readRangeMatrix?(range: RangeDependency): ArrayValue;
  /** 定义名称解析:返回 undefined 视为 #NAME? */
  resolveName?(name: string): FormulaValue | undefined;
}

interface EvaluationRange {
  readonly kind: 'range';
  readonly range: RangeDependency;
}

type EvaluationValue = FormulaValue | EvaluationRange;

export function evaluateFormula(ast: FormulaAst, context: FormulaEvaluationContext): FormulaValue {
  const result = evaluateNode(ast, context);
  if (isEvaluationRange(result)) {
    // If a top-level range is returned, return the top-left cell value or matrix
    const matrix = readRangeAsMatrix(result.range, context);
    return matrix.length === 1 && matrix[0]?.length === 1 ? matrix[0][0]! : matrix;
  }
  return result;
}

function evaluateNode(node: FormulaAst, context: FormulaEvaluationContext): EvaluationValue {
  switch (node.type) {
    case 'number-literal':
      return node.value;
    case 'string-literal':
      return node.value;
    case 'boolean-literal':
      return node.value;
    case 'cell-reference':
      return context.readCell(resolveCellReference(node.reference, context.currentCell));
    case 'range-reference':
      return { kind: 'range', range: resolveRangeReference(node, context.currentCell) };
    case 'unary-expression':
      return evaluateUnary(node.operator, evaluateNode(node.operand, context));
    case 'binary-expression':
      return evaluateBinary(
        node.operator,
        evaluateNode(node.left, context),
        evaluateNode(node.right, context),
      );
    case 'function-call':
      return evaluateFunction(node.name, node.arguments, context);
    case 'name-reference': {
      const resolved = context.resolveName?.(node.name.toUpperCase());
      return resolved === undefined ? createFormulaError('#NAME?', 'Unknown name: ' + node.name) : resolved;
    }
  }
}

function evaluateUnary(operator: '+' | '-' | '%', operand: EvaluationValue): FormulaValue {
  if (isFormulaError(operand)) return operand;
  if (isEvaluationRange(operand)) return createFormulaError('#VALUE!', 'A range cannot be used with a unary operator');
  const number = toNumber(operand);
  if (isFormulaError(number)) return number;
  if (operator === '-') return -number;
  if (operator === '%') return number / 100;
  return number;
}

function evaluateBinary(
  operator: BinaryOperator,
  left: EvaluationValue,
  right: EvaluationValue,
): FormulaValue {
  if (isFormulaError(left)) return left;
  if (isFormulaError(right)) return right;
  if (isEvaluationRange(left) || isEvaluationRange(right)) {
    return createFormulaError('#VALUE!', 'A range cannot be used in binary operator');
  }

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
      return leftNumber + rightNumber;
    case '-':
      return leftNumber - rightNumber;
    case '*':
      return leftNumber * rightNumber;
    case '/':
      return rightNumber === 0 ? createFormulaError('#DIV/0!', 'Division by zero') : leftNumber / rightNumber;
    case '^':
      return Math.pow(leftNumber, rightNumber);
  }
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
): FormulaValue | EvaluationRange {
  // 需要原始 AST / 返回区间的引用类函数:在求值器内原生实现
  const native = evaluateReferenceFunction(name, argumentsList, context);
  if (native !== undefined) return native;

  const fn = getBuiltinFunction(name);
  const evaluatedArgs: FormulaValue[] = [];
  const rawRanges: EvaluationValue[] = [];
  for (const argument of argumentsList) {
    const value = evaluateNode(argument, context);
    rawRanges.push(value);
    if (isEvaluationRange(value)) {
      evaluatedArgs.push(readRangeAsMatrix(value.range, context));
    } else {
      evaluatedArgs.push(value);
    }
  }

  if (fn) {
    try {
      return fn(evaluatedArgs);
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

/** ROW / COLUMN / ADDRESS / OFFSET / INDIRECT:需要 AST 或返回区间引用 */
function evaluateReferenceFunction(
  name: string,
  args: readonly FormulaAst[],
  context: FormulaEvaluationContext,
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
        const value = evaluateNode(argument, context);
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
        const value = evaluateNode(node, context);
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
      const value = evaluateNode(first, context);
      if (isEvaluationRange(value)) return createFormulaError('#VALUE!', 'INDIRECT expects text');
      if (typeof value !== 'string') return createFormulaError('#REF!', 'INDIRECT text required');
      try {
        const parsed = parseFormula('=' + value);
        const resolved = evaluateNode(parsed, context);
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
  if (isFormulaError(value)) return value;
  if (value === null) return 0;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const n = Number(value);
    if (!Number.isNaN(n)) return n;
  }
  return createFormulaError('#VALUE!', 'Expected a number');
}

function isEvaluationRange(value: EvaluationValue): value is EvaluationRange {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'range';
}
