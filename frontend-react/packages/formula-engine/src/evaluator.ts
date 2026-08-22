import type { BinaryOperator, CellAddress, FormulaAst } from './ast';
import { resolveCellReference, resolveRangeReference } from './dependencies';
import { getBuiltinFunction } from './functions';
import type { RangeDependency } from './range-index';
import { createFormulaError, isFormulaError, type ArrayValue, type FormulaValue } from './values';

export interface FormulaEvaluationContext {
  readonly currentCell: CellAddress;
  readCell(address: CellAddress): FormulaValue;
  readRange(range: RangeDependency): Iterable<FormulaValue>;
  readRangeMatrix?(range: RangeDependency): ArrayValue;
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
): FormulaValue {
  const fn = getBuiltinFunction(name);
  if (!fn) return createFormulaError('#NAME?', `Unknown function: ${name}`);

  const evaluatedArgs: FormulaValue[] = [];
  for (const argument of argumentsList) {
    const value = evaluateNode(argument, context);
    if (isEvaluationRange(value)) {
      evaluatedArgs.push(readRangeAsMatrix(value.range, context));
    } else {
      evaluatedArgs.push(value);
    }
  }

  try {
    return fn(evaluatedArgs);
  } catch (err) {
    return createFormulaError('#VALUE!', err instanceof Error ? err.message : 'Function evaluation error');
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
