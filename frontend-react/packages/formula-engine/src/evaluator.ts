import type { CellAddress, FormulaAst } from './ast';
import { resolveCellReference, resolveRangeReference } from './dependencies';
import type { RangeDependency } from './range-index';
import { createFormulaError, isFormulaError, type FormulaValue } from './values';

export interface FormulaEvaluationContext {
  readonly currentCell: CellAddress;
  readCell(address: CellAddress): FormulaValue;
  readRange(range: RangeDependency): Iterable<FormulaValue>;
}

interface EvaluationRange {
  readonly kind: 'range';
  readonly range: RangeDependency;
}

type EvaluationValue = FormulaValue | EvaluationRange;

export function evaluateFormula(ast: FormulaAst, context: FormulaEvaluationContext): FormulaValue {
  const result = evaluateNode(ast, context);
  return isEvaluationRange(result)
    ? createFormulaError('#VALUE!', 'A range can only be used as a function argument')
    : result;
}

function evaluateNode(node: FormulaAst, context: FormulaEvaluationContext): EvaluationValue {
  switch (node.type) {
    case 'number-literal':
      return node.value;
    case 'string-literal':
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

function evaluateUnary(operator: '+' | '-', operand: EvaluationValue): FormulaValue {
  if (isFormulaError(operand)) return operand;
  if (isEvaluationRange(operand)) return createFormulaError('#VALUE!', 'A range cannot be used with a unary operator');
  const number = toNumber(operand);
  if (isFormulaError(number)) return number;
  return operator === '-' ? -number : number;
}

function evaluateBinary(
  operator: '+' | '-' | '*' | '/',
  left: EvaluationValue,
  right: EvaluationValue,
): FormulaValue {
  if (isFormulaError(left)) return left;
  if (isFormulaError(right)) return right;
  if (isEvaluationRange(left) || isEvaluationRange(right)) {
    return createFormulaError('#VALUE!', 'A range cannot be used in arithmetic');
  }

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
  }
}

function evaluateFunction(
  name: string,
  argumentsList: readonly FormulaAst[],
  context: FormulaEvaluationContext,
): FormulaValue {
  if (name.toUpperCase() !== 'SUM') return createFormulaError('#NAME?', `Unknown function: ${name}`);

  let sum = 0;
  for (const argument of argumentsList) {
    const value = evaluateNode(argument, context);
    if (isFormulaError(value)) return value;
    if (isEvaluationRange(value)) {
      for (const rangeValue of context.readRange(value.range)) {
        if (isFormulaError(rangeValue)) return rangeValue;
        if (typeof rangeValue === 'number') sum += rangeValue;
      }
      continue;
    }
    if (typeof value === 'number') sum += value;
  }
  return sum;
}

function toNumber(value: FormulaValue): number | ReturnType<typeof createFormulaError> {
  if (isFormulaError(value)) return value;
  if (value === null) return 0;
  return typeof value === 'number' ? value : createFormulaError('#VALUE!', 'Expected a number');
}

function isEvaluationRange(value: EvaluationValue): value is EvaluationRange {
  return typeof value === 'object' && value !== null && value.kind === 'range';
}
