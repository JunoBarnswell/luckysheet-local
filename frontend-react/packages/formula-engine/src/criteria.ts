import { coerceExcelNumber } from './numeric';
import { isFormulaError, type FormulaValue } from './values';

export type CriteriaOperator = 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le';

export interface CriteriaExpression {
  readonly operator: CriteriaOperator;
  readonly operand: FormulaValue;
  readonly wildcard?: string;
}

export interface CriteriaRange {
  readonly values: FormulaValue[][];
  readonly rows: number;
  readonly columns: number;
}

export function parseCriteria(value: FormulaValue): CriteriaExpression {
  if (isFormulaError(value)) return { operator: 'eq', operand: value };
  if (value === null) return { operator: 'eq', operand: '' };
  if (typeof value !== 'string') return { operator: 'eq', operand: value };
  const match = /^(<>|>=|<=|=|>|<)(.*)$/.exec(value.trim());
  const operator = match?.[1] ?? '=';
  const operandText = (match?.[2] ?? value).trim();
  const parsedOperator: CriteriaOperator = operator === '<>' ? 'ne'
    : operator === '>' ? 'gt'
      : operator === '>=' ? 'ge'
        : operator === '<' ? 'lt'
          : operator === '<=' ? 'le' : 'eq';
  const wildcard = hasWildcard(operandText) ? operandText : undefined;
  if (wildcard !== undefined) return { operator: parsedOperator, operand: operandText, wildcard };
  const numeric = coerceExcelNumber(operandText);
  return { operator: parsedOperator, operand: isFormulaError(numeric) ? operandText : numeric };
}

export function matchesCriteria(value: FormulaValue, expression: CriteriaExpression): boolean {
  if (isFormulaError(value) || isFormulaError(expression.operand)) return false;
  if (expression.wildcard !== undefined) {
    const matched = wildcardMatches(String(value ?? ''), expression.wildcard);
    return expression.operator === 'ne' ? !matched : expression.operator === 'eq' && matched;
  }
  const numericValue = coerceExcelNumber(value);
  const numericOperand = coerceExcelNumber(expression.operand);
  if (!isFormulaError(numericValue) && !isFormulaError(numericOperand)) {
    return compareNumbers(expression.operator, numericValue, numericOperand);
  }
  const left = String(value ?? '').toLocaleLowerCase('en-US');
  const right = String(expression.operand ?? '').toLocaleLowerCase('en-US');
  return compareStrings(expression.operator, left, right);
}

export function toCriteriaRange(value: FormulaValue): CriteriaRange {
  if (!Array.isArray(value)) return { values: [[value]], rows: 1, columns: 1 };
  const values = (value as unknown[]).map((row) => Array.isArray(row) ? row.map((cell) => (cell ?? null) as FormulaValue) : [(row ?? null) as FormulaValue]);
  const columns = values.length === 0 ? 0 : values[0]!.length;
  if (values.some((row) => row.length !== columns)) return { values, rows: values.length, columns: -1 };
  return { values, rows: values.length, columns };
}

export function sameCriteriaShape(left: CriteriaRange, right: CriteriaRange): boolean {
  return left.columns >= 0 && right.columns >= 0 && left.rows === right.rows && left.columns === right.columns;
}

/** SUMIF/AVERAGEIF use the criteria range dimensions as a top-left projection. */
export function projectCriteriaRange(source: CriteriaRange, rows: number, columns: number): CriteriaRange {
  const values = Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => source.values[row]?.[column] ?? null));
  return { values, rows, columns };
}

function hasWildcard(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    if (value[index] !== '~' && (value[index] === '*' || value[index] === '?')) return true;
    if (value[index] === '~') index += 1;
  }
  return false;
}

function wildcardMatches(value: string, pattern: string): boolean {
  let regex = '^';
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index]!;
    if (character === '~' && index + 1 < pattern.length) {
      regex += escapeRegex(pattern[++index]!);
    } else if (character === '*') regex += '.*';
    else if (character === '?') regex += '.';
    else regex += escapeRegex(character);
  }
  regex += '$';
  return new RegExp(regex, 'iu').test(value);
}

function compareNumbers(operator: CriteriaOperator, left: number, right: number): boolean {
  switch (operator) {
    case 'ne': return left !== right;
    case 'gt': return left > right;
    case 'ge': return left >= right;
    case 'lt': return left < right;
    case 'le': return left <= right;
    default: return left === right;
  }
}

function compareStrings(operator: CriteriaOperator, left: string, right: string): boolean {
  switch (operator) {
    case 'ne': return left !== right;
    case 'gt': return left > right;
    case 'ge': return left >= right;
    case 'lt': return left < right;
    case 'le': return left <= right;
    default: return left === right;
  }
}

function escapeRegex(value: string): string {
  return value.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&');
}
