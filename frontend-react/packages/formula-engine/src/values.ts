export type ScalarValue = number | string | boolean | null;

export type FormulaErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#CALC!'
  | '#BLOCKED!'
  | '#SPILL!'
  | '#PARSE!'
  | '#CYCLE!';

export interface FormulaError {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
  readonly message: string;
  readonly position?: number;
}

export type ArrayValue = FormulaValue[][];

export type FormulaValue = ScalarValue | FormulaError | ArrayValue;

export function createFormulaError(code: FormulaErrorCode, message: string, position?: number): FormulaError {
  return position === undefined ? { kind: 'error', code, message } : { kind: 'error', code, message, position };
}

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'error';
}

export function isArrayValue(value: unknown): value is ArrayValue {
  return Array.isArray(value) && value.length > 0 && Array.isArray(value[0]);
}
