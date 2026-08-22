export type ScalarValue = number | string | boolean | null;

export type FormulaErrorCode = '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#PARSE!' | '#CYCLE!';

export interface FormulaError {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
  readonly message: string;
  readonly position?: number;
}

export type FormulaValue = ScalarValue | FormulaError;

export function createFormulaError(code: FormulaErrorCode, message: string, position?: number): FormulaError {
  return position === undefined ? { kind: 'error', code, message } : { kind: 'error', code, message, position };
}

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'error';
}
