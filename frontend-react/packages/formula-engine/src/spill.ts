export type SpillState = 'ok' | 'blocked' | 'spill-error';

export interface SpillRange {
  sheetId: string;
  anchor: { row: number; column: number };
  range: { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number };
  values: SpillValue[][];
  state: SpillState;
}

/** Formula error codes are kept local so formula-engine remains dependency-free. */
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
  readonly message?: string;
  readonly position?: number;
}

export type ScalarValue = number | string | boolean | null;
export type ArrayValue = FormulaValue[][];
export type FormulaValue = ScalarValue | FormulaError | ArrayValue;
export type SpillValue = ScalarValue | FormulaError;

export interface SpillModel {
  anchor: { row: number; column: number };
  range: SpillRange['range'];
  values: SpillValue[][];
  state: SpillState;
  blocker?: { row: number; column: number };
}

export function createFormulaError(code: FormulaErrorCode, message: string, position?: number): FormulaError {
  return position === undefined ? { kind: 'error', code, message } : { kind: 'error', code, message, position };
}

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'error';
}

export function isArrayValue(value: unknown): value is ArrayValue {
  return Array.isArray(value) && value.length > 0 && Array.isArray(value[0]);
}

export function spillBlocked(state: SpillState): boolean {
  return state === 'blocked';
}

export const STANDARD_FORMULA_ERRORS: readonly FormulaErrorCode[] = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#CALC!',
  '#BLOCKED!',
  '#SPILL!',
  '#PARSE!',
  '#CYCLE!',
];
