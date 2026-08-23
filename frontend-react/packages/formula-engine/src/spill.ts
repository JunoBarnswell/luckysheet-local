import type { FormulaValue as CoreFormulaValue, FormulaErrorCode as CoreFormulaErrorCode, SpillRange, SpillState } from '@react-sheets/core-model';

export type { SpillRange, SpillState };

/** 与 core-model FormulaErrorCode 对齐的引擎错误码 */
export type FormulaErrorCode =
  | CoreFormulaErrorCode
  | '#PARSE!'
  | '#CYCLE!';

export interface FormulaError {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
  readonly message: string;
  readonly position?: number;
}

export type ScalarValue = number | string | boolean | null;
export type ArrayValue = FormulaValue[][];
export type FormulaValue = ScalarValue | FormulaError | ArrayValue;

export interface SpillModel {
  anchor: { row: number; column: number };
  range: SpillRange['range'];
  values: CoreFormulaValue[][];
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
