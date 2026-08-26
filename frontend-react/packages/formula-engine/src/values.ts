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
  | '#PARSE!';

export interface FormulaError {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
  readonly message: string;
  readonly position?: number;
}

/**
 * A first-class reference value.  It is intentionally not a matrix: callers
 * choose scalar dereference, materialization, or dependency registration at
 * the point where the consuming function requires it.
 */
export interface ReferenceValue {
  readonly kind: 'reference';
  readonly references: readonly ReferenceSegment[];
}

export type ReferenceSegment =
  | { readonly kind: 'cell'; readonly sheetId: string; readonly row: number; readonly column: number }
  | { readonly kind: 'range'; readonly sheetId: string; readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number }
  | { readonly kind: 'whole-row'; readonly sheetId: string; readonly startRow: number; readonly endRow: number }
  | { readonly kind: 'whole-column'; readonly sheetId: string; readonly startColumn: number; readonly endColumn: number }
  | { readonly kind: 'sheet-range'; readonly startSheetId: string; readonly endSheetId: string; readonly expression: string }
  | { readonly kind: 'external'; readonly workbookId: string; readonly sheetId?: string; readonly expression: string };

export type ArrayValue = FormulaValue[][];

export type FormulaValue = ScalarValue | FormulaError | ArrayValue | ReferenceValue;

export function createFormulaError(code: FormulaErrorCode, message: string, position?: number): FormulaError {
  return position === undefined ? { kind: 'error', code, message } : { kind: 'error', code, message, position };
}

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'error';
}

export function isArrayValue(value: unknown): value is ArrayValue {
  return Array.isArray(value) && value.length > 0 && Array.isArray(value[0]);
}

export function isReferenceValue(value: unknown): value is ReferenceValue {
  return typeof value === 'object'
    && value !== null
    && 'kind' in value
    && (value as { kind: string }).kind === 'reference'
    && 'references' in value;
}
