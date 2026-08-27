export type CellEntryErrorCode =
  | 'CELL_ENTRY_INVALID_INPUT'
  | 'CELL_ENTRY_SPILL_CHILD'
  | 'CELL_ENTRY_VALIDATION_BLOCKED'
  | 'CELL_ENTRY_CONFIRMATION_REQUIRED'
  | 'CELL_ENTRY_EDITOR_MISMATCH'
  | 'CELL_ENTRY_RICH_TEXT_INVALID';

export interface CellEntryFailure {
  code: CellEntryErrorCode;
  message: string;
  sheetId: string;
  row: number;
  column: number;
  recovery: string;
  alertStyle?: 'stop' | 'warning' | 'information';
  title?: string;
}
export class CellEntryError extends Error implements CellEntryFailure {
  readonly code: CellEntryErrorCode;
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly recovery: string;
  readonly alertStyle?: 'stop' | 'warning' | 'information';
  readonly title?: string;

  constructor(failure: CellEntryFailure, options?: ErrorOptions) {
    super(failure.message, options);
    this.name = 'CellEntryError';
    this.code = failure.code;
    this.sheetId = failure.sheetId;
    this.row = failure.row;
    this.column = failure.column;
    this.recovery = failure.recovery;
    this.alertStyle = failure.alertStyle;
    this.title = failure.title;
  }
}

export function isCellEntryError(value: unknown): value is CellEntryError {
  return value instanceof CellEntryError;
}
