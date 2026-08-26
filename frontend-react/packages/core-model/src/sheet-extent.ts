export const DEFAULT_SHEET_ROW_COUNT = 1_000 as const;
export const DEFAULT_SHEET_COLUMN_COUNT = 26 as const;
export const MAX_SHEET_ROW_COUNT = 1_048_576 as const;
export const MAX_SHEET_COLUMN_COUNT = 16_384 as const;
export const SHEET_ROW_GROWTH_CHUNK = 1_000 as const;
export const SHEET_COLUMN_GROWTH_CHUNK = 26 as const;

/**
 * Runtime worksheet geometry. Cells remain sparse while the addressable
 * extent grows on demand up to the native Excel/OOXML worksheet boundary.
 */
export class SheetExtent {
  private _rowCount: number;
  private _columnCount: number;

  constructor(rowCount: number = DEFAULT_SHEET_ROW_COUNT, columnCount: number = DEFAULT_SHEET_COLUMN_COUNT) {
    this._rowCount = SheetExtent.normalizeCount(rowCount, 'rowCount');
    this._columnCount = SheetExtent.normalizeCount(columnCount, 'columnCount');
  }

  get rowCount(): number { return this._rowCount; }
  get columnCount(): number { return this._columnCount; }

  set rowCount(value: number) {
    this._rowCount = SheetExtent.normalizeCount(value, 'rowCount');
  }

  set columnCount(value: number) {
    this._columnCount = SheetExtent.normalizeCount(value, 'columnCount');
  }

  ensureCell(row: number, column: number): void {
    SheetExtent.assertCoordinate(row, 'row');
    SheetExtent.assertCoordinate(column, 'column');
    if (row + 1 > this._rowCount) this._rowCount = row + 1;
    if (column + 1 > this._columnCount) this._columnCount = column + 1;
  }

  ensureRange(startRow: number, endRow: number, startColumn: number, endColumn: number): void {
    SheetExtent.assertCoordinate(startRow, 'startRow');
    SheetExtent.assertCoordinate(endRow, 'endRow');
    SheetExtent.assertCoordinate(startColumn, 'startColumn');
    SheetExtent.assertCoordinate(endColumn, 'endColumn');
    if (endRow < startRow || endColumn < startColumn) throw new Error('Sheet extent range must be ordered');
    this.ensureCell(endRow, endColumn);
  }

  contains(row: number, column: number): boolean {
    return Number.isSafeInteger(row) && row >= 0 && row < this._rowCount
      && Number.isSafeInteger(column) && column >= 0 && column < this._columnCount;
  }

  private static normalizeCount(value: number, name: string): number {
    if (!Number.isSafeInteger(value) || value <= 0) throw new Error(`Sheet ${name} must be a positive safe integer`);
    const maximum = name === 'rowCount' ? MAX_SHEET_ROW_COUNT : MAX_SHEET_COLUMN_COUNT;
    if (value > maximum) throw new Error(`UNSUPPORTED_FEATURE: Sheet ${name} exceeds the Excel-compatible limit of ${maximum}`);
    return value;
  }

  private static assertCoordinate(value: number, name: string): void {
    if (!Number.isSafeInteger(value) || value < 0) throw new Error(`Sheet ${name} must be a non-negative safe integer`);
    const rowCoordinate = name.toLowerCase().includes('row');
    const maximum = (rowCoordinate ? MAX_SHEET_ROW_COUNT : MAX_SHEET_COLUMN_COUNT) - 1;
    if (value > maximum) throw new Error(`UNSUPPORTED_FEATURE: Sheet ${name} exceeds the Excel-compatible index limit of ${maximum}`);
  }
}
