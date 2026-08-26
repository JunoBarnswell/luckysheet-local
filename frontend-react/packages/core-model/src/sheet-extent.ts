export const DEFAULT_SHEET_ROW_COUNT = 1_000 as const;
export const DEFAULT_SHEET_COLUMN_COUNT = 26 as const;

/**
 * Runtime worksheet geometry.  This is deliberately independent from the
 * OOXML exchange limits: a workbook may address a coordinate that has not
 * been materialized yet and grow its runtime extent when the owning command
 * requires it.
 */
export class SheetExtent {
  private _rowCount: number;
  private _columnCount: number;

  constructor(rowCount = DEFAULT_SHEET_ROW_COUNT, columnCount = DEFAULT_SHEET_COLUMN_COUNT) {
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
    return value;
  }

  private static assertCoordinate(value: number, name: string): void {
    if (!Number.isSafeInteger(value) || value < 0) throw new Error(`Sheet ${name} must be a non-negative safe integer`);
  }
}
