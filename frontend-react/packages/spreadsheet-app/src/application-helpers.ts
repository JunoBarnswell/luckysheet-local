import type { CellData, RangeRef, TableFieldType, WorksheetModel } from '@react-sheets/core-model';
import { columnLabel } from './address';

export { columnLabel } from './address';

export function usedRangeOfSheet(sheet: WorksheetModel): RangeRef {
  return sheet.usedRange;
}

/**
 * Excel-style current region. A table/data region always wins; otherwise the
 * region grows over adjacent nonblank rows and columns around the active cell.
 */
export function currentRegionOfSheet(
  sheet: WorksheetModel,
  row: number,
  column: number,
  isOccupied: (row: number, column: number) => boolean,
): RangeRef {
  const table = sheet.sheetTables.find((entry) =>
    row >= entry.range.startRow && row <= entry.range.endRow
    && column >= entry.range.startColumn && column <= entry.range.endColumn);
  if (table) return structuredClone(table.range);
  const dataRegion = sheet.dataRegions.find((entry) =>
    row >= entry.range.startRow && row <= entry.range.endRow
    && column >= entry.range.startColumn && column <= entry.range.endColumn);
  if (dataRegion) return structuredClone(dataRegion.range);
  if (!isOccupied(row, column)) {
    return { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column };
  }

  let startRow = row;
  let endRow = row;
  let startColumn = column;
  let endColumn = column;
  const rowHasData = (candidate: number): boolean => {
    for (let current = startColumn; current <= endColumn; current += 1) {
      if (isOccupied(candidate, current)) return true;
    }
    return false;
  };
  const columnHasData = (candidate: number): boolean => {
    for (let current = startRow; current <= endRow; current += 1) {
      if (isOccupied(current, candidate)) return true;
    }
    return false;
  };
  let grew = true;
  while (grew) {
    grew = false;
    while (startRow > 0 && rowHasData(startRow - 1)) { startRow -= 1; grew = true; }
    while (endRow + 1 < sheet.rowCount && rowHasData(endRow + 1)) { endRow += 1; grew = true; }
    while (startColumn > 0 && columnHasData(startColumn - 1)) { startColumn -= 1; grew = true; }
    while (endColumn + 1 < sheet.columnCount && columnHasData(endColumn + 1)) { endColumn += 1; grew = true; }
  }
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

export function inferTableFieldType(values: CellData['value'][]): TableFieldType {
  const present = values.filter((value) => value != null && value !== '');
  if (present.length === 0) return 'mixed';
  if (present.every((value) => typeof value === 'number')) return 'number';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value))) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

export function nextId(prefix: string): string {
  return prefix + '-' + Math.random().toString(36).slice(2, 8);
}
