import type { CellData, RangeRef, TableFieldType, WorksheetModel } from '@react-sheets/core-model';
import { columnLabel } from './address';

export { columnLabel } from './address';

export function usedRangeOfSheet(sheet: WorksheetModel): RangeRef {
  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxColumn = 0;
  sheet.cells.forEach((_cell, row, column) => {
    minRow = Math.min(minRow, row);
    minColumn = Math.min(minColumn, column);
    maxRow = Math.max(maxRow, row);
    maxColumn = Math.max(maxColumn, column);
  });
  return {
    sheetId: sheet.id,
    startRow: Number.isFinite(minRow) ? minRow : 0,
    endRow: Number.isFinite(minRow) ? maxRow : 0,
    startColumn: Number.isFinite(minColumn) ? minColumn : 0,
    endColumn: Number.isFinite(minColumn) ? maxColumn : 0,
  };
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
