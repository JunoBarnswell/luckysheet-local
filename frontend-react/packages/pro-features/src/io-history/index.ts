import type { RangeRef, WorkbookSnapshotV1 } from '@react-sheets/core-model';

export * from './xlsx';

export interface PrintLayout {
  paper: 'A4' | 'Letter' | 'Legal';
  orientation: 'portrait' | 'landscape';
  margin: { top: number; right: number; bottom: number; left: number };
  repeatRows?: RangeRef;
  scale?: number;
  fitToWidth?: boolean;
  fitToHeight?: boolean;
}

export interface PrintPage {
  page: number;
  range: RangeRef;
}

export interface RevisionEntry {
  revision: number;
  operationId: string;
  unitId: string;
  createdAt: string;
  mutationIds: string[];
}

export function paginateRange(range: RangeRef, rowsPerPage: number, columnsPerPage: number): PrintPage[] {
  if (rowsPerPage <= 0 || columnsPerPage <= 0) throw new Error('Page dimensions must be positive');
  const pages: PrintPage[] = [];
  let page = 1;
  for (let row = range.startRow; row <= range.endRow; row += rowsPerPage) {
    for (let column = range.startColumn; column <= range.endColumn; column += columnsPerPage) {
      pages.push({
        page: page++,
        range: {
          sheetId: range.sheetId,
          startRow: row,
          endRow: Math.min(range.endRow, row + rowsPerPage - 1),
          startColumn: column,
          endColumn: Math.min(range.endColumn, column + columnsPerPage - 1),
        },
      });
    }
  }
  return pages;
}

export function serializeSnapshot(snapshot: WorkbookSnapshotV1): string {
  return JSON.stringify(snapshot);
}

export function deserializeSnapshot(input: string): WorkbookSnapshotV1 {
  const snapshot = JSON.parse(input) as WorkbookSnapshotV1;
  if (snapshot.schema !== 'WorkbookSnapshotV1' || !snapshot.unitId || !Array.isArray(snapshot.sheets)) {
    throw new Error('Invalid WorkbookSnapshotV1');
  }
  return snapshot;
}

export function diffSnapshots(
  before: WorkbookSnapshotV1,
  after: WorkbookSnapshotV1,
): Array<{ sheetId: string; row: number; column: number; oldValue?: unknown; newValue?: unknown }> {
  const changes: Array<{ sheetId: string; row: number; column: number; oldValue?: unknown; newValue?: unknown }> = [];
  const beforeSheets = new Map(before.sheets.map((sheet) => [sheet.id, sheet]));

  for (const sheet of after.sheets) {
    const previous = beforeSheets.get(sheet.id);
    const rows = new Set([...Object.keys(previous?.cells ?? {}), ...Object.keys(sheet.cells)]);
    for (const row of rows) {
      const beforeColumns = previous?.cells[row] ?? {};
      const afterColumns = sheet.cells[row] ?? {};
      const columns = new Set([...Object.keys(beforeColumns), ...Object.keys(afterColumns)]);
      for (const column of columns) {
        const oldCell = beforeColumns[column];
        const newCell = afterColumns[column];
        if (JSON.stringify(oldCell) !== JSON.stringify(newCell)) {
          changes.push({
            sheetId: sheet.id,
            row: Number(row),
            column: Number(column),
            oldValue: oldCell?.value,
            newValue: newCell?.value,
          });
        }
      }
    }
  }
  return changes;
}
