import {
  MAX_SHEET_COLUMN_COUNT,
  MAX_SHEET_ROW_COUNT,
  SHEET_COLUMN_GROWTH_CHUNK,
  SHEET_ROW_GROWTH_CHUNK,
} from '@react-sheets/core-model';

export interface SheetExtentState {
  sheetId: string;
  rowCount: number;
  columnCount: number;
}

export interface SheetExtentGrowthAxes {
  rows?: boolean;
  columns?: boolean;
}

/** Plans one canonical extent mutation and suppresses duplicate in-flight requests. */
export function planSheetExtentGrowth(
  current: SheetExtentState,
  pending: SheetExtentState,
  axes: SheetExtentGrowthAxes,
): SheetExtentState | null {
  const ownedPending = pending.sheetId === current.sheetId ? pending : current;
  const rowCount = axes.rows && ownedPending.rowCount <= current.rowCount
    ? Math.min(MAX_SHEET_ROW_COUNT, current.rowCount + SHEET_ROW_GROWTH_CHUNK)
    : Math.max(current.rowCount, ownedPending.rowCount);
  const columnCount = axes.columns && ownedPending.columnCount <= current.columnCount
    ? Math.min(MAX_SHEET_COLUMN_COUNT, current.columnCount + SHEET_COLUMN_GROWTH_CHUNK)
    : Math.max(current.columnCount, ownedPending.columnCount);
  if (rowCount === ownedPending.rowCount && columnCount === ownedPending.columnCount) return null;
  return { sheetId: current.sheetId, rowCount, columnCount };
}

export function resolveAutoScrollExtentGrowth(input: {
  right: boolean;
  bottom: boolean;
  viewport: { width: number; height: number; scrollX: number; scrollY: number };
  content: { width: number; height: number };
  defaultRowHeight: number;
  defaultColumnWidth: number;
}): SheetExtentGrowthAxes {
  return {
    rows: input.bottom
      && input.content.height - input.viewport.scrollY - input.viewport.height
        <= Math.max(input.viewport.height, input.defaultRowHeight * 50),
    columns: input.right
      && input.content.width - input.viewport.scrollX - input.viewport.width
        <= Math.max(input.viewport.width, input.defaultColumnWidth * 8),
  };
}
