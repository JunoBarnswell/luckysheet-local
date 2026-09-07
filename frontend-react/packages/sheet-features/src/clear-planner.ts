import type {
  CellData,
  CellHyperlink,
  CellNote,
  CommentThread,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorksheetModel,
} from '@react-sheets/core-model';
import { sheetRuleRegistry } from '@react-sheets/core-model';

/** The only clear semantics accepted by the worksheet range command. */
export type ClearFamily = 'contents' | 'formats' | 'all' | 'comments-and-notes' | 'hyperlinks';

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  family: ClearFamily;
}

export interface ClearRangeSnapshot {
  cells: Array<{ row: number; column: number; value?: CellData }>;
  notes: Array<{ row: number; column: number; note: CellNote }>;
  hyperlinks: Array<{ row: number; column: number; hyperlink: CellHyperlink }>;
  comments: CommentThread[];
  /** Complete rule snapshots are required when formats/all crop rule ranges. */
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
}

export interface ClearRangePlan {
  params: ClearRangeParams;
  range: RangeRef;
  snapshot: ClearRangeSnapshot;
}

function normalizeRange(range: RangeRef): RangeRef {
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new Error('Clear range is invalid');
  }
  return {
    ...range,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function contains(range: RangeRef, row: number, column: number): boolean {
  return range.startRow <= row && row <= range.endRow && range.startColumn <= column && column <= range.endColumn;
}

function snapshotCells(sheet: WorksheetModel, range: RangeRef): ClearRangeSnapshot['cells'] {
  const cells: ClearRangeSnapshot['cells'] = [];
  sheet.cells.forEach((cell, row, column) => {
    if (contains(range, row, column)) cells.push({ row, column, value: structuredClone(cell) });
  });
  return cells;
}

export function createClearRangePlan(sheet: WorksheetModel, input: ClearRangeParams): ClearRangePlan {
  const range = normalizeRange(input.range);
  if (range.sheetId !== sheet.id || input.sheetId !== sheet.id) throw new Error('Clear range targets another worksheet');
  const notes: ClearRangeSnapshot['notes'] = [];
  const hyperlinks: ClearRangeSnapshot['hyperlinks'] = [];
  for (const { row, column, note } of sheet.review.noteEntries()) {
    if (contains(range, row, column)) notes.push({ row, column, note });
  }
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const parts = key.split(':');
    const row = Number(parts[0]);
    const column = Number(parts[1]);
    if (Number.isInteger(row) && Number.isInteger(column) && contains(range, row, column)) hyperlinks.push({ row, column, hyperlink: structuredClone(hyperlink) });
  }
  const comments = sheet.review.threadEntries().filter((thread) => contains(range, thread.row, thread.column));
  return {
    params: { ...input, range },
    range,
    snapshot: {
      cells: snapshotCells(sheet, range),
      notes,
      hyperlinks,
      comments,
      ...(input.family === 'formats' || input.family === 'all' ? { conditionalFormats: structuredClone(sheet.conditionalFormats) } : {}),
      ...(input.family === 'formats' || input.family === 'all' ? { dataValidations: structuredClone(sheet.dataValidations) } : {}),
    },
  };
}
