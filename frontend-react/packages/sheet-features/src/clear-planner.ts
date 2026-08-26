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
  for (const [key, note] of sheet.notes) {
    const parts = key.split(':');
    const row = Number(parts[0]);
    const column = Number(parts[1]);
    if (Number.isInteger(row) && Number.isInteger(column) && contains(range, row, column)) notes.push({ row, column, note: structuredClone(note) });
  }
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const parts = key.split(':');
    const row = Number(parts[0]);
    const column = Number(parts[1]);
    if (Number.isInteger(row) && Number.isInteger(column) && contains(range, row, column)) hyperlinks.push({ row, column, hyperlink: structuredClone(hyperlink) });
  }
  const comments = sheet.commentThreads
    .filter((thread) => contains(range, thread.row, thread.column))
    .map((thread) => structuredClone(thread));
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

function clearCellContents(cell: CellData): CellData {
  const next = { ...cell, value: null };
  delete next.formula;
  delete next.displayValue;
  return next;
}

function clearCellFormats(cell: CellData): CellData {
  const next = { ...cell };
  delete next.style;
  delete next.styleId;
  delete next.numberFormat;
  delete next.displayValue;
  return next;
}

export function applyClearRangePlan(sheet: WorksheetModel, plan: ClearRangePlan): void {
  const { range, params } = plan;
  const cells: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (contains(range, row, column)) cells.push({ row, column, cell });
  });
  if (params.family === 'comments-and-notes') {
    for (const entry of cells) {
      const { row, column, cell: current } = entry;
      if (current?.note !== undefined || current?.comment !== undefined) {
        const next = { ...current };
        delete next.note;
        delete next.comment;
        sheet.cells.set(row, column, next);
      }
      sheet.notes.delete(`${row}:${column}`);
    }
    sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) => !contains(range, thread.row, thread.column)));
    return;
  }
  if (params.family === 'hyperlinks') {
    for (const entry of cells) {
      const { row, column, cell: current } = entry;
      if (current?.hyperlink !== undefined || current?.hyperlinkDetail !== undefined) {
        const next = { ...current };
        delete next.hyperlink;
        delete next.hyperlinkDetail;
        sheet.cells.set(row, column, next);
      }
      sheet.hyperlinks.delete(`${row}:${column}`);
    }
    return;
  }
  for (const entry of cells) {
    const { row, column, cell: current } = entry;
    if (params.family === 'contents') sheet.cells.set(row, column, clearCellContents(current));
    else if (params.family === 'formats') sheet.cells.set(row, column, clearCellFormats(current));
    else sheet.cells.delete(row, column);
  }
  if (params.family === 'all') {
    for (const key of [...sheet.notes.keys()]) {
      const [row, column] = key.split(':').map(Number);
      if (contains(range, row!, column!)) sheet.notes.delete(key);
    }
    for (const key of [...sheet.hyperlinks.keys()]) {
      const [row, column] = key.split(':').map(Number);
      if (contains(range, row!, column!)) sheet.hyperlinks.delete(key);
    }
    sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) => !contains(range, thread.row, thread.column)));
  }
  if (params.family === 'formats' || params.family === 'all') {
    sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...sheetRuleRegistry.cropRules(sheet.conditionalFormats, range));
    sheet.dataValidations.splice(0, sheet.dataValidations.length, ...sheetRuleRegistry.cropRules(sheet.dataValidations, range));
  }
}

export function restoreClearRangeSnapshot(sheet: WorksheetModel, range: RangeRef, snapshot: ClearRangeSnapshot): void {
  const cells: Array<{ row: number; column: number }> = [];
  sheet.cells.forEach((_cell, row, column) => {
    if (contains(range, row, column)) cells.push({ row, column });
  });
  for (const { row, column } of cells) sheet.cells.delete(row, column);
  for (const key of [...sheet.notes.keys()]) {
    const [row, column] = key.split(':').map(Number);
    if (contains(range, row!, column!)) sheet.notes.delete(key);
  }
  for (const key of [...sheet.hyperlinks.keys()]) {
    const [row, column] = key.split(':').map(Number);
    if (contains(range, row!, column!)) sheet.hyperlinks.delete(key);
  }
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) => !contains(range, thread.row, thread.column)));
  for (const item of snapshot.cells) if (item.value !== undefined) sheet.cells.set(item.row, item.column, structuredClone(item.value));
  for (const item of snapshot.notes) sheet.notes.set(`${item.row}:${item.column}`, structuredClone(item.note));
  for (const item of snapshot.hyperlinks) sheet.hyperlinks.set(`${item.row}:${item.column}`, structuredClone(item.hyperlink));
  sheet.commentThreads.push(...structuredClone(snapshot.comments));
  if (snapshot.conditionalFormats !== undefined) sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...structuredClone(snapshot.conditionalFormats));
  if (snapshot.dataValidations !== undefined) sheet.dataValidations.splice(0, sheet.dataValidations.length, ...structuredClone(snapshot.dataValidations));
}
