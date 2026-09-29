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
import { clearFormulaProvenance, sheetRuleRegistry } from '@react-sheets/core-model';

/** The only clear semantics accepted by the worksheet range command. */
export type ClearFamily = 'contents' | 'formats' | 'all' | 'comments-and-notes' | 'hyperlinks';

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  family: ClearFamily;
}

export interface ClearRangeSnapshot {
  /** Omitted for metadata-only clear operations so inverse replay never touches cell storage. */
  cells?: Array<{ row: number; column: number; value?: CellData }>;
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

function snapshotCells(sheet: WorksheetModel, range: RangeRef): NonNullable<ClearRangeSnapshot['cells']> {
  const cells: NonNullable<ClearRangeSnapshot['cells']> = [];
  sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn,
    (cell, row, column) => cells.push({ row, column, value: structuredClone(cell) }));
  return cells;
}

export function createClearRangePlan(sheet: WorksheetModel, input: ClearRangeParams): ClearRangePlan {
  const range = normalizeRange(input.range);
  if (range.sheetId !== sheet.id || input.sheetId !== sheet.id) throw new Error('Clear range targets another worksheet');
  const notes: ClearRangeSnapshot['notes'] = [];
  const hyperlinks: ClearRangeSnapshot['hyperlinks'] = [];
  const comments: ClearRangeSnapshot['comments'] = [];
  for (const entry of sheet.review.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
    if (entry.note) notes.push({ row: entry.row, column: entry.column, note: entry.note });
    comments.push(...entry.threads);
  }
  for (const entry of sheet.hyperlinks.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
    hyperlinks.push({ row: entry.row, column: entry.column, hyperlink: structuredClone(entry.hyperlink) });
  }
  return {
    params: { ...input, range },
    range,
    snapshot: {
      ...(input.family === 'comments-and-notes' || input.family === 'hyperlinks' ? {} : { cells: snapshotCells(sheet, range) }),
      notes,
      hyperlinks,
      comments,
      ...(input.family === 'formats' || input.family === 'all' ? { conditionalFormats: structuredClone(sheet.conditionalFormats) } : {}),
      ...(input.family === 'formats' || input.family === 'all' ? { dataValidations: structuredClone(sheet.dataValidations) } : {}),
    },
  };
}

export function clearCellContents(cell: CellData): CellData {
  const next = clearFormulaProvenance({ ...cell, value: null });
  delete next.formula;
  delete next.formulaValue;
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
  if (params.family === 'comments-and-notes') {
    for (const entry of sheet.review.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
      if (entry.note) sheet.review.removeNote(entry.row, entry.column);
      for (const thread of entry.threads) sheet.review.removeThread(thread.id);
    }
    return;
  }
  if (params.family === 'hyperlinks') {
    for (const entry of sheet.hyperlinks.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
      sheet.hyperlinks.delete(entry.key);
    }
    return;
  }
  sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (current, row, column) => {
    if (params.family === 'contents') sheet.cells.set(row, column, clearCellContents(current));
    else if (params.family === 'formats') sheet.cells.set(row, column, clearCellFormats(current));
    else sheet.cells.delete(row, column);
  });
  if (params.family === 'all') {
    for (const entry of sheet.review.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
      if (entry.note) sheet.review.removeNote(entry.row, entry.column);
      for (const thread of entry.threads) sheet.review.removeThread(thread.id);
    }
    for (const entry of sheet.hyperlinks.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
      sheet.hyperlinks.delete(entry.key);
    }
  }
  if (params.family === 'formats' || params.family === 'all') {
    sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...sheetRuleRegistry.cropRules(sheet.conditionalFormats, range));
    sheet.dataValidations.splice(0, sheet.dataValidations.length, ...sheetRuleRegistry.cropRules(sheet.dataValidations, range));
  }
}

export function restoreClearRangeSnapshot(sheet: WorksheetModel, range: RangeRef, snapshot: ClearRangeSnapshot): void {
  if (snapshot.cells !== undefined) {
    sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn,
      (_cell, row, column) => sheet.cells.delete(row, column));
  }
  for (const entry of sheet.review.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
    if (entry.note) sheet.review.removeNote(entry.row, entry.column);
    for (const thread of entry.threads) sheet.review.removeThread(thread.id);
  }
  for (const entry of sheet.hyperlinks.entriesInRange(range.startRow, range.endRow, range.startColumn, range.endColumn)) {
    sheet.hyperlinks.delete(entry.key);
  }
  for (const item of snapshot.cells ?? []) if (item.value !== undefined) sheet.cells.set(item.row, item.column, structuredClone(item.value));
  for (const item of snapshot.notes) sheet.review.setNote(item.row, item.column, item.note);
  for (const item of snapshot.hyperlinks) sheet.hyperlinks.set(`${item.row}:${item.column}`, structuredClone(item.hyperlink));
  for (const thread of snapshot.comments) sheet.review.addThread(thread);
  if (snapshot.conditionalFormats !== undefined) sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...structuredClone(snapshot.conditionalFormats));
  if (snapshot.dataValidations !== undefined) sheet.dataValidations.splice(0, sheet.dataValidations.length, ...structuredClone(snapshot.dataValidations));
}
