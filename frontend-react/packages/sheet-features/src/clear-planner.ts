import type {
  CellData,
  CellHyperlink,
  CellNote,
  CommentThread,
  ConditionalFormatRule,
  RangeRef,
  WorksheetModel,
} from '@react-sheets/core-model';

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
  /** A complete rule snapshot is required when formats/all crop conditional formats. */
  conditionalFormats?: ConditionalFormatRule[];
}

export interface ClearRangePlan {
  params: ClearRangeParams;
  range: RangeRef;
  snapshot: ClearRangeSnapshot;
}

const MAX_CLEAR_CELLS = 100_000;

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

function intersects(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
}

/** Exact rectangle subtraction; the result never widens a non-contiguous remainder. */
export function subtractClearRange(source: RangeRef, clear: RangeRef): RangeRef[] {
  if (!intersects(source, clear)) return [structuredClone(source)];
  const top = Math.max(source.startRow, clear.startRow);
  const bottom = Math.min(source.endRow, clear.endRow);
  const left = Math.max(source.startColumn, clear.startColumn);
  const right = Math.min(source.endColumn, clear.endColumn);
  const result: RangeRef[] = [];
  if (source.startRow < top) result.push({ ...source, endRow: top - 1 });
  if (bottom < source.endRow) result.push({ ...source, startRow: bottom + 1 });
  if (source.startColumn < left) result.push({ ...source, startRow: top, endRow: bottom, endColumn: left - 1 });
  if (right < source.endColumn) result.push({ ...source, startRow: top, endRow: bottom, startColumn: right + 1 });
  return result;
}

/** Rebuilds every rule with the exact portions outside the clear rectangle. */
export function cropConditionalFormats(
  rules: readonly ConditionalFormatRule[],
  range: RangeRef,
): ConditionalFormatRule[] {
  return rules.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((candidate) => subtractClearRange(candidate, range));
    return ranges.length === 0 ? [] : [{ ...structuredClone(rule), ranges }];
  });
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
  const area = (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
  if (!Number.isSafeInteger(area) || area > MAX_CLEAR_CELLS) throw new Error('Clear range is too large');
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
  if (params.family === 'comments-and-notes') {
    for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const current = sheet.cells.get(row, column);
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
    for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const current = sheet.cells.get(row, column);
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
  for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const current = sheet.cells.get(row, column);
    if (!current) continue;
    if (params.family === 'contents') sheet.cells.set(row, column, clearCellContents(current));
    else if (params.family === 'formats') sheet.cells.set(row, column, clearCellFormats(current));
    else sheet.cells.delete(row, column);
  }
  if (params.family === 'all') {
    for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      sheet.notes.delete(`${row}:${column}`);
      sheet.hyperlinks.delete(`${row}:${column}`);
    }
    sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) => !contains(range, thread.row, thread.column)));
  }
  if (params.family === 'formats' || params.family === 'all') {
    sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...cropConditionalFormats(sheet.conditionalFormats, range));
  }
}

export function restoreClearRangeSnapshot(sheet: WorksheetModel, range: RangeRef, snapshot: ClearRangeSnapshot): void {
  for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    sheet.cells.delete(row, column);
    sheet.notes.delete(`${row}:${column}`);
    sheet.hyperlinks.delete(`${row}:${column}`);
  }
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) => !contains(range, thread.row, thread.column)));
  for (const item of snapshot.cells) if (item.value !== undefined) sheet.cells.set(item.row, item.column, structuredClone(item.value));
  for (const item of snapshot.notes) sheet.notes.set(`${item.row}:${item.column}`, structuredClone(item.note));
  for (const item of snapshot.hyperlinks) sheet.hyperlinks.set(`${item.row}:${item.column}`, structuredClone(item.hyperlink));
  sheet.commentThreads.push(...structuredClone(snapshot.comments));
  if (snapshot.conditionalFormats !== undefined) sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...structuredClone(snapshot.conditionalFormats));
}
