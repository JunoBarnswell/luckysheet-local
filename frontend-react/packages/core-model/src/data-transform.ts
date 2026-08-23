import type {
  CellData,
  RangeRef,
  Row,
  WorksheetModel,
} from './index';
import type { CellNote, CommentThread, DrawingObject, SpillRange } from './domain';

export interface RowPermutation {
  /** Target row order expressed as source row indexes. */
  readonly sourceRows: readonly Row[];
  readonly range: RangeRef;
}

function inRange(range: RangeRef, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow
    && column >= range.startColumn && column <= range.endColumn;
}

function rowInRange(range: RangeRef, row: number): boolean {
  return row >= range.startRow && row <= range.endRow;
}

function remapRow(row: number, rowMap: ReadonlyMap<number, number>): number {
  return rowMap.get(row) ?? row;
}

function remapRangeRows(range: RangeRef, rowMap: ReadonlyMap<number, number>): RangeRef {
  if (range.startRow === range.endRow) {
    const next = remapRow(range.startRow, rowMap);
    return { ...range, startRow: next, endRow: next };
  }
  const rows: number[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) rows.push(remapRow(row, rowMap));
  return { ...range, startRow: Math.min(...rows), endRow: Math.max(...rows) };
}

function remapDrawingAnchor(drawing: DrawingObject, rowMap: ReadonlyMap<number, number>): void {
  if (drawing.anchor.kind === 'absolute') return;
  if (drawing.anchor.row !== undefined) drawing.anchor.row = remapRow(drawing.anchor.row, rowMap);
  if (drawing.anchor.endRow !== undefined) drawing.anchor.endRow = remapRow(drawing.anchor.endRow, rowMap);
}

function remapComment(thread: CommentThread, rowMap: ReadonlyMap<number, number>): CommentThread {
  return { ...thread, row: remapRow(thread.row, rowMap) };
}

function remapSpill(spill: SpillRange, rowMap: ReadonlyMap<number, number>): SpillRange {
  return {
    ...spill,
    anchor: { ...spill.anchor, row: remapRow(spill.anchor.row, rowMap) },
    range: remapRangeRows(spill.range, rowMap),
  };
}

/**
 * Move complete row records inside a selection according to a stable
 * permutation. This is used by sort and deliberately operates on the
 * canonical worksheet metadata as well as cells. Callers must validate that
 * partially intersecting metadata is supported before calling it.
 */
export function applyRowPermutation(sheet: WorksheetModel, permutation: RowPermutation): void {
  const { range, sourceRows } = permutation;
  const expectedCount = range.endRow - range.startRow + 1;
  if (sourceRows.length !== expectedCount) throw new Error('Row permutation length does not match the range');
  const expected = new Set<number>();
  for (let row = range.startRow; row <= range.endRow; row += 1) expected.add(row);
  if (sourceRows.some((row) => !expected.has(row)) || new Set(sourceRows).size !== expectedCount) {
    throw new Error('Row permutation must contain every selected row exactly once');
  }

  const rowMap = new Map<number, number>();
  sourceRows.forEach((sourceRow, targetOffset) => rowMap.set(sourceRow, range.startRow + targetOffset));

  const cellsByRow = new Map<number, Array<{ column: number; cell: CellData }>>();
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const entries: Array<{ column: number; cell: CellData }> = [];
    sheet.cells.forEach((cell, cellRow, column) => {
      if (cellRow === row && column >= range.startColumn && column <= range.endColumn) {
        entries.push({ column, cell: structuredClone(cell) });
      }
    });
    cellsByRow.set(row, entries);
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  sourceRows.forEach((sourceRow, targetOffset) => {
    const targetRow = range.startRow + targetOffset;
    for (const entry of cellsByRow.get(sourceRow) ?? []) sheet.cells.set(targetRow, entry.column, entry.cell);
  });

  // Cell-attached metadata follows the row record.
  const notes = new Map<string, CellNote>();
  for (const [key, note] of sheet.notes) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (!inRange(range, row, column)) {
      notes.set(key, note);
      continue;
    }
    notes.set(`${remapRow(row, rowMap)}:${column}`, note);
  }
  sheet.notes.clear();
  for (const [key, note] of notes) sheet.notes.set(key, note);
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.map((thread) =>
    rowInRange(range, thread.row) ? remapComment(thread, rowMap) : thread));
  for (const drawing of sheet.drawings) remapDrawingAnchor(drawing, rowMap);
  for (const sparkline of sheet.sparklines) {
    if (rowInRange(range, sparkline.anchor.row)) sparkline.anchor.row = remapRow(sparkline.anchor.row, rowMap);
    if (rowInRange(range, sparkline.sourceRange.startRow) && rowInRange(range, sparkline.sourceRange.endRow)) {
      sparkline.sourceRange = remapRangeRows(sparkline.sourceRange, rowMap);
    }
  }
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...sheet.spillRanges.map((spill) =>
    rowInRange(range, spill.anchor.row) ? remapSpill(spill, rowMap) : spill));
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    rule.ranges = rule.ranges.map((ruleRange) =>
      rowInRange(range, ruleRange.startRow) && rowInRange(range, ruleRange.endRow)
        ? remapRangeRows(ruleRange, rowMap)
        : ruleRange);
  }
  if (sheet.filter && rowInRange(range, sheet.filter.range.startRow) && rowInRange(range, sheet.filter.range.endRow)) {
    sheet.filter.range = remapRangeRows(sheet.filter.range, rowMap);
  }
  for (const table of sheet.sheetTables) {
    if (rowInRange(range, table.range.startRow) && rowInRange(range, table.range.endRow)) {
      table.range = remapRangeRows(table.range, rowMap);
    }
  }
  for (const merge of sheet.merges) {
    if (rowInRange(range, merge.range.startRow) && rowInRange(range, merge.range.endRow)) {
      merge.range = remapRangeRows(merge.range, rowMap);
      merge.anchor.row = remapRow(merge.anchor.row, rowMap);
    }
  }
  if (sheet.outline) {
    for (const group of sheet.outline.groups) {
      if (group.axis === 'row' && rowInRange(range, group.start) && rowInRange(range, group.end)) {
        group.start = remapRow(group.start, rowMap);
        group.end = remapRow(group.end, rowMap);
        if (group.start > group.end) [group.start, group.end] = [group.end, group.start];
      }
    }
  }
  for (const rule of sheet.protectionRules) {
    if (rule.range && rowInRange(range, rule.range.startRow) && rowInRange(range, rule.range.endRow)) {
      rule.range = remapRangeRows(rule.range, rowMap);
    }
  }
}

export function validatePermutationMetadata(sheet: WorksheetModel, range: RangeRef): void {
  const intersects = (candidate: RangeRef): boolean => candidate.sheetId === range.sheetId
    && candidate.startRow <= range.endRow && candidate.endRow >= range.startRow
    && candidate.startColumn <= range.endColumn && candidate.endColumn >= range.startColumn;
  for (const merge of sheet.merges) {
    if (intersects(merge.range) && !(merge.range.startRow >= range.startRow && merge.range.endRow <= range.endRow)) {
      throw new Error('Sort cannot partially intersect a merged range');
    }
  }
  for (const table of sheet.sheetTables) {
    if (intersects(table.range) && !(table.range.startRow === range.startRow && table.range.endRow === range.endRow)) {
      throw new Error('Sort requires the complete table row range');
    }
  }
  for (const group of sheet.outline?.groups ?? []) {
    if (group.axis === 'row' && group.start <= range.endRow && group.end >= range.startRow
      && !(group.start >= range.startRow && group.end <= range.endRow)) {
      throw new Error('Sort cannot partially intersect an outline group');
    }
  }
}
