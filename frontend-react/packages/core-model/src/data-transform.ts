import type { CellData, RangeRef, Row, WorksheetModel } from './index';
import { noteCellKey } from './index';
import type { CellNote, CommentThread, DrawingObject, SpillRange } from './domain';

/** Canonical, prevalidated permutation shared by local execution and replay. */
export interface RowPermutationPlan {
  readonly range: RangeRef;
  readonly sourceRows: readonly Row[];
  readonly sourceToTarget: ReadonlyMap<Row, Row>;
}

const MAX_SEGMENT_CELLS = 100_000;

function normalizeRange(range: RangeRef): RangeRef {
  if (![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isInteger)) {
    throw new Error('Row permutation range must contain integer coordinates');
  }
  if (range.startRow < 0 || range.startColumn < 0) throw new Error('Row permutation range is outside worksheet bounds');
  return { ...range, startRow: Math.min(range.startRow, range.endRow), endRow: Math.max(range.startRow, range.endRow), startColumn: Math.min(range.startColumn, range.endColumn), endColumn: Math.max(range.startColumn, range.endColumn) };
}

export function createRowPermutationPlan(range: RangeRef, sourceRows: readonly Row[]): RowPermutationPlan {
  const normalized = normalizeRange(range);
  const expectedCount = normalized.endRow - normalized.startRow + 1;
  if (sourceRows.length !== expectedCount) throw new Error('Row permutation length does not match the range');
  const expected = new Set<number>();
  for (let row = normalized.startRow; row <= normalized.endRow; row += 1) expected.add(row);
  const sourceToTarget = new Map<Row, Row>();
  sourceRows.forEach((sourceRow, targetOffset) => {
    if (!Number.isInteger(sourceRow) || !expected.has(sourceRow) || sourceToTarget.has(sourceRow)) throw new Error('Row permutation must contain every selected row exactly once');
    sourceToTarget.set(sourceRow, normalized.startRow + targetOffset);
  });
  if (sourceToTarget.size !== expectedCount) throw new Error('Row permutation must contain every selected row exactly once');
  return Object.freeze({ range: Object.freeze(normalized), sourceRows: Object.freeze([...sourceRows]), sourceToTarget });
}

function inRange(range: RangeRef, row: number, column: number): boolean {
  return range.startRow <= row && row <= range.endRow && range.startColumn <= column && column <= range.endColumn;
}

function rangesIntersect(a: RangeRef, b: RangeRef): boolean {
  return a.sheetId === b.sheetId && a.startRow <= b.endRow && b.startRow <= a.endRow && a.startColumn <= b.endColumn && b.startColumn <= a.endColumn;
}

function remapRow(row: number, plan: RowPermutationPlan): number { return plan.sourceToTarget.get(row) ?? row; }

function cloneRange(range: RangeRef, startRow: number, endRow: number, startColumn = range.startColumn, endColumn = range.endColumn): RangeRef {
  return { ...range, startRow, endRow, startColumn, endColumn };
}

function mergeExactSegments(segments: RangeRef[]): RangeRef[] {
  const result = [...segments];
  let changed = true;
  while (changed) {
    changed = false;
    outer: for (let left = 0; left < result.length; left += 1) {
      for (let right = left + 1; right < result.length; right += 1) {
        const a = result[left]!; const b = result[right]!;
        const sameRows = a.startRow === b.startRow && a.endRow === b.endRow;
        const sameColumns = a.startColumn === b.startColumn && a.endColumn === b.endColumn;
        if ((sameRows && (a.endColumn + 1 === b.startColumn || b.endColumn + 1 === a.startColumn))
          || (sameColumns && (a.endRow + 1 === b.startRow || b.endRow + 1 === a.startRow))) {
          result[left] = {
            ...a,
            startRow: Math.min(a.startRow, b.startRow), endRow: Math.max(a.endRow, b.endRow),
            startColumn: Math.min(a.startColumn, b.startColumn), endColumn: Math.max(a.endColumn, b.endColumn),
          };
          result.splice(right, 1);
          changed = true;
          break outer;
        }
      }
    }
  }
  return result;
}

/** Exact disjoint rectangle cover. It never uses min/max over non-contiguous rows. */
function remapRangeExact(range: RangeRef, plan: RowPermutationPlan): RangeRef[] {
  if (range.sheetId !== plan.range.sheetId || !rangesIntersect(range, plan.range)) return [structuredClone(range)];
  const selected = plan.range;
  const firstRow = Math.max(range.startRow, selected.startRow);
  const lastRow = Math.min(range.endRow, selected.endRow);
  const firstColumn = Math.max(range.startColumn, selected.startColumn);
  const lastColumn = Math.min(range.endColumn, selected.endColumn);
  const result: RangeRef[] = [];
  if (range.startRow < firstRow) result.push(cloneRange(range, range.startRow, firstRow - 1));
  if (lastRow < range.endRow) result.push(cloneRange(range, lastRow + 1, range.endRow));
  if (range.startColumn < firstColumn) result.push(cloneRange(range, firstRow, lastRow, range.startColumn, firstColumn - 1));
  if (lastColumn < range.endColumn) result.push(cloneRange(range, firstRow, lastRow, lastColumn + 1, range.endColumn));
  const area = (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
  if (!Number.isSafeInteger(area) || area > MAX_SEGMENT_CELLS) throw new Error('Row permutation metadata range cannot be represented exactly within the bounded plan');
  const rows = new Map<number, [number, number]>();
  for (let row = firstRow; row <= lastRow; row += 1) rows.set(remapRow(row, plan), [firstColumn, lastColumn]);
  const orderedRows = [...rows.keys()].sort((a, b) => a - b);
  let i = 0;
  while (i < orderedRows.length) {
    const start = orderedRows[i]!;
    const columns = rows.get(start)!;
    let end = start;
    while (i + 1 < orderedRows.length && orderedRows[i + 1] === end + 1 && rows.get(orderedRows[i + 1]!)?.[0] === columns[0] && rows.get(orderedRows[i + 1]!)?.[1] === columns[1]) {
      end = orderedRows[++i]!;
    }
    result.push(cloneRange(range, start, end, columns[0], columns[1]));
    i += 1;
  }
  return mergeExactSegments(result);
}

function remapRangeList(ranges: readonly RangeRef[], plan: RowPermutationPlan): RangeRef[] { return ranges.flatMap((range) => remapRangeExact(range, plan)); }

function remapSingleRange(owner: string, range: RangeRef, plan: RowPermutationPlan): RangeRef {
  const segments = remapRangeExact(range, plan);
  if (segments.length !== 1) throw new Error(`Sort cannot exactly remap ${owner} into a single range`);
  return segments[0]!;
}

function remapDrawingAnchor(drawing: DrawingObject, plan: RowPermutationPlan): DrawingObject {
  if (drawing.anchor.kind === 'absolute' || drawing.anchor.row === undefined || drawing.anchor.column === undefined) return drawing;
  const startInside = inRange(plan.range, drawing.anchor.row, drawing.anchor.column);
  const endRow = drawing.anchor.endRow ?? drawing.anchor.row;
  const endColumn = drawing.anchor.endColumn ?? drawing.anchor.column;
  const endInside = inRange(plan.range, endRow, endColumn);
  if (!startInside && !endInside) return drawing;
  if (!startInside || !endInside) throw new Error(`Sort cannot exactly remap drawing ${drawing.id}`);
  return { ...drawing, anchor: { ...drawing.anchor, row: remapRow(drawing.anchor.row, plan), endRow: drawing.anchor.endRow === undefined ? undefined : remapRow(drawing.anchor.endRow, plan) } };
}

function remapSpill(spill: SpillRange, plan: RowPermutationPlan): SpillRange {
  const segments = remapRangeExact(spill.range, plan);
  if (segments.length !== 1) throw new Error('Sort cannot exactly remap a spill range with disjoint segments');
  const anchorInside = inRange(plan.range, spill.anchor.row, spill.anchor.column);
  if (!anchorInside && segments[0]!.startRow !== spill.range.startRow) throw new Error('Sort cannot detach a spill anchor from its range');
  return { ...spill, anchor: anchorInside ? { ...spill.anchor, row: remapRow(spill.anchor.row, plan) } : spill.anchor, range: segments[0]! };
}

function remapCellMap<T>(source: ReadonlyMap<string, T>, plan: RowPermutationPlan): Map<string, T> {
  const next = new Map<string, T>();
  for (const [key, value] of source) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText); const column = Number(columnText);
    const nextKey = noteCellKey(inRange(plan.range, row, column) ? remapRow(row, plan) : row, column);
    if (next.has(nextKey)) throw new Error(`Sort produced duplicate cell metadata at ${nextKey}`);
    next.set(nextKey, value);
  }
  return next;
}

/** Validate all owners before the first cell changes. */
export function validatePermutationMetadata(sheet: WorksheetModel, plan: RowPermutationPlan): void {
  const range = plan.range;
  if (range.sheetId !== sheet.id || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw new Error('Row permutation range is outside worksheet bounds');
  for (const merge of sheet.merges) {
    if (!rangesIntersect(merge.range, range)) continue;
    if (!(merge.range.startRow >= range.startRow && merge.range.endRow <= range.endRow)) throw new Error('Sort cannot partially intersect a merged range');
    remapSingleRange('merge', merge.range, plan);
  }
  for (const table of sheet.sheetTables) {
    if (rangesIntersect(table.range, range)) {
      if (!(table.range.startRow === range.startRow && table.range.endRow === range.endRow)) throw new Error('Sort requires the complete table row range');
      remapSingleRange(`table ${table.id}`, table.range, plan);
    }
    if (table.autoFilter) remapSingleRange(`table ${table.id} filter`, table.autoFilter.range, plan);
  }
  for (const group of sheet.outline?.groups ?? []) if (group.axis === 'row' && group.start <= range.endRow && group.end >= range.startRow && !(group.start >= range.startRow && group.end <= range.endRow)) throw new Error('Sort cannot partially intersect an outline group');
  for (const drawing of sheet.drawings) remapDrawingAnchor(drawing, plan);
  for (const sparkline of sheet.sparklines) remapSingleRange(`sparkline ${sparkline.id}`, sparkline.sourceRange, plan);
  for (const spill of sheet.spillRanges) remapSpill(spill, plan);
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const ruleRange of rule.ranges) remapRangeExact(ruleRange, plan);
  }
  for (const rule of sheet.dataValidations) if (rule.listSource?.kind === 'range') remapSingleRange(`validation ${rule.id} source`, rule.listSource.range, plan);
  if (sheet.autoFilter) remapSingleRange('auto filter', sheet.autoFilter.range, plan);
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range') remapSingleRange(`pivot ${pivot.id} source`, pivot.source.range, plan);
    if (pivot.source.kind === 'worksheet-ranges') for (const source of pivot.source.ranges) remapSingleRange(`pivot ${pivot.id} source`, source.range, plan);
  }
  for (const rule of sheet.protectionRules) if (rule.range) remapSingleRange(`protection ${rule.id}`, rule.range, plan);
  if (sheet.bandedRule) remapSingleRange('banded rule', sheet.bandedRule.range, plan);
}

export function applyRowPermutation(sheet: WorksheetModel, plan: RowPermutationPlan): void {
  validatePermutationMetadata(sheet, plan);
  const { range, sourceRows } = plan;
  const cellsByRow = new Map<number, Array<{ column: number; cell: CellData }>>();
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const entries: Array<{ column: number; cell: CellData }> = [];
    sheet.cells.forEach((cell, cellRow, column) => { if (cellRow === row && column >= range.startColumn && column <= range.endColumn) entries.push({ column, cell: structuredClone(cell) }); });
    cellsByRow.set(row, entries);
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) sheet.cells.delete(row, column);
  sourceRows.forEach((sourceRow, targetOffset) => { for (const entry of cellsByRow.get(sourceRow) ?? []) sheet.cells.set(range.startRow + targetOffset, entry.column, entry.cell); });

  const notes = remapCellMap(sheet.notes, plan); sheet.notes.clear(); for (const [key, value] of notes) sheet.notes.set(key, value as CellNote);
  const hyperlinks = remapCellMap(sheet.hyperlinks, plan); sheet.hyperlinks.clear(); for (const [key, value] of hyperlinks) sheet.hyperlinks.set(key, value);
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.map((thread: CommentThread) => inRange(range, thread.row, thread.column) ? { ...thread, row: remapRow(thread.row, plan) } : thread));
  for (const drawing of sheet.drawings) Object.assign(drawing, remapDrawingAnchor(drawing, plan));
  for (const sparkline of sheet.sparklines) { sparkline.sourceRange = remapSingleRange(`sparkline ${sparkline.id}`, sparkline.sourceRange, plan); if (inRange(range, sparkline.anchor.row, sparkline.anchor.column)) sparkline.anchor.row = remapRow(sparkline.anchor.row, plan); }
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...sheet.spillRanges.map((spill) => remapSpill(spill, plan)));
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) rule.ranges = remapRangeList(rule.ranges, plan);
  for (const rule of sheet.dataValidations) if (rule.listSource?.kind === 'range') rule.listSource.range = remapSingleRange(`validation ${rule.id} source`, rule.listSource.range, plan);
  if (sheet.autoFilter) sheet.autoFilter.range = remapSingleRange('auto filter', sheet.autoFilter.range, plan);
  for (const table of sheet.sheetTables) { table.range = remapSingleRange(`table ${table.id}`, table.range, plan); if (table.autoFilter) table.autoFilter.range = remapSingleRange(`table ${table.id} filter`, table.autoFilter.range, plan); }
  for (const pivot of sheet.pivots) { if (pivot.source.kind === 'worksheet-range') pivot.source.range = remapSingleRange(`pivot ${pivot.id} source`, pivot.source.range, plan); if (pivot.source.kind === 'worksheet-ranges') for (const source of pivot.source.ranges) source.range = remapSingleRange(`pivot ${pivot.id} source`, source.range, plan); if (pivot.target.sheetId === sheet.id && inRange(range, pivot.target.anchor.row, pivot.target.anchor.column)) pivot.target.anchor.row = remapRow(pivot.target.anchor.row, plan); }
  for (const merge of sheet.merges) { merge.range = remapSingleRange('merge', merge.range, plan); if (inRange(range, merge.anchor.row, merge.anchor.column)) merge.anchor.row = remapRow(merge.anchor.row, plan); }
  for (const group of sheet.outline?.groups ?? []) if (group.axis === 'row' && group.start >= range.startRow && group.end <= range.endRow) { const mapped = remapRangeExact({ sheetId: sheet.id, startRow: group.start, endRow: group.end, startColumn: range.startColumn, endColumn: range.endColumn }, plan); if (mapped.length !== 1) throw new Error('Sort cannot exactly remap outline group'); group.start = mapped[0]!.startRow; group.end = mapped[0]!.endRow; }
  for (const rule of sheet.protectionRules) if (rule.range) rule.range = remapSingleRange(`protection ${rule.id}`, rule.range, plan);
  if (sheet.bandedRule) sheet.bandedRule.range = remapSingleRange('banded rule', sheet.bandedRule.range, plan);
}
