import type { CellData, RangeRef, Row, Column, WorksheetModel } from './index';
import type { CellNote, DrawingObject, StructuralTransformParams, CommentThread, SheetTableModel, SpillRange, ProtectionRule, OutlineGroup } from './domain';
import { WorkbookModel, noteCellKey } from './index';
import {
  formatFormula,
  mapAstReferences,
  offsetAst,
  parseFormula,
  remapAst,
  type ParsedCellReference,
  type StructuralShift,
} from '@react-sheets/formula-engine';

export interface StructuralTransformResult {
  removedCells: Array<{ row: Row; column: Column; cell: CellData }>;
}

/** 结构变换唯一入口 — 一次更新 cells/merges/CF/validation/filter/freeze/charts/pivots/tables/names/drawings/notes/comments/protection/公式引用 */
export class StructuralTransform {
  static apply(workbook: WorkbookModel, params: StructuralTransformParams): StructuralTransformResult {
    const sheet = workbook.getSheet(params.sheetId);
    switch (params.kind) {
      case 'insert-rows':
        return applyAxis(workbook, sheet, 'row', params.at ?? 0, params.count ?? 1, 1);
      case 'delete-rows':
        return applyAxis(workbook, sheet, 'row', params.at ?? 0, params.count ?? 1, -1);
      case 'insert-columns':
        return applyAxis(workbook, sheet, 'column', params.at ?? 0, params.count ?? 1, 1);
      case 'delete-columns':
        return applyAxis(workbook, sheet, 'column', params.at ?? 0, params.count ?? 1, -1);
      case 'move-range':
        return applyMoveRange(workbook, sheet, params.sourceRange!, params.targetOrigin!);
      case 'shift-cells-down':
      case 'shift-cells-up':
      case 'shift-cells-right':
      case 'shift-cells-left':
        if (!params.sourceRange) throw new Error('Shift-cells requires sourceRange');
        return applyShiftCells(workbook, sheet, params.sourceRange, params.kind);
      default:
        throw new Error(`Unknown structural op: ${(params as StructuralTransformParams).kind}`);
    }
  }
}

function validateAxisBounds(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  if (!Number.isInteger(at) || at < 0) throw new Error(`Structural ${axis} index must be a non-negative integer`);
  if (!Number.isInteger(count) || count <= 0) throw new Error('Structural count must be a positive integer');
  if (direction === -1) {
    const limit = axis === 'row' ? sheet.rowCount : sheet.columnCount;
    if (at >= limit || at + count > limit) {
      throw new Error(`Cannot delete ${count} ${axis}(s) at ${at}: outside worksheet bounds`);
    }
  }
}

function intersectsAxisRange(
  range: RangeRef,
  axis: 'row' | 'column',
  at: number,
  count: number,
): boolean {
  const start = axis === 'row' ? range.startRow : range.startColumn;
  const end = axis === 'row' ? range.endRow : range.endColumn;
  return start <= at + count - 1 && end >= at;
}

/**
 * A structural delete must never silently drop a non-cell object.  Objects
 * whose ranges merely move are handled by the transform helpers below; an
 * object anchored in deleted coordinates is rejected before any cell is
 * changed.  The caller can then choose an explicit object-delete operation.
 */
function validateAxisMetadataPreservation(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  if (direction === 1) return;
  const deleted = (position: number): boolean => position >= at && position < at + count;
  for (const sparkline of sheet.sparklines) {
    const position = axis === 'row' ? sparkline.anchor.row : sparkline.anchor.column;
    if (deleted(position)) throw new Error(`Cannot delete ${axis} ${position}: sparkline ${sparkline.id} would be lost`);
  }
  for (const pivot of sheet.pivots) {
    if (pivot.targetAnchor) {
      const position = axis === 'row' ? pivot.targetAnchor.row : pivot.targetAnchor.column;
      if (deleted(position)) throw new Error(`Cannot delete ${axis} ${position}: pivot ${pivot.id} would be lost`);
    }
  }
  for (const spill of sheet.spillRanges) {
    const position = axis === 'row' ? spill.anchor.row : spill.anchor.column;
    if (deleted(position)) throw new Error(`Cannot delete ${axis} ${position}: spill range would be lost`);
  }
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute') continue;
    const start = axis === 'row' ? drawing.anchor.row : drawing.anchor.column;
    const end = axis === 'row' ? drawing.anchor.endRow : drawing.anchor.endColumn;
    if ((start !== undefined && deleted(start)) || (end !== undefined && deleted(end))) {
      throw new Error(`Cannot delete ${axis} ${at}: drawing ${drawing.id} would lose its anchor`);
    }
  }
  for (const [key] of sheet.notes) {
    const [row, column] = key.split(':').map(Number);
    if (deleted(axis === 'row' ? row! : column!)) {
      throw new Error(`Cannot delete ${axis} ${at}: note at ${key} would be lost`);
    }
  }
  for (const thread of sheet.commentThreads) {
    if (deleted(axis === 'row' ? thread.row : thread.column)) {
      throw new Error(`Cannot delete ${axis} ${at}: comment thread ${thread.id} would be lost`);
    }
  }
  for (const table of sheet.sheetTables) {
    if (intersectsAxisRange(table.range, axis, at, count)) {
      throw new Error(`Cannot delete ${axis} ${at}: table ${table.id} requires an explicit table operation`);
    }
  }
}

function validateShiftPreservation(
  sheet: WorksheetModel,
  selection: RangeRef,
  kind: 'shift-cells-down' | 'shift-cells-up' | 'shift-cells-right' | 'shift-cells-left',
): void {
  const rowDelta = kind === 'shift-cells-down' ? 1 : kind === 'shift-cells-up' ? -1 : 0;
  const columnDelta = kind === 'shift-cells-right' ? 1 : kind === 'shift-cells-left' ? -1 : 0;
  const inside = (row: number, column: number): boolean =>
    row >= selection.startRow && row <= selection.endRow && column >= selection.startColumn && column <= selection.endColumn;
  const remainsInside = (row: number, column: number): boolean => inside(row + rowDelta, column + columnDelta);
  for (const sparkline of sheet.sparklines) {
    if (inside(sparkline.anchor.row, sparkline.anchor.column) && !remainsInside(sparkline.anchor.row, sparkline.anchor.column)) {
      throw new Error(`Cannot shift ${selection.sheetId}: sparkline ${sparkline.id} would leave the selected range`);
    }
  }
  for (const spill of sheet.spillRanges) {
    if (inside(spill.anchor.row, spill.anchor.column) && !remainsInside(spill.anchor.row, spill.anchor.column)) {
      throw new Error(`Cannot shift ${selection.sheetId}: spill range would leave the selected range`);
    }
  }
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute') continue;
    const row = drawing.anchor.row;
    const column = drawing.anchor.column;
    if (row !== undefined && column !== undefined && inside(row, column) && !remainsInside(row, column)) {
      throw new Error(`Cannot shift ${selection.sheetId}: drawing ${drawing.id} would leave the selected range`);
    }
    const endRow = drawing.anchor.endRow;
    const endColumn = drawing.anchor.endColumn;
    if (endRow !== undefined && endColumn !== undefined && inside(endRow, endColumn) && !remainsInside(endRow, endColumn)) {
      throw new Error(`Cannot shift ${selection.sheetId}: drawing ${drawing.id} would lose its extent`);
    }
  }
  for (const [key] of sheet.notes) {
    const [row, column] = key.split(':').map(Number);
    if (row !== undefined && column !== undefined && inside(row, column) && !remainsInside(row, column)) {
      throw new Error(`Cannot shift ${selection.sheetId}: note at ${key} would leave the selected range`);
    }
  }
  for (const thread of sheet.commentThreads) {
    if (inside(thread.row, thread.column) && !remainsInside(thread.row, thread.column)) {
      throw new Error(`Cannot shift ${selection.sheetId}: comment thread ${thread.id} would leave the selected range`);
    }
  }
}

function applyAxis(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): StructuralTransformResult {
  validateAxisBounds(sheet, axis, at, count, direction);
  validateAxisMetadataPreservation(sheet, axis, at, count, direction);
  if (count <= 0) return { removedCells: [] };
  const end = at + count - 1;
  let removed: Array<{ row: Row; column: Column; cell: CellData }> = [];

  if (direction === 1) {
    if (axis === 'row') {
      sheet.cells.shiftRows(at, count, 1);
      sheet.rowCount += count;
    } else {
      sheet.cells.shiftColumns(at, count, 1);
      sheet.columnCount += count;
    }
  } else if (axis === 'row') {
    removed = sheet.cells.extractRegion(at, end, 0, Math.max(sheet.columnCount - 1, 0));
    sheet.cells.shiftRows(end + 1, count, -1);
    sheet.rowCount = Math.max(1, sheet.rowCount - count);
  } else {
    removed = sheet.cells.extractRegion(0, Math.max(sheet.rowCount - 1, 0), at, end);
    sheet.cells.shiftColumns(end + 1, count, -1);
    sheet.columnCount = Math.max(1, sheet.columnCount - count);
  }

  shiftMerges(sheet, axis, at, count, direction);
  shiftRuleRanges(sheet.conditionalFormats, axis, at, count, direction);
  shiftRuleRanges(sheet.dataValidations, axis, at, count, direction);
  shiftFilter(sheet, axis, at, count, direction);
  shiftFreeze(sheet, axis, at, count, direction);
  shiftHiddenAndSizes(sheet, axis, at, count, direction);
  shiftSparklines(sheet, axis, at, count, direction);
  shiftPivots(sheet, axis, at, count, direction);
  shiftChartPayloads(sheet, axis, at, count, direction);
  shiftDrawings(sheet, axis, at, count, direction);
  shiftSheetTables(sheet, axis, at, count, direction);
  shiftNotes(sheet, axis, at, count, direction);
  shiftComments(sheet, axis, at, count, direction);
  shiftSpills(sheet, axis, at, count, direction);
  shiftProtection(sheet, axis, at, count, direction);
  shiftBanded(sheet, axis, at, count, direction);
  shiftOutline(sheet, axis, at, count, direction);
  rewriteFormulas(workbook, sheet.id, axis, at, count, direction);
  rewriteDefinedNames(workbook, sheet, {
    axis,
    at,
    count,
    op: direction === 1 ? 'insert' : 'delete',
  });
  return { removedCells: removed };
}

function applyShiftCells(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  range: RangeRef,
  kind: 'shift-cells-down' | 'shift-cells-up' | 'shift-cells-right' | 'shift-cells-left',
): StructuralTransformResult {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startColumn = Math.min(range.startColumn, range.endColumn);
  const endColumn = Math.max(range.startColumn, range.endColumn);
  const selection: RangeRef = {
    sheetId: range.sheetId,
    startRow,
    endRow,
    startColumn,
    endColumn,
  };
  validateShiftPreservation(sheet, selection, kind);
  const isVertical = kind === 'shift-cells-down' || kind === 'shift-cells-up';
  const delta = kind === 'shift-cells-down' || kind === 'shift-cells-right' ? 1 : -1;
  const rowDelta = isVertical ? delta : 0;
  const columnDelta = isVertical ? 0 : delta;

  const sourceCells: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
      sourceCells.push({ row, column, cell: structuredClone(cell) });
    }
  });

  // Clear the bounded region before placing shifted cells. This avoids stale
  // cells at the newly-empty edge and makes the transform deterministic even
  // when source and destination overlap.
  for (let row = startRow; row <= endRow; row++) {
    for (let column = startColumn; column <= endColumn; column++) sheet.cells.delete(row, column);
  }

  const removedCells: Array<{ row: Row; column: Column; cell: CellData }> = [];
  for (const entry of sourceCells) {
    const nextRow = entry.row + rowDelta;
    const nextColumn = entry.column + columnDelta;
    const inside = nextRow >= startRow && nextRow <= endRow && nextColumn >= startColumn && nextColumn <= endColumn;
    if (!inside) {
      removedCells.push(entry);
      continue;
    }
    const cell = entry.cell.formula
      ? { ...entry.cell, formula: offsetFormulaText(entry.cell.formula, rowDelta, columnDelta) }
      : entry.cell;
    sheet.cells.set(nextRow, nextColumn, cell);
  }

  shiftBoundedMetadata(sheet, selection, rowDelta, columnDelta);
  rewriteReferencesForMovedRegion(workbook, sheet, selection, {
    ...selection,
    startRow: selection.startRow + rowDelta,
    endRow: selection.endRow + rowDelta,
    startColumn: selection.startColumn + columnDelta,
    endColumn: selection.endColumn + columnDelta,
  }, rowDelta, columnDelta);
  return { removedCells };
}

function offsetFormulaText(formula: string, rowOffset: number, columnOffset: number): string {
  if (!formula.trim().startsWith('=')) return formula;
  try {
    return formatFormula(offsetAst(parseFormula(formula), rowOffset, columnOffset));
  } catch {
    return formula;
  }
}

function rangeContains(outer: RangeRef, inner: RangeRef): boolean {
  return outer.sheetId === inner.sheetId
    && inner.startRow >= outer.startRow
    && inner.endRow <= outer.endRow
    && inner.startColumn >= outer.startColumn
    && inner.endColumn <= outer.endColumn;
}

function shiftContainedRange(range: RangeRef, selection: RangeRef, rowDelta: number, columnDelta: number): void {
  if (!rangeContains(selection, range)) return;
  range.startRow += rowDelta;
  range.endRow += rowDelta;
  range.startColumn += columnDelta;
  range.endColumn += columnDelta;
}

function shiftBoundedMetadata(sheet: WorksheetModel, selection: RangeRef, rowDelta: number, columnDelta: number): void {
  for (const merge of sheet.merges) {
    const contained = rangeContains(selection, merge.range);
    shiftContainedRange(merge.range, selection, rowDelta, columnDelta);
    if (contained) {
      merge.anchor.row += rowDelta;
      merge.anchor.column += columnDelta;
    }
  }
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) shiftContainedRange(range, selection, rowDelta, columnDelta);
  }
  if (sheet.filter) shiftContainedRange(sheet.filter.range, selection, rowDelta, columnDelta);
  for (const table of sheet.sheetTables) shiftContainedRange(table.range, selection, rowDelta, columnDelta);
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind !== 'chart') continue;
    for (const range of payload.sourceRanges) shiftContainedRange(range, selection, rowDelta, columnDelta);
    if (payload.categoryRange) shiftContainedRange(payload.categoryRange, selection, rowDelta, columnDelta);
    for (const series of payload.series ?? []) shiftContainedRange(series.range, selection, rowDelta, columnDelta);
  }
  for (const pivot of sheet.pivots) {
    shiftContainedRange(pivot.sourceRange, selection, rowDelta, columnDelta);
    if (pivot.dataSource?.kind === 'worksheet-range') shiftContainedRange(pivot.dataSource.range, selection, rowDelta, columnDelta);
    if (pivot.dataSource?.kind === 'worksheet-ranges') {
      for (const range of pivot.dataSource.ranges) shiftContainedRange(range, selection, rowDelta, columnDelta);
    }
    if (pivot.targetAnchor && pivot.targetAnchor.row >= selection.startRow && pivot.targetAnchor.row <= selection.endRow
      && pivot.targetAnchor.column >= selection.startColumn && pivot.targetAnchor.column <= selection.endColumn) {
      pivot.targetAnchor.row += rowDelta;
      pivot.targetAnchor.column += columnDelta;
    }
  }
  for (const sparkline of sheet.sparklines) {
    shiftContainedRange(sparkline.sourceRange, selection, rowDelta, columnDelta);
    if (sparkline.anchor.row >= selection.startRow && sparkline.anchor.row <= selection.endRow
      && sparkline.anchor.column >= selection.startColumn && sparkline.anchor.column <= selection.endColumn) {
      sparkline.anchor.row += rowDelta;
      sparkline.anchor.column += columnDelta;
    }
  }
  for (const spill of sheet.spillRanges) {
    shiftContainedRange(spill.range, selection, rowDelta, columnDelta);
    if (spill.anchor.row >= selection.startRow && spill.anchor.row <= selection.endRow
      && spill.anchor.column >= selection.startColumn && spill.anchor.column <= selection.endColumn) {
      spill.anchor.row += rowDelta;
      spill.anchor.column += columnDelta;
    }
  }
  for (const rule of sheet.protectionRules) {
    if (rule.range) shiftContainedRange(rule.range, selection, rowDelta, columnDelta);
  }
  if (sheet.bandedRule) shiftContainedRange(sheet.bandedRule.range, selection, rowDelta, columnDelta);
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute') continue;
    const row = drawing.anchor.row;
    const column = drawing.anchor.column;
    if (row == null || column == null) continue;
    if (row < selection.startRow || row > selection.endRow || column < selection.startColumn || column > selection.endColumn) continue;
    drawing.anchor.row = row + rowDelta;
    drawing.anchor.column = column + columnDelta;
    if (drawing.anchor.endRow != null) drawing.anchor.endRow += rowDelta;
    if (drawing.anchor.endColumn != null) drawing.anchor.endColumn += columnDelta;
  }

  const nextNotes = new Map<string, CellNote>();
  for (const [key, note] of sheet.notes) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (row >= selection.startRow && row <= selection.endRow && column >= selection.startColumn && column <= selection.endColumn) {
      const nextRow = row + rowDelta;
      const nextColumn = column + columnDelta;
      if (nextRow >= selection.startRow && nextRow <= selection.endRow && nextColumn >= selection.startColumn && nextColumn <= selection.endColumn) {
        nextNotes.set(noteCellKey(nextRow, nextColumn), note);
      }
    } else {
      nextNotes.set(key, note);
    }
  }
  sheet.notes.clear();
  for (const [key, note] of nextNotes) sheet.notes.set(key, note);
  const nextThreads: CommentThread[] = [];
  for (const thread of sheet.commentThreads) {
    if (thread.row >= selection.startRow && thread.row <= selection.endRow && thread.column >= selection.startColumn && thread.column <= selection.endColumn) {
      const nextRow = thread.row + rowDelta;
      const nextColumn = thread.column + columnDelta;
      if (nextRow < selection.startRow || nextRow > selection.endRow || nextColumn < selection.startColumn || nextColumn > selection.endColumn) continue;
      nextThreads.push({ ...thread, row: nextRow, column: nextColumn });
      continue;
    }
    nextThreads.push(thread);
  }
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...nextThreads);

  if (rowDelta !== 0) {
    remapBoundedSet(sheet.hiddenRows, selection.startRow, selection.endRow, rowDelta);
    remapBoundedMap(sheet.rowHeights, selection.startRow, selection.endRow, rowDelta);
  }
  if (columnDelta !== 0) {
    remapBoundedSet(sheet.hiddenColumns, selection.startColumn, selection.endColumn, columnDelta);
    remapBoundedMap(sheet.columnWidths, selection.startColumn, selection.endColumn, columnDelta);
  }
}

function remapBoundedSet(set: Set<number>, start: number, end: number, delta: number): void {
  const next = new Set<number>();
  for (const value of set) next.add(value >= start && value <= end ? value + delta : value);
  set.clear();
  for (const value of next) set.add(value);
}

function remapBoundedMap(map: Record<number, number>, start: number, end: number, delta: number): void {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const numericKey = Number(key);
    next[numericKey >= start && numericKey <= end ? numericKey + delta : numericKey] = value;
  }
  for (const key of Object.keys(map)) delete map[Number(key)];
  Object.assign(map, next);
}

function referenceBelongsToSheet(reference: ParsedCellReference, owner: WorksheetModel, target: WorksheetModel): boolean {
  if (reference.sheetId === undefined) return owner.id === target.id;
  const normalized = reference.sheetId.trim().toLocaleLowerCase();
  return normalized === target.id.toLocaleLowerCase() || normalized === target.name.toLocaleLowerCase();
}

function rewriteReferencesForMovedRegion(
  workbook: WorkbookModel,
  targetSheet: WorksheetModel,
  selection: RangeRef,
  destination: RangeRef,
  rowDelta: number,
  columnDelta: number,
): void {
  const mapper = (owner: WorksheetModel) => (reference: ParsedCellReference): ParsedCellReference => {
    if (!referenceBelongsToSheet(reference, owner, targetSheet)) return reference;
    if (reference.row < selection.startRow || reference.row > selection.endRow
      || reference.column < selection.startColumn || reference.column > selection.endColumn) return reference;
    return { ...reference, row: reference.row + rowDelta, column: reference.column + columnDelta };
  };

  for (const owner of workbook.getSheets()) {
    owner.cells.forEach((cell, row, column) => {
      if (!cell.formula) return;
      if (owner.id === targetSheet.id
        && (insideCell(selection, row, column) || insideCell(destination, row, column))) return;
      const next = transformFormula(cell.formula, (ast) => mapAstReferences(ast, mapper(owner)));
      if (next !== cell.formula) owner.cells.set(row, column, { ...cell, formula: next });
    });
  }
  for (const [name, formula] of Object.entries(workbook.definedNames)) {
    workbook.definedNames[name] = transformFormula(formula, (ast) => mapAstReferences(ast, mapper(targetSheet)));
  }
  for (const entry of workbook.definedNameModels) {
    if (entry.scope === 'sheet' && entry.sheetId !== targetSheet.id) continue;
    entry.formula = transformFormula(entry.formula, (ast) => mapAstReferences(ast, mapper(targetSheet)));
  }
}

function transformFormula(formula: string, transform: (ast: ReturnType<typeof parseFormula>) => ReturnType<typeof parseFormula>): string {
  if (!formula.trim().startsWith('=')) return formula;
  try {
    return formatFormula(transform(parseFormula(formula)));
  } catch {
    return formula;
  }
}

function shiftRangeRef(range: RangeRef, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): boolean {
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  if (direction === 1) {
    if (range[startKey] >= at) {
      range[startKey] += count;
      range[endKey] += count;
    } else if (range[endKey] >= at) {
      range[endKey] += count;
    }
    return range[endKey] >= range[startKey];
  }
  const end = at + count - 1;
  if (range[endKey] < at) return true;
  if (range[startKey] > end) {
    range[startKey] -= count;
    range[endKey] -= count;
    return true;
  }
  range[endKey] = Math.max(at - 1, range[endKey] - count);
  range[startKey] = Math.min(Math.max(at - 1, range[startKey]), range[endKey]);
  return range[endKey] >= range[startKey];
}

function shiftMerges(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.merges.length - 1; index >= 0; index--) {
    const merge = sheet.merges[index]!;
    const keep = shiftRangeRef(merge.range, axis, at, count, direction);
    if (!keep) {
      sheet.merges.splice(index, 1);
      continue;
    }
    if (axis === 'row') {
      if (direction === 1 && merge.anchor.row >= at) merge.anchor.row += count;
      if (direction === -1) {
        if (merge.anchor.row >= at && merge.anchor.row < at + count) merge.anchor.row = merge.range.startRow;
        else if (merge.anchor.row >= at + count) merge.anchor.row -= count;
      }
    } else {
      if (direction === 1 && merge.anchor.column >= at) merge.anchor.column += count;
      if (direction === -1) {
        if (merge.anchor.column >= at && merge.anchor.column < at + count) merge.anchor.column = merge.range.startColumn;
        else if (merge.anchor.column >= at + count) merge.anchor.column -= count;
      }
    }
  }
}

function shiftRuleRanges(rules: Array<{ ranges: RangeRef[] }>, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const rule of rules) {
    rule.ranges = rule.ranges.filter((range) => shiftRangeRef(range, axis, at, count, direction));
  }
}

function shiftFilter(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.filter) return;
  if (!shiftRangeRef(sheet.filter.range, axis, at, count, direction)) {
    sheet.filter = undefined;
    return;
  }
  if (axis === 'column') {
    const next: typeof sheet.filter.criteria = {};
    for (const [key, condition] of Object.entries(sheet.filter.criteria)) {
      const column = Number(key);
      const shifted = shiftIndex(column, at, count, direction);
      if (shifted == null) continue;
      next[shifted] = { ...condition, column: shifted };
    }
    sheet.filter.criteria = next;
  }
}

function shiftFreeze(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    if (direction === 1 && sheet.freeze.ySplit >= at) sheet.freeze.ySplit += count;
    if (direction === -1 && sheet.freeze.ySplit > at) sheet.freeze.ySplit = Math.max(0, sheet.freeze.ySplit - count);
    if (direction === 1 && sheet.freeze.startRow >= at) sheet.freeze.startRow += count;
    if (direction === -1 && sheet.freeze.startRow > at) sheet.freeze.startRow = Math.max(0, sheet.freeze.startRow - count);
    return;
  }
  if (direction === 1 && sheet.freeze.xSplit >= at) sheet.freeze.xSplit += count;
  if (direction === -1 && sheet.freeze.xSplit > at) sheet.freeze.xSplit = Math.max(0, sheet.freeze.xSplit - count);
  if (direction === 1 && sheet.freeze.startColumn >= at) sheet.freeze.startColumn += count;
  if (direction === -1 && sheet.freeze.startColumn > at) sheet.freeze.startColumn = Math.max(0, sheet.freeze.startColumn - count);
}

function shiftHiddenAndSizes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    remapIndexSet(sheet.hiddenRows, at, count, direction);
    remapSizeMap(sheet.rowHeights, at, count, direction);
    return;
  }
  remapIndexSet(sheet.hiddenColumns, at, count, direction);
  remapSizeMap(sheet.columnWidths, at, count, direction);
}

function remapIndexSet(set: Set<number>, at: number, count: number, direction: 1 | -1): void {
  const next = new Set<number>();
  for (const value of set) {
    const shifted = shiftIndex(value, at, count, direction);
    if (shifted != null) next.add(shifted);
  }
  set.clear();
  for (const value of next) set.add(value);
}

function remapSizeMap(map: Record<number, number>, at: number, count: number, direction: 1 | -1): void {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const shifted = shiftIndex(Number(key), at, count, direction);
    if (shifted != null) next[shifted] = value;
  }
  for (const key of Object.keys(map)) delete map[Number(key)];
  Object.assign(map, next);
}

function shiftIndex(value: number, at: number, count: number, direction: 1 | -1): number | null {
  if (direction === 1) return value >= at ? value + count : value;
  if (value < at) return value;
  if (value < at + count) return null;
  return value - count;
}

function shiftSparklines(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.sparklines.length - 1; index >= 0; index--) {
    const sparkline = sheet.sparklines[index]!;
    shiftRangeRef(sparkline.sourceRange, axis, at, count, direction);
    const position = axis === 'row' ? sparkline.anchor.row : sparkline.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) {
      sheet.sparklines.splice(index, 1);
      continue;
    }
    if (axis === 'row') sparkline.anchor.row = shifted;
    else sparkline.anchor.column = shifted;
  }
}

function shiftPivots(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const pivot of sheet.pivots) {
    shiftRangeRef(pivot.sourceRange, axis, at, count, direction);
    if (pivot.dataSource?.kind === 'worksheet-range') shiftRangeRef(pivot.dataSource.range, axis, at, count, direction);
    if (pivot.dataSource?.kind === 'worksheet-ranges') {
      for (const range of pivot.dataSource.ranges) shiftRangeRef(range, axis, at, count, direction);
    }
    if (pivot.targetAnchor) {
      const position = axis === 'row' ? pivot.targetAnchor.row : pivot.targetAnchor.column;
      const shifted = shiftIndex(position, at, count, direction);
      if (shifted != null) {
        if (axis === 'row') pivot.targetAnchor.row = shifted;
        else pivot.targetAnchor.column = shifted;
      }
    }
  }
}

function shiftChartPayloads(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind !== 'chart') continue;
    for (const range of payload.sourceRanges) shiftRangeRef(range, axis, at, count, direction);
    if (payload.categoryRange) shiftRangeRef(payload.categoryRange, axis, at, count, direction);
    for (const series of payload.series ?? []) shiftRangeRef(series.range, axis, at, count, direction);
  }
}

function shiftDrawings(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.drawings.length - 1; index >= 0; index--) {
    const drawing = sheet.drawings[index]!;
    shiftDrawingAnchor(drawing, axis, at, count, direction);
  }
}

function shiftDrawingAnchor(drawing: DrawingObject, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (drawing.anchor.kind === 'absolute') return;
  const start = axis === 'row' ? drawing.anchor.row : drawing.anchor.column;
  if (start != null) {
    const shifted = shiftIndex(start, at, count, direction);
    if (shifted == null) {
      drawing.anchor.row = axis === 'row' ? at : drawing.anchor.row;
      drawing.anchor.column = axis === 'column' ? at : drawing.anchor.column;
    } else if (axis === 'row') drawing.anchor.row = shifted;
    else drawing.anchor.column = shifted;
  }
  const end = axis === 'row' ? drawing.anchor.endRow : drawing.anchor.endColumn;
  if (end != null) {
    const shifted = shiftIndex(end, at, count, direction);
    if (shifted != null) {
      if (axis === 'row') drawing.anchor.endRow = shifted;
      else drawing.anchor.endColumn = shifted;
    }
  }
}

function shiftSheetTables(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const kept = sheet.sheetTables.filter((table) => shiftRangeRef(table.range, axis, at, count, direction));
  sheet.sheetTables.splice(0, sheet.sheetTables.length, ...kept);
}

function shiftNotes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next = new Map<string, CellNote>();
  for (const [key, note] of sheet.notes) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    const nextRow = axis === 'row' ? shifted : row;
    const nextColumn = axis === 'column' ? shifted : column;
    next.set(noteCellKey(nextRow, nextColumn), note);
  }
  sheet.notes.clear();
  for (const [key, note] of next) sheet.notes.set(key, note);
}

function shiftComments(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next: CommentThread[] = [];
  for (const thread of sheet.commentThreads) {
    const position = axis === 'row' ? thread.row : thread.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    next.push(axis === 'row' ? { ...thread, row: shifted } : { ...thread, column: shifted });
  }
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...next);
}

function shiftSpills(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next: SpillRange[] = [];
  for (const spill of sheet.spillRanges) {
    const keep = shiftRangeRef(spill.range, axis, at, count, direction);
    if (!keep) continue;
    const position = axis === 'row' ? spill.anchor.row : spill.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    if (axis === 'row') spill.anchor.row = shifted;
    else spill.anchor.column = shifted;
    next.push(spill);
  }
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...next);
}

function shiftProtection(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const kept = sheet.protectionRules.filter((rule) => {
    if (!rule.range) return true;
    return shiftRangeRef(rule.range, axis, at, count, direction);
  });
  sheet.protectionRules.splice(0, sheet.protectionRules.length, ...kept);
}

function shiftBanded(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.bandedRule) return;
  if (!shiftRangeRef(sheet.bandedRule.range, axis, at, count, direction)) sheet.bandedRule = undefined;
}

function shiftOutline(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.outline) return;
  const next: OutlineGroup[] = [];
  for (const group of sheet.outline.groups) {
    if (group.axis !== axis) {
      next.push(group);
      continue;
    }
    const range: RangeRef = axis === 'row'
      ? { sheetId: sheet.id, startRow: group.start, endRow: group.end, startColumn: 0, endColumn: 0 }
      : { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: group.start, endColumn: group.end };
    if (!shiftRangeRef(range, axis, at, count, direction)) continue;
    next.push(axis === 'row'
      ? { ...group, start: range.startRow, end: range.endRow }
      : { ...group, start: range.startColumn, end: range.endColumn });
  }
  sheet.outline.groups = next;
}

function rewriteFormulas(workbook: WorkbookModel, sheetId: string, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const targetSheet = workbook.getSheet(sheetId);
  const shift: StructuralShift = {
    axis,
    at,
    count,
    op: direction === 1 ? 'insert' : 'delete',
  };
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      if (!cell.formula) return;
      const formula = transformFormula(cell.formula, (ast) => remapAst(
        ast,
        shift,
        (reference) => referenceBelongsToSheet(reference, sheet, targetSheet),
      ));
      if (formula !== cell.formula) sheet.cells.set(row, column, { ...cell, formula });
    });
  }
  rewriteDefinedNames(workbook, targetSheet, shift);
}

function rewriteDefinedNames(workbook: WorkbookModel, targetSheet: WorksheetModel, shift: StructuralShift): void {
  for (const [name, value] of Object.entries(workbook.definedNames)) {
    workbook.definedNames[name] = transformFormula(value, (ast) => remapAst(
      ast,
      shift,
      (reference) => reference.sheetId === undefined
        || reference.sheetId.trim().toLocaleLowerCase() === targetSheet.id.toLocaleLowerCase()
        || reference.sheetId.trim().toLocaleLowerCase() === targetSheet.name.toLocaleLowerCase(),
    ));
  }
  for (const entry of workbook.definedNameModels) {
    if (entry.scope === 'sheet' && entry.sheetId !== targetSheet.id) continue;
    entry.formula = transformFormula(entry.formula, (ast) => remapAst(
      ast,
      shift,
      (reference) => reference.sheetId === undefined
        || reference.sheetId.trim().toLocaleLowerCase() === targetSheet.id.toLocaleLowerCase()
        || reference.sheetId.trim().toLocaleLowerCase() === targetSheet.name.toLocaleLowerCase(),
    ));
  }
}

function applyMoveRange(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  source: RangeRef,
  targetOrigin: { row: Row; column: Column },
): StructuralTransformResult {
  const normalizedSource = normalizeRange(source);
  const height = normalizedSource.endRow - normalizedSource.startRow + 1;
  const width = normalizedSource.endColumn - normalizedSource.startColumn + 1;
  const target: RangeRef = {
    sheetId: sheet.id,
    startRow: targetOrigin.row,
    endRow: targetOrigin.row + height - 1,
    startColumn: targetOrigin.column,
    endColumn: targetOrigin.column + width - 1,
  };
  if (normalizedSource.sheetId !== sheet.id) throw new Error('Move range source must belong to the target worksheet');
  if (target.startRow < 0 || target.startColumn < 0 || target.endRow >= sheet.rowCount || target.endColumn >= sheet.columnCount) {
    throw new Error('Move range target is outside worksheet bounds');
  }
  validateMoveMetadataPreservation(sheet, normalizedSource, target);

  const rowDelta = target.startRow - normalizedSource.startRow;
  const colDelta = target.startColumn - normalizedSource.startColumn;
  const extracted = sheet.cells.extractRegion(
    normalizedSource.startRow,
    normalizedSource.endRow,
    normalizedSource.startColumn,
    normalizedSource.endColumn,
  );
  // Moving a range replaces every destination coordinate, including cells
  // that were empty in the source. This prevents stale target values.
  for (let row = target.startRow; row <= target.endRow; row += 1) {
    for (let column = target.startColumn; column <= target.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (const item of extracted) {
    const cell = item.cell.formula
      ? { ...item.cell, formula: offsetFormulaText(item.cell.formula, rowDelta, colDelta) }
      : item.cell;
    sheet.cells.set(item.row + rowDelta, item.column + colDelta, structuredClone(cell));
  }

  const relocate = (range: RangeRef): void => {
    if (!rangeContains(normalizedSource, range)) return;
    range.startRow += rowDelta;
    range.endRow += rowDelta;
    range.startColumn += colDelta;
    range.endColumn += colDelta;
  };
  for (const merge of sheet.merges) {
    if (rangeContains(normalizedSource, merge.range)) {
      relocate(merge.range);
      merge.anchor.row += rowDelta;
      merge.anchor.column += colDelta;
    }
  }
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) relocate(range);
  }
  if (sheet.filter) relocate(sheet.filter.range);
  for (const table of sheet.sheetTables) relocate(table.range);
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind !== 'chart') continue;
    for (const range of payload.sourceRanges) relocate(range);
    if (payload.categoryRange) relocate(payload.categoryRange);
    for (const series of payload.series ?? []) relocate(series.range);
  }
  for (const pivot of sheet.pivots) {
    relocate(pivot.sourceRange);
    if (pivot.dataSource?.kind === 'worksheet-range') relocate(pivot.dataSource.range);
    if (pivot.dataSource?.kind === 'worksheet-ranges') for (const range of pivot.dataSource.ranges) relocate(range);
    if (pivot.targetAnchor && insideCell(normalizedSource, pivot.targetAnchor.row, pivot.targetAnchor.column)) {
      pivot.targetAnchor.row += rowDelta;
      pivot.targetAnchor.column += colDelta;
    }
  }
  for (const sparkline of sheet.sparklines) {
    relocate(sparkline.sourceRange);
    if (insideCell(normalizedSource, sparkline.anchor.row, sparkline.anchor.column)) {
      sparkline.anchor.row += rowDelta;
      sparkline.anchor.column += colDelta;
    }
  }
  for (const spill of sheet.spillRanges) {
    relocate(spill.range);
    if (insideCell(normalizedSource, spill.anchor.row, spill.anchor.column)) {
      spill.anchor.row += rowDelta;
      spill.anchor.column += colDelta;
    }
  }
  for (const rule of sheet.protectionRules) if (rule.range) relocate(rule.range);
  if (sheet.bandedRule) relocate(sheet.bandedRule.range);
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute' || drawing.anchor.row === undefined || drawing.anchor.column === undefined) continue;
    if (!insideCell(normalizedSource, drawing.anchor.row, drawing.anchor.column)) continue;
    drawing.anchor.row += rowDelta;
    drawing.anchor.column += colDelta;
    if (drawing.anchor.endRow !== undefined) drawing.anchor.endRow += rowDelta;
    if (drawing.anchor.endColumn !== undefined) drawing.anchor.endColumn += colDelta;
  }
  const nextNotes = new Map<string, CellNote>();
  for (const [key, note] of sheet.notes) {
    const [row, column] = key.split(':').map(Number);
    if (insideCell(normalizedSource, row!, column!)) nextNotes.set(noteCellKey(row! + rowDelta, column! + colDelta), note);
    else nextNotes.set(key, note);
  }
  sheet.notes.clear();
  for (const [key, note] of nextNotes) sheet.notes.set(key, note);
  for (const thread of sheet.commentThreads) {
    if (insideCell(normalizedSource, thread.row, thread.column)) {
      thread.row += rowDelta;
      thread.column += colDelta;
    }
  }
  rewriteReferencesForMovedRegion(workbook, sheet, normalizedSource, target, rowDelta, colDelta);
  return { removedCells: extracted };
}

function normalizeRange(range: RangeRef): RangeRef {
  return {
    sheetId: range.sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function insideCell(range: RangeRef, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
}

function validateMoveMetadataPreservation(sheet: WorksheetModel, source: RangeRef, target: RangeRef): void {
  const validateRange = (range: RangeRef, label: string): void => {
    if (range.sheetId !== sheet.id) return;
    if (rangesIntersect(range, source) && !rangeContains(source, range)) {
      throw new Error(`Cannot move range: ${label} partially intersects the source`);
    }
    if (rangesIntersect(range, target) && !rangeContains(source, range)) {
      throw new Error(`Cannot move range: ${label} would be overwritten at the target`);
    }
  };
  for (const merge of sheet.merges) validateRange(merge.range, `merge ${merge.range.sheetId}`);
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) validateRange(range, 'rule range');
  }
  if (sheet.filter) validateRange(sheet.filter.range, 'filter');
  for (const table of sheet.sheetTables) validateRange(table.range, `table ${table.id}`);
  for (const pivot of sheet.pivots) {
    validateRange(pivot.sourceRange, `pivot ${pivot.id} source`);
    if (pivot.dataSource?.kind === 'worksheet-range') validateRange(pivot.dataSource.range, `pivot ${pivot.id} source`);
    if (pivot.dataSource?.kind === 'worksheet-ranges') for (const range of pivot.dataSource.ranges) validateRange(range, `pivot ${pivot.id} source`);
  }
  for (const sparkline of sheet.sparklines) validateRange(sparkline.sourceRange, `sparkline ${sparkline.id} source`);
  for (const spill of sheet.spillRanges) validateRange(spill.range, 'spill range');
  for (const rule of sheet.protectionRules) if (rule.range) validateRange(rule.range, `protection ${rule.id}`);
  if (sheet.bandedRule) validateRange(sheet.bandedRule.range, 'banded rule');
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind !== 'chart') continue;
    for (const range of payload.sourceRanges) validateRange(range, 'chart source');
    if (payload.categoryRange) validateRange(payload.categoryRange, 'chart category source');
    for (const series of payload.series ?? []) validateRange(series.range, 'chart series source');
  }
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute' || drawing.anchor.row === undefined || drawing.anchor.column === undefined) continue;
    const anchor = {
      sheetId: sheet.id,
      startRow: drawing.anchor.row,
      endRow: drawing.anchor.endRow ?? drawing.anchor.row,
      startColumn: drawing.anchor.column,
      endColumn: drawing.anchor.endColumn ?? drawing.anchor.column,
    };
    validateRange(anchor, `drawing ${drawing.id} anchor`);
  }
  for (const [key] of sheet.notes) {
    const [row, column] = key.split(':').map(Number);
    if (insideCell(target, row!, column!) && !insideCell(source, row!, column!)) throw new Error(`Cannot move range: note ${key} would be overwritten`);
  }
  for (const thread of sheet.commentThreads) {
    if (insideCell(target, thread.row, thread.column) && !insideCell(source, thread.row, thread.column)) {
      throw new Error(`Cannot move range: comment ${thread.id} would be overwritten`);
    }
  }
}

export function ensureDrawing(sheet: WorksheetModel, kind: DrawingObject['kind'], payloadId: string, transform: DrawingObject['transform']): DrawingObject {
  const existing = sheet.drawings.find((item) => item.payloadId === payloadId && item.kind === kind);
  if (existing) {
    existing.transform = { ...transform };
    return existing;
  }
  const drawing: DrawingObject = {
    id: `drawing-${payloadId}`,
    sheetId: sheet.id,
    kind,
    anchor: { kind: 'absolute' },
    transform: { ...transform },
    zIndex: sheet.drawings.length,
    payloadId,
  };
  sheet.drawings.push(drawing);
  return drawing;
}

export function usedSheetTables(sheet: WorksheetModel): SheetTableModel[] {
  return sheet.sheetTables;
}

export function usedProtection(sheet: WorksheetModel): ProtectionRule[] {
  return sheet.protectionRules;
}
