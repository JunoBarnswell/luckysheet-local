import type { CellData, RangeRef, Row, Column, WorksheetModel } from './index';
import type { CellNote, DrawingObject, StructuralTransformParams, CommentThread, SheetTableModel, SpillRange, ProtectionRule } from './domain';
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

function applyAxis(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): StructuralTransformResult {
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
  shiftCharts(sheet, axis, at, count, direction);
  shiftDrawings(sheet, axis, at, count, direction);
  shiftSheetTables(sheet, axis, at, count, direction);
  shiftNotes(sheet, axis, at, count, direction);
  shiftComments(sheet, axis, at, count, direction);
  shiftSpills(sheet, axis, at, count, direction);
  shiftProtection(sheet, axis, at, count, direction);
  shiftBanded(sheet, axis, at, count, direction);
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
  rewriteReferencesForMovedRegion(workbook, sheet, selection, rowDelta, columnDelta, sourceCells);
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
  for (const chart of sheet.charts) {
    for (const range of chart.sourceRanges) shiftContainedRange(range, selection, rowDelta, columnDelta);
    if (chart.categoryRange) shiftContainedRange(chart.categoryRange, selection, rowDelta, columnDelta);
    for (const series of chart.series ?? []) shiftContainedRange(series.range, selection, rowDelta, columnDelta);
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
  rowDelta: number,
  columnDelta: number,
  movedCells: ReadonlyArray<{ row: number; column: number; cell: CellData }>,
): void {
  const movedDestinationKeys = new Set(
    movedCells.map((entry) => JSON.stringify([entry.row + rowDelta, entry.column + columnDelta])),
  );
  const mapper = (owner: WorksheetModel) => (reference: ParsedCellReference): ParsedCellReference => {
    if (!referenceBelongsToSheet(reference, owner, targetSheet)) return reference;
    if (reference.row < selection.startRow || reference.row > selection.endRow
      || reference.column < selection.startColumn || reference.column > selection.endColumn) return reference;
    const destinationKey = JSON.stringify([reference.row + rowDelta, reference.column + columnDelta]);
    // A reference to the cell that fell off the shifted edge has no surviving
    // destination. Leave it untouched rather than inventing a value.
    if (!movedDestinationKeys.has(destinationKey)) return reference;
    return { ...reference, row: reference.row + rowDelta, column: reference.column + columnDelta };
  };

  for (const owner of workbook.getSheets()) {
    owner.cells.forEach((cell, row, column) => {
      if (!cell.formula || (owner.id === targetSheet.id && movedDestinationKeys.has(JSON.stringify([row, column])))) return;
      const next = transformFormula(cell.formula, (ast) => mapAstReferences(ast, mapper(owner)));
      if (next !== cell.formula) owner.cells.set(row, column, { ...cell, formula: next });
    });
  }
  for (const [name, formula] of Object.entries(workbook.definedNames)) {
    workbook.definedNames[name] = transformFormula(formula, (ast) => mapAstReferences(ast, mapper(targetSheet)));
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

function shiftCharts(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const chart of sheet.charts) {
    for (const range of chart.sourceRanges) shiftRangeRef(range, axis, at, count, direction);
    if (chart.categoryRange) shiftRangeRef(chart.categoryRange, axis, at, count, direction);
    for (const series of chart.series ?? []) shiftRangeRef(series.range, axis, at, count, direction);
  }
}

function shiftDrawings(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.drawings.length - 1; index >= 0; index--) {
    const drawing = sheet.drawings[index]!;
    shiftDrawingAnchor(drawing, axis, at, count, direction);
    syncDrawingPayload(sheet, drawing);
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

function syncDrawingPayload(sheet: WorksheetModel, drawing: DrawingObject): void {
  const bounds = { x: drawing.transform.x, y: drawing.transform.y, width: drawing.transform.width, height: drawing.transform.height };
  if (drawing.kind === 'chart') {
    const chart = sheet.charts.find((item) => item.id === drawing.payloadId);
    if (chart) chart.bounds = bounds;
  } else if (drawing.kind === 'shape') {
    const shape = sheet.shapes.find((item) => item.id === drawing.payloadId);
    if (shape) shape.bounds = bounds;
  } else if (drawing.kind === 'image') {
    const image = sheet.images.find((item) => item.id === drawing.payloadId);
    if (image) image.bounds = bounds;
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
}

function applyMoveRange(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  source: RangeRef,
  targetOrigin: { row: Row; column: Column },
): StructuralTransformResult {
  const rowDelta = targetOrigin.row - source.startRow;
  const colDelta = targetOrigin.column - source.startColumn;
  const extracted = sheet.cells.extractRegion(source.startRow, source.endRow, source.startColumn, source.endColumn);
  for (const item of extracted) {
    if (item.cell.formula) {
      item.cell.formula = offsetFormulaText(item.cell.formula, rowDelta, colDelta);
    }
    sheet.cells.set(item.row + rowDelta, item.column + colDelta, item.cell);
  }
  const moved: RangeRef = {
    sheetId: source.sheetId,
    startRow: source.startRow,
    endRow: source.endRow,
    startColumn: source.startColumn,
    endColumn: source.endColumn,
  };
  const relocate = (range: RangeRef): void => {
    if (range.sheetId !== source.sheetId) return;
    if (range.startRow >= source.startRow && range.endRow <= source.endRow && range.startColumn >= source.startColumn && range.endColumn <= source.endColumn) {
      range.startRow += rowDelta;
      range.endRow += rowDelta;
      range.startColumn += colDelta;
      range.endColumn += colDelta;
    }
  };
  for (const merge of sheet.merges) relocate(merge.range);
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) relocate(range);
  }
  if (sheet.filter) relocate(sheet.filter.range);
  for (const table of sheet.sheetTables) relocate(table.range);
  void workbook;
  void moved;
  return { removedCells: extracted };
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
