import type { CellData, RangeRef, Row, Column, WorksheetModel } from './index';
import type { CellHyperlink, CellNote, DrawingObject, StructuralTransformParams, CommentThread, SheetTableModel, SpillRange, ProtectionRule, OutlineGroup, CellShiftSpec } from './domain';
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

export interface CellShiftPlan {
  spec: CellShiftSpec;
  selection: RangeRef;
  band: RangeRef;
  count: number;
  direction: 1 | -1;
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
      case 'cell-shift':
        if (!params.sourceRange || !params.operation || !params.axis) throw new Error('Cell shift requires range, operation and axis');
        return applyCellShift(workbook, sheet, {
          sheetId: params.sheetId,
          range: params.sourceRange,
          operation: params.operation,
          axis: params.axis,
        });
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
  workbook: WorkbookModel,
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
    if (pivot.target.sheetId === sheet.id) {
      const position = axis === 'row' ? pivot.target.anchor.row : pivot.target.anchor.column;
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
  for (const [key] of sheet.hyperlinks) {
    const [row, column] = key.split(':').map(Number);
    if (deleted(axis === 'row' ? row! : column!)) {
      throw new Error(`Cannot delete ${axis} ${at}: hyperlink at ${key} would be lost`);
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
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id && intersectsAxisRange(table.sourceRange, axis, at, count)) {
      throw new Error(`Cannot delete ${axis} ${at}: workbook table ${table.id} requires an explicit table operation`);
    }
  }
}

/**
 * Block bytes are immutable during a worksheet structural transform.  A row
 * or column insertion/deletion may therefore move a complete data region when
 * it is entirely before the operation, but it must reject an operation that
 * would add/remove rows or columns inside the region.  Silently expanding a
 * metadata range without rewriting the block would expose the wrong records.
 */
function validateDataRegionAxisPreservation(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  for (const region of sheet.dataRegions) {
    const start = axis === 'row' ? region.range.startRow : region.range.startColumn;
    const end = axis === 'row' ? region.range.endRow : region.range.endColumn;
    const operationEnd = at + count - 1;
    const shiftsEntireRegion = direction === 1 ? at <= start : operationEnd < start;
    const isAfterRegion = direction === 1 ? at > end : at > end;
    if (shiftsEntireRegion || isAfterRegion) continue;
    throw new Error(`Cannot structurally transform ${axis} ${at}: data region ${region.id} requires a data-block transaction`);
  }
}

function shiftDataRegionAxis(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  for (const region of sheet.dataRegions) {
    const start = axis === 'row' ? region.range.startRow : region.range.startColumn;
    const shouldShift = direction === 1 ? at <= start : at + count - 1 < start;
    if (!shouldShift) continue;
    const delta = direction * count;
    if (axis === 'row') {
      region.range.startRow += delta;
      region.range.endRow += delta;
      region.headerRow += delta;
    } else {
      region.range.startColumn += delta;
      region.range.endColumn += delta;
    }
    const source = workbook.dataModel.sources.get(region.sourceId);
    if (source?.sourceRange?.sheetId !== sheet.id) continue;
    if (axis === 'row') {
      source.sourceRange.startRow += delta;
      source.sourceRange.endRow += delta;
    } else {
      source.sourceRange.startColumn += delta;
      source.sourceRange.endColumn += delta;
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
  validateAxisMetadataPreservation(workbook, sheet, axis, at, count, direction);
  validateDataRegionAxisPreservation(sheet, axis, at, count, direction);
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

  // Only complete regions outside the structural edit are moved.  Any
  // intersecting edit was rejected above because it needs a block rewrite.
  shiftDataRegionAxis(workbook, sheet, axis, at, count, direction);

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
  shiftWorkbookTables(workbook, sheet.id, axis, at, count, direction);
  shiftNotes(sheet, axis, at, count, direction);
  shiftHyperlinks(sheet, axis, at, count, direction);
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

export function planCellShift(workbook: WorkbookModel, spec: CellShiftSpec): CellShiftPlan {
  const sheet = workbook.getSheet(spec.sheetId);
  if (spec.range.sheetId !== spec.sheetId) throw new Error('Cell shift range must belong to the target worksheet');
  const selection = normalizeRange(spec.range);
  const count = spec.axis === 'row'
    ? selection.endRow - selection.startRow + 1
    : selection.endColumn - selection.startColumn + 1;
  if (selection.startRow < 0 || selection.startColumn < 0
    || selection.endRow >= sheet.rowCount || selection.endColumn >= sheet.columnCount) {
    throw new Error('Cell shift selection is outside worksheet bounds');
  }
  // Cell insertion shifts the complete tail of the affected band.  Grow the
  // runtime extent before calculating that band so the same plan can be
  // consumed by the model, command runtime, and inverse operation.
  if (spec.operation === 'insert') {
    if (spec.axis === 'row') sheet.ensureCellExtent(sheet.rowCount + count - 1, 0);
    else sheet.ensureCellExtent(0, sheet.columnCount + count - 1);
  }
  const direction: 1 | -1 = spec.operation === 'insert' ? 1 : -1;
  const band: RangeRef = spec.axis === 'row'
    ? { sheetId: sheet.id, startRow: selection.startRow, endRow: sheet.rowCount - 1, startColumn: selection.startColumn, endColumn: selection.endColumn }
    : { sheetId: sheet.id, startRow: selection.startRow, endRow: selection.endRow, startColumn: selection.startColumn, endColumn: sheet.columnCount - 1 };
  validateCellShiftBounds(sheet, selection, band, spec.axis, spec.operation, count);
  validateDataRegionCellShift(sheet, band);
  return { spec: { ...spec, range: selection }, selection, band, count, direction };
}

function applyCellShift(workbook: WorkbookModel, sheet: WorksheetModel, spec: CellShiftSpec): StructuralTransformResult {
  const plan = planCellShift(workbook, spec);
  const sourceCells: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (insideCell(plan.band, row, column)) sourceCells.push({ row, column, cell: structuredClone(cell) });
  });
  const removedCells: Array<{ row: Row; column: Column; cell: CellData }> = [];
  for (let row = plan.band.startRow; row <= plan.band.endRow; row += 1) {
    for (let column = plan.band.startColumn; column <= plan.band.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (const entry of sourceCells) {
    const destination = mapCellShiftCoordinate(plan, entry.row, entry.column);
    if (!destination) {
      removedCells.push(entry);
      continue;
    }
    sheet.cells.set(destination.row, destination.column, entry.cell);
  }
  shiftCellBandMetadata(workbook, sheet, plan);
  rewriteFormulas(workbook, sheet.id, plan.spec.axis, plan.spec.axis === 'row' ? plan.selection.startRow : plan.selection.startColumn, plan.count, plan.direction);
  return { removedCells };
}

function validateCellShiftBounds(
  sheet: WorksheetModel,
  selection: RangeRef,
  band: RangeRef,
  axis: CellShiftSpec['axis'],
  operation: CellShiftSpec['operation'],
  count: number,
): void {
  sheet.cells.forEach((_cell, row, column) => {
    if (!insideCell(band, row, column)) return;
    const destination = mapCellShiftCoordinate({ spec: { sheetId: sheet.id, range: selection, operation, axis }, selection, band, count, direction: operation === 'insert' ? 1 : -1 }, row, column);
    if (destination && !insideCell(band, destination.row, destination.column)) throw new Error('Cell shift would move data outside worksheet bounds');
    if (!destination && operation === 'insert') throw new Error('Cell shift would discard data outside worksheet bounds');
  });
}

function validateDataRegionCellShift(sheet: WorksheetModel, band: RangeRef): void {
  for (const region of sheet.dataRegions) {
    if (rangesIntersect(region.range, band)) throw new Error(`Cannot shift cells across data region ${region.id}: requires a data-block transaction`);
  }
}

function mapCellShiftCoordinate(plan: CellShiftPlan, row: number, column: number): { row: number; column: number } | null {
  if (!insideCell(plan.band, row, column)) return null;
  const inSelection = insideCell(plan.selection, row, column);
  if (plan.spec.axis === 'row') {
    if (plan.spec.operation === 'delete' && inSelection) return null;
    if (row < plan.selection.startRow) return { row, column };
    return { row: plan.spec.operation === 'insert' ? row + plan.count : row - plan.count, column };
  }
  if (plan.spec.operation === 'delete' && inSelection) return null;
  if (column < plan.selection.startColumn) return { row, column };
  return { row, column: plan.spec.operation === 'insert' ? column + plan.count : column - plan.count };
}

function offsetFormulaText(formula: string, rowOffset: number, columnOffset: number): string {
  if (!formula.trim().startsWith('=')) return formula;
  try {
    return formatFormula(offsetAst(parseFormula(formula), rowOffset, columnOffset));
  } catch (error) {
    throw new Error(`Formula relocation failed: ${formula}`, { cause: error as Error });
  }
}

function rangeContains(outer: RangeRef, inner: RangeRef): boolean {
  return outer.sheetId === inner.sheetId
    && inner.startRow >= outer.startRow
    && inner.endRow <= outer.endRow
    && inner.startColumn >= outer.startColumn
    && inner.endColumn <= outer.endColumn;
}

function shiftCellBandMetadata(workbook: WorkbookModel, sheet: WorksheetModel, plan: CellShiftPlan): void {
  const shiftRange = (range: RangeRef): boolean => {
    if (range.sheetId !== sheet.id || !rangesIntersect(range, plan.band)) return true;
    if (plan.spec.axis === 'row'
      ? range.startColumn < plan.selection.startColumn || range.endColumn > plan.selection.endColumn
      : range.startRow < plan.selection.startRow || range.endRow > plan.selection.endRow) {
      throw new Error('Cell shift intersects metadata outside the affected band');
    }
    return shiftRangeRef(range, plan.spec.axis, plan.spec.axis === 'row' ? plan.selection.startRow : plan.selection.startColumn, plan.count, plan.direction);
  };
  const mapAnchor = (row: number, column: number): { row: number; column: number } | null => mapCellShiftCoordinate(plan, row, column);

  for (let index = sheet.merges.length - 1; index >= 0; index -= 1) {
    const merge = sheet.merges[index]!;
    if (!shiftRange(merge.range)) { sheet.merges.splice(index, 1); continue; }
    const anchor = mapAnchor(merge.anchor.row, merge.anchor.column);
    if (anchor) { merge.anchor.row = anchor.row; merge.anchor.column = anchor.column; }
  }
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    rule.ranges = rule.ranges.filter(shiftRange);
  }
  if (sheet.autoFilter) shiftRange(sheet.autoFilter.range);
  for (const table of sheet.sheetTables) {
    if (!shiftRange(table.range)) throw new Error(`Cell shift would remove sheet table ${table.id}`);
    if (table.autoFilter) shiftRange(table.autoFilter.range);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id && !shiftRange(table.sourceRange)) throw new Error(`Cell shift would remove workbook table ${table.id}`);
  }
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind === 'chart') {
      payload.sourceRanges = payload.sourceRanges.filter(shiftRange);
      if (payload.categoryRange) shiftRange(payload.categoryRange);
      payload.series = payload.series?.filter((series) => {
        const valid = shiftRange(series.range);
        if (series.xRange) shiftRange(series.xRange);
        if (series.yRange) shiftRange(series.yRange);
        if (series.errorBars?.plusRange) shiftRange(series.errorBars.plusRange);
        if (series.errorBars?.minusRange) shiftRange(series.errorBars.minusRange);
        return valid;
      });
    } else if (payload.kind === 'data-chart' && payload.source.kind === 'report-sheet') {
      if (!shiftRange(payload.source.range)) throw new Error('Cell shift would remove Data Chart report binding');
    }
  }
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range') shiftRange(pivot.source.range);
    if (pivot.source.kind === 'worksheet-ranges') pivot.source.ranges = pivot.source.ranges.filter((sourceRange) => shiftRange(sourceRange.range));
    if (pivot.target.sheetId === sheet.id) {
      const anchor = mapAnchor(pivot.target.anchor.row, pivot.target.anchor.column);
      if (anchor) { pivot.target.anchor.row = anchor.row; pivot.target.anchor.column = anchor.column; }
    }
  }
  for (const sparkline of sheet.sparklines) {
    shiftRange(sparkline.sourceRange);
    const anchor = mapAnchor(sparkline.anchor.row, sparkline.anchor.column);
    if (anchor) { sparkline.anchor.row = anchor.row; sparkline.anchor.column = anchor.column; }
  }
  for (const spill of sheet.spillRanges) {
    if (!shiftRange(spill.range)) throw new Error('Cell shift would remove a spill range');
    const anchor = mapAnchor(spill.anchor.row, spill.anchor.column);
    if (anchor) { spill.anchor.row = anchor.row; spill.anchor.column = anchor.column; }
  }
  for (const rule of sheet.protectionRules) if (rule.range) shiftRange(rule.range);
  if (sheet.bandedRule) shiftRange(sheet.bandedRule.range);
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute' || drawing.anchor.row == null || drawing.anchor.column == null) continue;
    const anchor = mapAnchor(drawing.anchor.row, drawing.anchor.column);
    if (!anchor) throw new Error(`Cell shift would remove drawing ${drawing.id}`);
    drawing.anchor.row = anchor.row;
    drawing.anchor.column = anchor.column;
    if (drawing.anchor.endRow != null) drawing.anchor.endRow = mapAnchor(drawing.anchor.endRow, drawing.anchor.endColumn ?? drawing.anchor.column)?.row ?? drawing.anchor.endRow;
    if (drawing.anchor.endColumn != null) drawing.anchor.endColumn = mapAnchor(drawing.anchor.row, drawing.anchor.endColumn)?.column ?? drawing.anchor.endColumn;
  }
  const remapMap = <T,>(source: Map<string, T>): Map<string, T> => {
    const next = new Map<string, T>();
    for (const [key, value] of source) {
      const parts = key.split(':');
      const row = Number(parts[0]);
      const column = Number(parts[1]);
      const anchor = mapAnchor(row, column);
      if (anchor) next.set(noteCellKey(anchor.row, anchor.column), value);
      else if (!insideCell(plan.band, row, column)) next.set(key, value);
      else throw new Error('Cell shift would remove anchored metadata');
    }
    return next;
  };
  const nextNotes = remapMap(sheet.notes);
  sheet.notes.clear();
  for (const [key, value] of nextNotes) sheet.notes.set(key, value);
  const nextHyperlinks = remapMap(sheet.hyperlinks);
  sheet.hyperlinks.clear();
  for (const [key, value] of nextHyperlinks) sheet.hyperlinks.set(key, value);
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.map((thread) => {
    const anchor = mapAnchor(thread.row, thread.column);
    if (anchor) return { ...thread, row: anchor.row, column: anchor.column };
    if (insideCell(plan.band, thread.row, thread.column)) throw new Error(`Cell shift would remove comment thread ${thread.id}`);
    return thread;
  }));
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
  for (const entry of workbook.definedNameModels) {
    if (entry.scope === 'sheet' && entry.sheetId !== targetSheet.id) continue;
    entry.formula = transformFormula(entry.formula, (ast) => mapAstReferences(ast, mapper(targetSheet)));
  }
}

function transformFormula(formula: string, transform: (ast: ReturnType<typeof parseFormula>) => ReturnType<typeof parseFormula>): string {
  const hasFormulaPrefix = formula.trim().startsWith('=');
  try {
    const formatted = formatFormula(transform(parseFormula(hasFormulaPrefix ? formula : `=${formula}`)));
    return hasFormulaPrefix ? formatted : formatted.replace(/^=/, '');
  } catch (error) {
    if (!hasFormulaPrefix) return formula;
    throw new Error(`Formula transformation failed: ${formula}`, { cause: error as Error });
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
  if (!sheet.autoFilter) return;
  if (!shiftRangeRef(sheet.autoFilter.range, axis, at, count, direction)) {
    sheet.autoFilter = undefined;
    return;
  }
  shiftAutoFilterSortState(sheet.autoFilter, axis, at, count, direction);
  if (axis === 'column') {
    const next: typeof sheet.autoFilter.columns = {};
    for (const [key, columnDefinition] of Object.entries(sheet.autoFilter.columns)) {
      const column = Number(key);
      const shifted = shiftIndex(column, at, count, direction);
      if (shifted == null) continue;
      next[shifted] = { ...columnDefinition, column: shifted };
    }
    sheet.autoFilter.columns = next;
  }
}

function shiftTableAutoFilter(table: SheetTableModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const autoFilter = table.autoFilter;
  if (!autoFilter) return;
  if (!shiftRangeRef(autoFilter.range, axis, at, count, direction)) {
    table.autoFilter = undefined;
    return;
  }
  shiftAutoFilterSortState(autoFilter, axis, at, count, direction);
  if (axis !== 'column') return;
  const next: typeof autoFilter.columns = {};
  for (const [key, column] of Object.entries(autoFilter.columns)) {
    const shifted = shiftIndex(Number(key), at, count, direction);
    if (shifted == null) continue;
    next[shifted] = { ...column, column: shifted };
  }
  autoFilter.columns = next;
}

function shiftAutoFilterSortState(autoFilter: NonNullable<WorksheetModel['autoFilter']>, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const sortState = autoFilter.sortState;
  if (!sortState) return;
  if (!shiftRangeRef(sortState.ref, axis, at, count, direction)) {
    delete autoFilter.sortState;
    return;
  }
  sortState.conditions = sortState.conditions.filter((condition) => shiftRangeRef(condition.ref, axis, at, count, direction));
  if (sortState.conditions.length === 0) delete autoFilter.sortState;
}

function shiftFreeze(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (sheet.pane.kind === 'none') return;
  if (axis === 'row') {
    if (sheet.pane.kind === 'frozen' && direction === 1 && sheet.pane.ySplit >= at) sheet.pane.ySplit += count;
    if (sheet.pane.kind === 'frozen' && direction === -1 && sheet.pane.ySplit > at) sheet.pane.ySplit = Math.max(0, sheet.pane.ySplit - count);
    if (direction === 1 && sheet.pane.startRow >= at) sheet.pane.startRow += count;
    if (direction === -1 && sheet.pane.startRow > at) sheet.pane.startRow = Math.max(0, sheet.pane.startRow - count);
    return;
  }
  if (sheet.pane.kind === 'frozen' && direction === 1 && sheet.pane.xSplit >= at) sheet.pane.xSplit += count;
  if (sheet.pane.kind === 'frozen' && direction === -1 && sheet.pane.xSplit > at) sheet.pane.xSplit = Math.max(0, sheet.pane.xSplit - count);
  if (direction === 1 && sheet.pane.startColumn >= at) sheet.pane.startColumn += count;
  if (direction === -1 && sheet.pane.startColumn > at) sheet.pane.startColumn = Math.max(0, sheet.pane.startColumn - count);
}

function shiftHiddenAndSizes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    remapIndexSet(sheet.hiddenRows, at, count, direction);
    remapSizeMap(sheet.rowHeightsPx, at, count, direction);
    return;
  }
  remapIndexSet(sheet.hiddenColumns, at, count, direction);
  remapSizeMap(sheet.columnWidthsPx, at, count, direction);
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
    if (pivot.source.kind === 'worksheet-range') shiftRangeRef(pivot.source.range, axis, at, count, direction);
    if (pivot.source.kind === 'worksheet-ranges') {
      for (const sourceRange of pivot.source.ranges) shiftRangeRef(sourceRange.range, axis, at, count, direction);
    }
    if (pivot.target.sheetId === sheet.id) {
      const position = axis === 'row' ? pivot.target.anchor.row : pivot.target.anchor.column;
      const shifted = shiftIndex(position, at, count, direction);
      if (shifted != null) {
        if (axis === 'row') pivot.target.anchor.row = shifted;
        else pivot.target.anchor.column = shifted;
      }
    }
  }
}

function shiftChartPayloads(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind === 'chart') {
      for (const range of payload.sourceRanges) shiftRangeRef(range, axis, at, count, direction);
      if (payload.categoryRange) shiftRangeRef(payload.categoryRange, axis, at, count, direction);
      for (const series of payload.series ?? []) {
        shiftRangeRef(series.range, axis, at, count, direction);
        if (series.xRange) shiftRangeRef(series.xRange, axis, at, count, direction);
        if (series.yRange) shiftRangeRef(series.yRange, axis, at, count, direction);
        if (series.errorBars?.plusRange) shiftRangeRef(series.errorBars.plusRange, axis, at, count, direction);
        if (series.errorBars?.minusRange) shiftRangeRef(series.errorBars.minusRange, axis, at, count, direction);
      }
    } else if (payload.kind === 'data-chart' && payload.source.kind === 'report-sheet') {
      shiftRangeRef(payload.source.range, axis, at, count, direction);
    }
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
  for (const table of sheet.sheetTables) shiftTableAutoFilter(table, axis, at, count, direction);
}

function shiftWorkbookTables(
  workbook: WorkbookModel,
  sheetId: string,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId !== sheetId) continue;
    if (!shiftRangeRef(table.sourceRange, axis, at, count, direction)) {
      throw new Error(`Workbook table ${table.id} lost its source range`);
    }
  }
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

function shiftHyperlinks(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next = new Map<string, CellHyperlink>();
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    const nextRow = axis === 'row' ? shifted : row;
    const nextColumn = axis === 'column' ? shifted : column;
    next.set(noteCellKey(nextRow, nextColumn), hyperlink);
  }
  sheet.hyperlinks.clear();
  for (const [key, hyperlink] of next) sheet.hyperlinks.set(key, hyperlink);
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
  for (const entry of workbook.definedNameModels) {
    if (entry.scope === 'sheet' && entry.sheetId !== targetSheet.id) continue;
    if (entry.anchor?.sheetId === targetSheet.id) {
      const position = shift.axis === 'row' ? entry.anchor.row : entry.anchor.column;
      if (shift.op === 'delete' && position >= shift.at && position < shift.at + shift.count) {
        throw new Error(`Defined name ${entry.name} anchor is removed by structural mutation`);
      }
      const delta = shift.op === 'insert'
        ? position >= shift.at ? shift.count : 0
        : position >= shift.at + shift.count ? -shift.count : 0;
      if (delta !== 0) {
        entry.anchor = shift.axis === 'row'
          ? { ...entry.anchor, row: entry.anchor.row + delta }
          : { ...entry.anchor, column: entry.anchor.column + delta };
      }
    }
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
  if (target.startRow < 0 || target.startColumn < 0) throw new Error('Move range target is outside worksheet bounds');
  sheet.ensureRangeExtent(target.startRow, target.endRow, target.startColumn, target.endColumn);
  validateMoveMetadataPreservation(workbook, sheet, normalizedSource, target);
  validateDataRegionMovePreservation(sheet, normalizedSource, target);

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
  relocateDataRegions(workbook, sheet, normalizedSource, rowDelta, colDelta);
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) relocate(range);
  }
  if (sheet.autoFilter) relocate(sheet.autoFilter.range);
  for (const table of sheet.sheetTables) {
    relocate(table.range);
    if (table.autoFilter) relocate(table.autoFilter.range);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) relocate(table.sourceRange);
  }
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind === 'chart') {
      for (const range of payload.sourceRanges) relocate(range);
      if (payload.categoryRange) relocate(payload.categoryRange);
      for (const series of payload.series ?? []) {
        relocate(series.range);
        if (series.xRange) relocate(series.xRange);
        if (series.yRange) relocate(series.yRange);
        if (series.errorBars?.plusRange) relocate(series.errorBars.plusRange);
        if (series.errorBars?.minusRange) relocate(series.errorBars.minusRange);
      }
    } else if (payload.kind === 'data-chart' && payload.source.kind === 'report-sheet') {
      relocate(payload.source.range);
    }
  }
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range') relocate(pivot.source.range);
    if (pivot.source.kind === 'worksheet-ranges') for (const sourceRange of pivot.source.ranges) relocate(sourceRange.range);
    if (pivot.target.sheetId === sheet.id && insideCell(normalizedSource, pivot.target.anchor.row, pivot.target.anchor.column)) {
      pivot.target.anchor.row += rowDelta;
      pivot.target.anchor.column += colDelta;
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
  const nextHyperlinks = new Map<string, CellHyperlink>();
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (insideCell(normalizedSource, row, column)) nextHyperlinks.set(noteCellKey(row + rowDelta, column + colDelta), hyperlink);
    else nextHyperlinks.set(key, hyperlink);
  }
  sheet.hyperlinks.clear();
  for (const [key, hyperlink] of nextHyperlinks) sheet.hyperlinks.set(key, hyperlink);
  for (const thread of sheet.commentThreads) {
    if (insideCell(normalizedSource, thread.row, thread.column)) {
      thread.row += rowDelta;
      thread.column += colDelta;
    }
  }
  rewriteReferencesForMovedRegion(workbook, sheet, normalizedSource, target, rowDelta, colDelta);
  return { removedCells: extracted };
}

function validateDataRegionMovePreservation(sheet: WorksheetModel, source: RangeRef, target: RangeRef): void {
  for (const region of sheet.dataRegions) {
    const sourceContains = rangeContains(source, region.range);
    const sourceIntersects = rangesIntersect(source, region.range);
    const targetContains = rangeContains(target, region.range);
    const targetIntersects = rangesIntersect(target, region.range);
    if (sourceIntersects && !sourceContains) {
      throw new Error(`Cannot move range: data region ${region.id} is partially intersected and requires a data-block transaction`);
    }
    if (targetIntersects && !(sourceContains && targetContains)) {
      throw new Error(`Cannot move range: data region ${region.id} would be overwritten and requires a data-block transaction`);
    }
    if (sourceContains && targetContains && sourceIntersects) {
      throw new Error(`Cannot move range: data region ${region.id} cannot be moved onto itself`);
    }
  }
}

function relocateDataRegions(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  source: RangeRef,
  rowDelta: number,
  columnDelta: number,
): void {
  for (const region of sheet.dataRegions) {
    if (!rangeContains(source, region.range)) continue;
    region.range.startRow += rowDelta;
    region.range.endRow += rowDelta;
    region.range.startColumn += columnDelta;
    region.range.endColumn += columnDelta;
    region.headerRow += rowDelta;
    const manifest = workbook.dataModel.sources.get(region.sourceId);
    if (!manifest?.sourceRange || manifest.sourceRange.sheetId !== sheet.id || !rangeContains(source, manifest.sourceRange)) continue;
    manifest.sourceRange.startRow += rowDelta;
    manifest.sourceRange.endRow += rowDelta;
    manifest.sourceRange.startColumn += columnDelta;
    manifest.sourceRange.endColumn += columnDelta;
  }
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

function validateMoveMetadataPreservation(workbook: WorkbookModel, sheet: WorksheetModel, source: RangeRef, target: RangeRef): void {
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
  if (sheet.autoFilter) validateRange(sheet.autoFilter.range, 'filter');
  for (const table of sheet.sheetTables) {
    validateRange(table.range, `table ${table.id}`);
    if (table.autoFilter) validateRange(table.autoFilter.range, `table ${table.id} autoFilter`);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) validateRange(table.sourceRange, `workbook table ${table.id}`);
  }
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range') validateRange(pivot.source.range, `pivot ${pivot.id} source`);
    if (pivot.source.kind === 'worksheet-ranges') for (const sourceRange of pivot.source.ranges) validateRange(sourceRange.range, `pivot ${pivot.id} source`);
  }
  for (const sparkline of sheet.sparklines) validateRange(sparkline.sourceRange, `sparkline ${sparkline.id} source`);
  for (const spill of sheet.spillRanges) validateRange(spill.range, 'spill range');
  for (const rule of sheet.protectionRules) if (rule.range) validateRange(rule.range, `protection ${rule.id}`);
  if (sheet.bandedRule) validateRange(sheet.bandedRule.range, 'banded rule');
  for (const payload of sheet.drawingPayloads.values()) {
    if (payload.kind === 'chart') {
      for (const range of payload.sourceRanges) validateRange(range, 'chart source');
      if (payload.categoryRange) validateRange(payload.categoryRange, 'chart category source');
      for (const series of payload.series ?? []) {
        validateRange(series.range, 'chart series source');
        if (series.xRange) validateRange(series.xRange, 'chart x source');
        if (series.yRange) validateRange(series.yRange, 'chart y source');
        if (series.errorBars?.plusRange) validateRange(series.errorBars.plusRange, 'chart error bar plus source');
        if (series.errorBars?.minusRange) validateRange(series.errorBars.minusRange, 'chart error bar minus source');
      }
    } else if (payload.kind === 'data-chart' && payload.source.kind === 'report-sheet') {
      validateRange(payload.source.range, 'data chart report binding');
    }
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
