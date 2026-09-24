import type { CellData, RangeRef, Row, Column } from './index';
import type { CellHyperlink, DrawingObject, StructuralTransformParams, SheetTableModel, SpillRange, ProtectionRule, OutlineGroup, CellShiftSpec } from './domain';
import type { WorkbookTableModel } from './data-model';
import type { DataSourceManifest } from './data-source';
import type { PrintDocumentSnapshot } from './workbook-state';
import { WorkbookModel, WorksheetModel, cellKey } from './index';
import {
  formatFormula,
  mapAstMovedReferences,
  mapAstStructuralReferences,
  parseFormula,
  transformReferenceInterval,
  type CellShiftReferenceTransform,
  type StructuralShift,
} from '@react-sheets/formula-engine';

export interface StructuralTransformResult {
  readonly kind: 'structural-transform';
  readonly removedCells: Array<{ row: Row; column: Column; cell: CellData }>;
  /** Sparse calculation inputs to clear and repopulate after this applied patch. */
  readonly clearInputRanges: readonly RangeRef[];
  readonly populateInputRanges: readonly RangeRef[];
  /** Formula owners rewritten outside the cell ranges above. */
  readonly rewrittenFormulaOwners: readonly StructuralReferenceOwnerAddress[];
}

export interface StructuralReferenceOwnerAddress {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

/** Formula-owner queries supplied by the canonical formula runtime. */
export interface StructuralReferenceOwnerIndex {
  getStructuralDependents(sheetId: string, axis: 'row' | 'column', at: number): readonly StructuralReferenceOwnerAddress[];
  getRangeDependents(sheetId: string, range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): readonly StructuralReferenceOwnerAddress[];
  getInvalidFormulaOwners(): readonly StructuralReferenceOwnerAddress[];
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
  static apply(workbook: WorkbookModel, params: StructuralTransformParams, referenceOwners: StructuralReferenceOwnerIndex): StructuralTransformResult {
    const sheet = workbook.getSheet(params.sheetId);
    switch (params.kind) {
      case 'insert-rows':
      case 'delete-rows':
      case 'insert-columns':
      case 'delete-columns': {
        if (params.at === undefined || params.count === undefined) {
          throw new Error('Structural axis mutation requires an explicit index and count');
        }
        const axis = params.kind.endsWith('rows') ? 'row' : 'column';
        return applyAxis(workbook, sheet, axis, params.at, params.count, params.kind.startsWith('insert-') ? 1 : -1, referenceOwners);
      }
      case 'move-range':
        if (!params.sourceRange || !params.targetOrigin) throw new Error('Move range requires a source range and target origin');
        return applyMoveRange(workbook, sheet, params.sourceRange, params.targetOrigin, referenceOwners);
      case 'cell-shift':
        if (!params.sourceRange || !params.operation || !params.axis) throw new Error('Cell shift requires range, operation and axis');
        return applyCellShift(workbook, sheet, {
          sheetId: params.sheetId,
          range: params.sourceRange,
          operation: params.operation,
          axis: params.axis,
        }, referenceOwners);
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
  if (!Number.isSafeInteger(at) || at < 0) throw new Error(`Structural ${axis} index must be a non-negative integer`);
  if (!Number.isSafeInteger(count) || count <= 0) throw new Error('Structural count must be a positive integer');
  const limit = axis === 'row' ? sheet.rowCount : sheet.columnCount;
  const maximum = axis === 'row' ? 1_048_576 : 16_384;
  const occupied = sheet.cells.occupiedRange(sheet.id);
  const occupiedEnd = axis === 'row' ? occupied.endRow : occupied.endColumn;
  if (sheet.cells.count() > 0 && (occupied.startRow < 0 || occupied.startColumn < 0
    || occupied.endRow >= 1_048_576 || occupied.endColumn >= 16_384)) {
    throw new Error('Existing worksheet cells exceed worksheet bounds');
  }
  if (direction === 1) {
    if (at > limit || limit + count > maximum) {
      throw new Error(`Cannot insert ${count} ${axis}(s) at ${at}: outside worksheet bounds`);
    }
    if (occupiedEnd >= at && occupiedEnd + count >= maximum) {
      throw new Error(`Structural insert would move existing cells outside the ${axis} limit`);
    }
    return;
  }
  if (direction === -1) {
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
  for (const { key, row, column } of sheet.review.noteEntries()) {
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
  for (const thread of sheet.review.threadEntries()) {
    if (deleted(axis === 'row' ? thread.row : thread.column)) {
      throw new Error(`Cannot delete ${axis} ${at}: comment thread ${thread.id} would be lost`);
    }
  }
  for (const merge of sheet.merges) {
    if (!intersectsAxisRange(merge.range, axis, at, count)) continue;
    const start = axis === 'row' ? merge.range.startRow : merge.range.startColumn;
    const end = axis === 'row' ? merge.range.endRow : merge.range.endColumn;
    if (start >= at && end < at + count) {
      throw new Error(`Cannot delete ${axis} ${at}: merge ${merge.range.startRow}:${merge.range.startColumn}-${merge.range.endRow}:${merge.range.endColumn} would be lost`);
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
  workbook: WorkbookModel,
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
  for (const source of workbook.dataModel.sources.values()) {
    const range = source.sourceRange;
    if (range?.sheetId !== sheet.id) continue;
    const start = axis === 'row' ? range.startRow : range.startColumn;
    const end = axis === 'row' ? range.endRow : range.endColumn;
    const operationEnd = at + count - 1;
    const shiftsEntireSource = direction === 1 ? at <= start : operationEnd < start;
    const isAfterSource = at > end;
    if (!shiftsEntireSource && !isAfterSource) {
      throw new Error(`Cannot structurally transform ${axis} ${at}: data source ${source.id} requires a data-block transaction`);
    }
  }
}

function shiftDataRegionAxis(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
  sources: Map<string, DataSourceManifest> = workbook.dataModel.sources,
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
  }
  for (const source of sources.values()) {
    if (source.sourceRange?.sheetId === sheet.id
      && !shiftRangeRef(source.sourceRange, axis, at, count, direction)) {
      throw new Error(`Structural mutation removes data source range ${source.id}`);
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
  referenceOwners: StructuralReferenceOwnerIndex,
): StructuralTransformResult {
  validateAxisBounds(sheet, axis, at, count, direction);
  validateAxisMetadataPreservation(workbook, sheet, axis, at, count, direction);
  validateDataRegionAxisPreservation(workbook, sheet, axis, at, count, direction);
  if (count <= 0) return { kind: 'structural-transform', removedCells: [], clearInputRanges: [], populateInputRanges: [], rewrittenFormulaOwners: [] };
  const calculationRanges = structuralAxisInputRanges(sheet, axis, at, count, direction);
  const shift: StructuralShift = {
    axis,
    at,
    count,
    op: direction === 1 ? 'insert' : 'delete',
  };
  const formulaRewrite = preflightFormulaRewrite(workbook, sheet, shift, referenceOwners);
  preflightAxisMetadata(workbook, sheet, axis, at, count, direction);
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
  for (const owner of workbook.getSheets()) {
    shiftRuleRanges(workbook, owner.conditionalFormats, axis, at, count, direction, sheet.id);
    shiftRuleRanges(workbook, owner.dataValidations, axis, at, count, direction, sheet.id);
    shiftSparklines(owner, axis, at, count, direction, sheet.id);
    shiftPivots(owner, axis, at, count, direction, sheet.id);
    shiftDrawingPayloadReferences(owner, axis, at, count, direction, sheet.id);
  }
  shiftFilter(sheet, axis, at, count, direction);
  shiftFreeze(sheet, axis, at, count, direction);
  shiftHiddenAndSizes(sheet, axis, at, count, direction);
  shiftDrawings(sheet, axis, at, count, direction);
  shiftSheetTables(sheet, axis, at, count, direction);
  shiftWorkbookTables(workbook, sheet.id, axis, at, count, direction);
  shiftReview(sheet, axis, at, count, direction);
  shiftHyperlinks(sheet, axis, at, count, direction);
  shiftHyperlinkTargets(workbook, workbook.getSheets(), sheet.id, axis, at, count, direction);
  shiftSpills(sheet, axis, at, count, direction);
  shiftProtection(sheet, axis, at, count, direction);
  shiftBanded(sheet, axis, at, count, direction);
  shiftOutline(sheet, axis, at, count, direction);
  shiftPrintDocumentAxis(workbook.printDocuments.get(sheet.id), sheet.id, axis, at, count, direction);
  const rewrittenFormulaOwners = applyFormulaRewritePlan(workbook, sheet.id, shift, undefined, formulaRewrite);
  return {
    kind: 'structural-transform',
    removedCells: removed,
    clearInputRanges: calculationRanges.clearInputRanges,
    populateInputRanges: calculationRanges.populateInputRanges,
    rewrittenFormulaOwners,
  };
}

function structuralAxisInputRanges(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): { clearInputRanges: RangeRef[]; populateInputRanges: RangeRef[] } {
  if (sheet.cells.count() === 0) return { clearInputRanges: [], populateInputRanges: [] };
  const occupied = sheet.cells.occupiedRange(sheet.id);
  const start = axis === 'row' ? occupied.startRow : occupied.startColumn;
  const end = axis === 'row' ? occupied.endRow : occupied.endColumn;
  if (end < at) return { clearInputRanges: [], populateInputRanges: [] };
  const orthogonalStart = axis === 'row' ? occupied.startColumn : occupied.startRow;
  const orthogonalEnd = axis === 'row' ? occupied.endColumn : occupied.endRow;
  const makeRange = (rangeStart: number, rangeEnd: number): RangeRef => axis === 'row'
    ? { sheetId: sheet.id, startRow: rangeStart, endRow: rangeEnd, startColumn: orthogonalStart, endColumn: orthogonalEnd }
    : { sheetId: sheet.id, startRow: orthogonalStart, endRow: orthogonalEnd, startColumn: rangeStart, endColumn: rangeEnd };
  const clearStart = Math.max(at, start);
  const clearInputRanges = [makeRange(clearStart, end)];
  if (direction === 1) {
    return {
      clearInputRanges,
      populateInputRanges: [makeRange(clearStart + count, end + count)],
    };
  }
  const movedSourceStart = Math.max(at + count, start);
  return {
    clearInputRanges,
    populateInputRanges: movedSourceStart > end
      ? []
      : [makeRange(movedSourceStart - count, end - count)],
  };
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
  const direction: 1 | -1 = spec.operation === 'insert' ? 1 : -1;
  const band: RangeRef = spec.axis === 'row'
    ? { sheetId: sheet.id, startRow: selection.startRow, endRow: sheet.rowCount - 1, startColumn: selection.startColumn, endColumn: selection.endColumn }
    : { sheetId: sheet.id, startRow: selection.startRow, endRow: selection.endRow, startColumn: selection.startColumn, endColumn: sheet.columnCount - 1 };
  validateCellShiftBounds(sheet, selection, band, spec.axis, spec.operation, count);
  validateDataRegionCellShift(workbook, sheet, band);
  return { spec: { ...spec, range: selection }, selection, band, count, direction };
}

function applyCellShift(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  spec: CellShiftSpec,
  referenceOwners: StructuralReferenceOwnerIndex,
): StructuralTransformResult {
  const plan = planCellShift(workbook, spec);
  const shift: StructuralShift = {
    axis: plan.spec.axis,
    at: plan.spec.axis === 'row' ? plan.selection.startRow : plan.selection.startColumn,
    count: plan.count,
    op: plan.direction === 1 ? 'insert' : 'delete',
  };
  const referenceShift: CellShiftReferenceTransform = {
    axis: plan.spec.axis,
    selection: plan.selection,
    direction: plan.direction,
  };
  const formulaRewrite = preflightFormulaRewrite(workbook, sheet, shift, referenceOwners, referenceShift);
  preflightCellShiftMetadata(workbook, sheet, plan);
  const sourceCells = sheet.cells.extractRegion(
    plan.band.startRow,
    plan.band.endRow,
    plan.band.startColumn,
    plan.band.endColumn,
  );
  const removedCells: Array<{ row: Row; column: Column; cell: CellData }> = [];
  for (const entry of sourceCells) {
    const destination = mapCellShiftCoordinate(plan, entry.row, entry.column);
    if (!destination) {
      removedCells.push(entry);
      continue;
    }
    sheet.cells.set(destination.row, destination.column, entry.cell);
  }
  shiftCellBandMetadata(workbook, sheet, plan);
  const rewrittenFormulaOwners = applyFormulaRewritePlan(workbook, sheet.id, shift, referenceShift, formulaRewrite, plan);
  return {
    kind: 'structural-transform',
    removedCells,
    clearInputRanges: [structuredClone(plan.band)],
    populateInputRanges: [structuredClone(plan.band)],
    rewrittenFormulaOwners,
  };
}

function validateCellShiftBounds(
  sheet: WorksheetModel,
  selection: RangeRef,
  band: RangeRef,
  axis: CellShiftSpec['axis'],
  operation: CellShiftSpec['operation'],
  count: number,
): void {
  const plan: CellShiftPlan = {
    spec: { sheetId: sheet.id, range: selection, operation, axis },
    selection,
    band,
    count,
    direction: operation === 'insert' ? 1 : -1,
  };
  sheet.cells.forEachInRange(band.startRow, band.endRow, band.startColumn, band.endColumn, (_cell, row, column) => {
    const destination = mapCellShiftCoordinate(plan, row, column);
    if (destination && !insideCell(band, destination.row, destination.column)) throw new Error('Cell shift would move data outside worksheet bounds');
    if (!destination && operation === 'insert') throw new Error('Cell shift would discard data outside worksheet bounds');
  });
}

function validateDataRegionCellShift(workbook: WorkbookModel, sheet: WorksheetModel, band: RangeRef): void {
  for (const region of sheet.dataRegions) {
    if (rangesIntersect(region.range, band)) throw new Error(`Cannot shift cells across data region ${region.id}: requires a data-block transaction`);
  }
  for (const table of sheet.sheetTables) {
    if (rangesIntersect(table.range, band)) throw new Error(`UNSUPPORTED_FEATURE: cell shift intersects table ${table.id}; use an explicit table operation`);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id && rangesIntersect(table.sourceRange, band)) {
      throw new Error(`UNSUPPORTED_FEATURE: cell shift intersects workbook table ${table.id}; use an explicit table operation`);
    }
  }
  for (const source of workbook.dataModel.sources.values()) {
    if (source.sourceRange?.sheetId === sheet.id && rangesIntersect(source.sourceRange, band)) {
      throw new Error(`UNSUPPORTED_FEATURE: cell shift intersects data source ${source.id}; use a data-block transaction`);
    }
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

function rangeContains(outer: RangeRef, inner: RangeRef): boolean {
  return outer.sheetId === inner.sheetId
    && inner.startRow >= outer.startRow
    && inner.endRow <= outer.endRow
    && inner.startColumn >= outer.startColumn
    && inner.endColumn <= outer.endColumn;
}

function shiftCellRangeReference(range: RangeRef, workbook: WorkbookModel, sheet: WorksheetModel, plan: CellShiftPlan): boolean {
  const shift: StructuralShift = {
    axis: plan.spec.axis,
    at: plan.spec.axis === 'row' ? plan.selection.startRow : plan.selection.startColumn,
    count: plan.count,
    op: plan.direction === 1 ? 'insert' : 'delete',
  };
  const mapped = mapAstStructuralReferences({
    type: 'range-reference',
    start: {
      type: 'cell-reference',
      reference: { sheetId: sheet.id, row: range.startRow, column: range.startColumn, absoluteRow: false, absoluteColumn: false },
      span: { start: 0, end: 0 },
    },
    end: {
      type: 'cell-reference',
      reference: { sheetId: sheet.id, row: range.endRow, column: range.endColumn, absoluteRow: false, absoluteColumn: false },
      span: { start: 0, end: 0 },
    },
    span: { start: 0, end: 0 },
  }, {
    shift,
    cellShift: { axis: plan.spec.axis, selection: plan.selection, direction: plan.direction },
    ownerSheetId: sheet.id,
    targetSheetId: sheet.id,
    targetSheetName: sheet.name,
    sheetOrder: workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name })),
  });
  if (mapped.type === 'invalid-reference') return false;
  if (mapped.type !== 'range-reference') throw new Error('STRUCTURAL_PATCH_INVARIANT: range transform changed its AST kind');
  range.startRow = mapped.start.reference.row;
  range.startColumn = mapped.start.reference.column;
  range.endRow = mapped.end.reference.row;
  range.endColumn = mapped.end.reference.column;
  return true;
}

function transformRuleFormulas(
  workbook: WorkbookModel,
  rules: Array<{
    id: string;
    sheetId: string;
    type?: string;
    formulaAnchor?: { sheetId: string; row: number; column: number };
    operator?: string;
    value1?: string | number;
    value2?: string | number;
    formula1?: string;
    formula2?: string;
    listSource?: { kind: 'values'; values: string[] } | { kind: 'range'; range: RangeRef } | { kind: 'formula'; formula: string };
  }>,
  targetSheet: WorksheetModel,
  shift: StructuralShift,
  cellShift: CellShiftReferenceTransform,
): void {
  const mapFormula = (formula: string, ownerSheetId: string): string => transformFormula(formula, (ast) => mapAstStructuralReferences(ast, {
    shift,
    cellShift,
    ownerSheetId,
    targetSheetId: targetSheet.id,
    targetSheetName: targetSheet.name,
    sheetOrder: workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name })),
  }));
  for (const rule of rules) {
    const ownerSheetId = rule.formulaAnchor?.sheetId ?? rule.sheetId;
    if (rule.operator === 'formula' && typeof rule.value1 === 'string') rule.value1 = mapFormula(rule.value1, ownerSheetId);
    else {
      if (typeof rule.value1 === 'string' && rule.value1.trim().startsWith('=')) rule.value1 = mapFormula(rule.value1, ownerSheetId);
      if (typeof rule.value2 === 'string' && rule.value2.trim().startsWith('=')) rule.value2 = mapFormula(rule.value2, ownerSheetId);
    }
    if (rule.formula1 && (rule.formula1.trim().startsWith('=') || rule.operator === 'formula' || rule.type === 'custom')) {
      rule.formula1 = mapFormula(rule.formula1, ownerSheetId);
    }
    if (rule.formula2 && (rule.formula2.trim().startsWith('=') || rule.type === 'custom')) {
      rule.formula2 = mapFormula(rule.formula2, ownerSheetId);
    }
    if (rule.listSource?.kind === 'formula') rule.listSource.formula = mapFormula(rule.listSource.formula, ownerSheetId);
  }
}

function cloneStructuralMetadataSheet(sheet: WorksheetModel): WorksheetModel {
  const staged = new WorksheetModel(sheet.id, sheet.name, sheet.rowCount, sheet.columnCount);
  staged.replaceDataRegions(sheet.dataRegions);
  staged.merges.push(...structuredClone(sheet.merges));
  staged.pivots.push(...structuredClone(sheet.pivots));
  staged.sparklines.push(...structuredClone(sheet.sparklines));
  staged.conditionalFormats.push(...structuredClone(sheet.conditionalFormats));
  staged.dataValidations.push(...structuredClone(sheet.dataValidations));
  staged.sheetTables.push(...structuredClone(sheet.sheetTables));
  staged.drawings.push(...structuredClone(sheet.drawings));
  for (const [key, payload] of sheet.drawingPayloads) staged.drawingPayloads.set(key, structuredClone(payload));
  for (const [key, hyperlink] of sheet.hyperlinks) staged.hyperlinks.set(key, structuredClone(hyperlink));
  staged.spillRanges.push(...structuredClone(sheet.spillRanges));
  staged.protectionRules.push(...structuredClone(sheet.protectionRules));
  staged.autoFilter = sheet.autoFilter ? structuredClone(sheet.autoFilter) : undefined;
  staged.bandedRule = sheet.bandedRule ? structuredClone(sheet.bandedRule) : undefined;
  staged.outline = sheet.outline ? structuredClone(sheet.outline) : undefined;
  staged.pane = structuredClone(sheet.pane);
  staged.defaultRowHeightPx = sheet.defaultRowHeightPx;
  staged.defaultColumnWidthPx = sheet.defaultColumnWidthPx;
  Object.assign(staged.rowHeightsPx, sheet.rowHeightsPx);
  Object.assign(staged.columnWidthsPx, sheet.columnWidthsPx);
  for (const row of sheet.hiddenRows) staged.hiddenRows.add(row);
  for (const column of sheet.hiddenColumns) staged.hiddenColumns.add(column);
  staged.review.replaceNotes(structuredClone(sheet.review.noteEntries()));
  staged.review.replaceThreads(structuredClone(sheet.review.threadEntries()));
  return staged;
}

function preflightCellShiftMetadata(workbook: WorkbookModel, sheet: WorksheetModel, plan: CellShiftPlan): void {
  const stagedSheets = workbook.getSheets().map(cloneStructuralMetadataSheet);
  const staged = stagedSheets.find((candidate) => candidate.id === sheet.id);
  if (!staged) throw new Error(`STRUCTURAL_PATCH_INVARIANT: worksheet ${sheet.id} is absent from metadata preflight`);
  const tables = [...workbook.dataModel.tables.values()].map((table) => structuredClone(table));
  const sources = new Map<string, DataSourceManifest>();
  for (const [id, source] of workbook.dataModel.sources) sources.set(id, structuredClone(source));
  const printDocument = workbook.printDocuments.get(sheet.id);
  shiftCellBandMetadata(workbook, staged, plan, tables, stagedSheets, sources,
    printDocument ? structuredClone(printDocument) : undefined);
}

function preflightAxisMetadata(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  const stagedSheets = workbook.getSheets().map(cloneStructuralMetadataSheet);
  const staged = stagedSheets.find((candidate) => candidate.id === sheet.id);
  if (!staged) throw new Error(`STRUCTURAL_PATCH_INVARIANT: worksheet ${sheet.id} is absent from metadata preflight`);
  const tables = [...workbook.dataModel.tables.values()].map((table) => structuredClone(table));
  const sources = new Map<string, DataSourceManifest>();
  for (const [id, source] of workbook.dataModel.sources) sources.set(id, structuredClone(source));
  shiftDataRegionAxis(workbook, staged, axis, at, count, direction, sources);
  shiftMerges(staged, axis, at, count, direction);
  for (const owner of stagedSheets) {
    shiftRuleRanges(workbook, owner.conditionalFormats, axis, at, count, direction, staged.id);
    shiftRuleRanges(workbook, owner.dataValidations, axis, at, count, direction, staged.id);
    shiftSparklines(owner, axis, at, count, direction, staged.id);
    shiftPivots(owner, axis, at, count, direction, staged.id);
    shiftDrawingPayloadReferences(owner, axis, at, count, direction, staged.id);
  }
  shiftFilter(staged, axis, at, count, direction);
  shiftFreeze(staged, axis, at, count, direction);
  shiftHiddenAndSizes(staged, axis, at, count, direction);
  shiftDrawings(staged, axis, at, count, direction);
  shiftSheetTables(staged, axis, at, count, direction);
  shiftWorkbookTables(workbook, staged.id, axis, at, count, direction, tables);
  shiftReview(staged, axis, at, count, direction);
  shiftHyperlinks(staged, axis, at, count, direction);
  shiftHyperlinkTargets(workbook, stagedSheets, staged.id, axis, at, count, direction);
  shiftSpills(staged, axis, at, count, direction);
  shiftProtection(staged, axis, at, count, direction);
  shiftBanded(staged, axis, at, count, direction);
  shiftOutline(staged, axis, at, count, direction);
  const printDocument = workbook.printDocuments.get(sheet.id);
  if (printDocument) shiftPrintDocumentAxis(structuredClone(printDocument), staged.id, axis, at, count, direction);
}

function shiftCellBandMetadata(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  plan: CellShiftPlan,
  workbookTables: Iterable<WorkbookTableModel> = workbook.dataModel.tables.values(),
  ownerSheets: readonly WorksheetModel[] = workbook.getSheets(),
  sources: Map<string, DataSourceManifest> = workbook.dataModel.sources,
  printDocument: PrintDocumentSnapshot | undefined = workbook.printDocuments.get(sheet.id),
): void {
  const shiftRange = (range: RangeRef): boolean => {
    if (range.sheetId !== sheet.id) return true;
    return shiftCellRangeReference(range, workbook, sheet, plan);
  };
  const shiftFilterReferences = (filter: NonNullable<WorksheetModel['autoFilter']>, label: string): void => {
    if (filter.range.sheetId !== sheet.id) return;
    if (plan.spec.axis === 'column' && rangesIntersect(filter.range, plan.band)) {
      throw new Error(`UNSUPPORTED_FEATURE: cell shift intersects ${label}; filter-column ownership requires an explicit filter operation`);
    }
    if (!shiftRange(filter.range)) throw new Error(`Cell shift removes ${label} range`);
    if (!filter.sortState) return;
    if (!shiftRange(filter.sortState.ref)) throw new Error(`Cell shift removes ${label} sort reference`);
    for (const condition of filter.sortState.conditions) {
      if (!shiftRange(condition.ref)) throw new Error(`Cell shift removes a ${label} sort condition`);
    }
  };
  const mapAnchor = (row: number, column: number): { row: number; column: number } | null =>
    insideCell(plan.band, row, column) ? mapCellShiftCoordinate(plan, row, column) : { row, column };

  for (let index = sheet.merges.length - 1; index >= 0; index -= 1) {
    const merge = sheet.merges[index]!;
    if (!shiftRange(merge.range)) throw new Error('Cell shift removes a merged range');
    const anchor = mapAnchor(merge.anchor.row, merge.anchor.column);
    if (!anchor) throw new Error(`Cell shift removes merge anchor at ${merge.anchor.row}:${merge.anchor.column}`);
    merge.anchor.row = anchor.row;
    merge.anchor.column = anchor.column;
  }
  for (const owner of ownerSheets) {
    for (const rule of [...owner.conditionalFormats, ...owner.dataValidations]) {
      rule.ranges = rule.ranges.filter(shiftRange);
      if (rule.ranges.length === 0) throw new Error(`Cell shift removes every range owned by rule ${rule.id}`);
      if (rule.listSource?.kind === 'range' && rule.listSource.range.sheetId === sheet.id
        && !shiftRange(rule.listSource.range)) {
        throw new Error(`Cell shift removes the list source for data validation ${rule.id}`);
      }
      if (rule.formulaAnchor?.sheetId === sheet.id
        && insideCell(plan.band, rule.formulaAnchor.row, rule.formulaAnchor.column)) {
        const nextAnchor = mapCellShiftCoordinate(plan, rule.formulaAnchor.row, rule.formulaAnchor.column);
        if (!nextAnchor) throw new Error(`Rule ${rule.id} formula anchor is removed by cell shift`);
        rule.formulaAnchor = { ...rule.formulaAnchor, row: nextAnchor.row, column: nextAnchor.column };
      }
    }
  }
  if (sheet.autoFilter) shiftFilterReferences(sheet.autoFilter, 'worksheet AutoFilter');
  for (const table of sheet.sheetTables) {
    if (!shiftRange(table.range)) throw new Error(`Cell shift would remove sheet table ${table.id}`);
    if (table.autoFilter) shiftFilterReferences(table.autoFilter, `AutoFilter for table ${table.id}`);
  }
  for (const table of workbookTables) {
    if (table.sourceRange?.sheetId === sheet.id && !shiftRange(table.sourceRange)) throw new Error(`Cell shift would remove workbook table ${table.id}`);
  }
  for (const source of sources.values()) {
    if (source.sourceRange?.sheetId === sheet.id) {
      if (rangesIntersect(source.sourceRange, plan.band)) {
        throw new Error(`UNSUPPORTED_FEATURE: cell shift intersects data source ${source.id}; use a data-block transaction`);
      }
      const previous = { ...source.sourceRange };
      if (!shiftRange(source.sourceRange)) throw new Error(`Cell shift removes data source range ${source.id}`);
      const previousHeight = previous.endRow - previous.startRow;
      const nextHeight = source.sourceRange.endRow - source.sourceRange.startRow;
      const previousWidth = previous.endColumn - previous.startColumn;
      const nextWidth = source.sourceRange.endColumn - source.sourceRange.startColumn;
      if (previousHeight !== nextHeight || previousWidth !== nextWidth) {
        throw new Error(`UNSUPPORTED_FEATURE: cell shift changes the physical extent of data source ${source.id}; use a data-block transaction`);
      }
    }
  }
  for (const owner of ownerSheets) for (const payload of owner.drawingPayloads.values()) {
    if (payload.kind === 'camera' || payload.kind === 'screenshot') {
      if (!shiftRange(payload.sourceRange)) throw new Error(`Cell shift would remove ${payload.kind} source range`);
    } else if (payload.kind === 'chart') {
      if (payload.source.kind === 'worksheet-ranges') {
        for (const range of payload.source.ranges) {
          if (!shiftRange(range)) throw new Error(`Cell shift removes a worksheet source range for chart ${payload.id}`);
        }
        if (payload.source.ranges.length === 0) throw new Error(`Chart ${payload.id} has no worksheet source ranges`);
      }
      else if (payload.source.kind === 'report-range' && !shiftRange(payload.source.range)) throw new Error('Cell shift would remove Chart report binding');
      if (payload.categoryRange && !shiftRange(payload.categoryRange)) throw new Error(`Cell shift removes chart category range ${payload.id}`);
      for (const series of payload.series ?? []) {
        if (!shiftRange(series.range)) throw new Error(`Cell shift removes chart series range ${series.id}`);
        for (const range of [series.xRange, series.yRange, series.sizeRange, series.categoryRange,
          series.stockRoles?.open, series.stockRoles?.high, series.stockRoles?.low, series.stockRoles?.close,
          series.stockRoles?.volume, series.dataLabels?.valuesFromCells, series.errorBars?.plusRange,
          series.errorBars?.minusRange]) {
          if (range && !shiftRange(range)) throw new Error(`Cell shift removes a chart data range for series ${series.id}`);
        }
      }
    } else if (payload.kind === 'form-control') {
      if (payload.cellLink?.sheetId === sheet.id) {
        const anchor = mapAnchor(payload.cellLink.row, payload.cellLink.column);
        if (!anchor) throw new Error(`Cell shift removes form-control cell link ${payload.cellLink.sheetId}`);
        payload.cellLink = { ...payload.cellLink, row: anchor.row, column: anchor.column };
      }
      if ('inputRange' in payload && !shiftRange(payload.inputRange)) {
        throw new Error('Cell shift removes form-control input range');
      }
    }
  }
  for (const owner of ownerSheets) for (const pivot of owner.pivots) {
    if (pivot.source.kind === 'worksheet-range' && !shiftRange(pivot.source.range)) throw new Error(`Cell shift removes pivot source range ${pivot.id}`);
    if (pivot.source.kind === 'worksheet-ranges') {
      for (const sourceRange of pivot.source.ranges) {
        if (!shiftRange(sourceRange.range)) throw new Error(`Cell shift removes a source range for pivot ${pivot.id}`);
      }
      if (pivot.source.ranges.length === 0) throw new Error(`Pivot ${pivot.id} has no source ranges`);
    }
    if (pivot.target.sheetId === sheet.id) {
      const anchor = mapAnchor(pivot.target.anchor.row, pivot.target.anchor.column);
      if (!anchor) throw new Error(`Cell shift removes pivot target anchor ${pivot.id}`);
      pivot.target.anchor.row = anchor.row;
      pivot.target.anchor.column = anchor.column;
    }
  }
  for (const owner of ownerSheets) for (const sparkline of owner.sparklines) {
    if (!shiftRange(sparkline.sourceRange)) throw new Error(`Cell shift removes sparkline source range ${sparkline.id}`);
    if (sparkline.sheetId === sheet.id) {
      const anchor = mapAnchor(sparkline.anchor.row, sparkline.anchor.column);
      if (!anchor) throw new Error(`Cell shift removes sparkline anchor ${sparkline.id}`);
      sparkline.anchor.row = anchor.row;
      sparkline.anchor.column = anchor.column;
    }
  }
  for (const spill of sheet.spillRanges) {
    if (!shiftRange(spill.range)) throw new Error('Cell shift would remove a spill range');
    const anchor = mapAnchor(spill.anchor.row, spill.anchor.column);
    if (!anchor) throw new Error('Cell shift removes a spill anchor');
    spill.anchor.row = anchor.row;
    spill.anchor.column = anchor.column;
  }
  for (const rule of sheet.protectionRules) if (rule.range && !shiftRange(rule.range)) throw new Error(`Cell shift removes protection range ${rule.id}`);
  if (sheet.bandedRule && !shiftRange(sheet.bandedRule.range)) throw new Error('Cell shift removes the banded range');
  for (const drawing of sheet.drawings) {
    if (drawing.anchor.kind === 'absolute' || drawing.anchor.row == null || drawing.anchor.column == null) continue;
    const anchor = mapAnchor(drawing.anchor.row, drawing.anchor.column);
    if (!anchor) throw new Error(`Cell shift would remove drawing ${drawing.id}`);
    const endAnchor = drawing.anchor.endRow != null || drawing.anchor.endColumn != null
      ? mapAnchor(drawing.anchor.endRow ?? drawing.anchor.row, drawing.anchor.endColumn ?? drawing.anchor.column)
      : undefined;
    if (endAnchor === null) throw new Error(`Cell shift would remove drawing extent ${drawing.id}`);
    drawing.anchor.row = anchor.row;
    drawing.anchor.column = anchor.column;
    if (endAnchor) {
      if (drawing.anchor.endRow != null) drawing.anchor.endRow = endAnchor.row;
      if (drawing.anchor.endColumn != null) drawing.anchor.endColumn = endAnchor.column;
    }
  }
  const remapMap = <T,>(source: Map<string, T>): Map<string, T> => {
    const next = new Map<string, T>();
    for (const [key, value] of source) {
      const parts = key.split(':');
      const row = Number(parts[0]);
      const column = Number(parts[1]);
      const anchor = mapAnchor(row, column);
      if (anchor) next.set(cellKey(anchor.row, anchor.column), value);
      else if (!insideCell(plan.band, row, column)) next.set(key, value);
      else throw new Error('Cell shift would remove anchored metadata');
    }
    return next;
  };
  sheet.review.validateRemapCoordinates((row, column) => {
    const anchor = mapAnchor(row, column);
    if (anchor) return anchor;
    if (!insideCell(plan.band, row, column)) return { row, column };
    return undefined;
  });
  sheet.review.remapCoordinates((row, column) => {
    const anchor = mapAnchor(row, column);
    if (anchor) return anchor;
    if (!insideCell(plan.band, row, column)) return { row, column };
    throw new Error('Cell shift would remove anchored review metadata');
  });
  const nextHyperlinks = remapMap(sheet.hyperlinks);
  sheet.hyperlinks.clear();
  for (const [key, value] of nextHyperlinks) sheet.hyperlinks.set(key, value);
  const shift: StructuralShift = {
    axis: plan.spec.axis,
    at: plan.spec.axis === 'row' ? plan.selection.startRow : plan.selection.startColumn,
    count: plan.count,
    op: plan.direction === 1 ? 'insert' : 'delete',
  };
  const cellShift = { axis: plan.spec.axis, selection: plan.selection, direction: plan.direction } satisfies CellShiftReferenceTransform;
  for (const owner of ownerSheets) {
    transformRuleFormulas(workbook, owner.conditionalFormats, sheet, shift, cellShift);
    transformRuleFormulas(workbook, owner.dataValidations, sheet, shift, cellShift);
    shiftCellBandHyperlinkTargets(workbook, owner, sheet, plan, shift, cellShift);
  }
  shiftPrintDocumentCellRanges(printDocument, sheet.id, workbook, sheet, plan);
}

function shiftPrintDocumentAxis(
  document: PrintDocumentSnapshot | undefined,
  targetSheetId: string,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  if (!document || document.sheetId !== targetSheetId) return;
  for (let index = document.printAreas.length - 1; index >= 0; index -= 1) {
    const area = document.printAreas[index]!;
    if (area.range.sheetId === targetSheetId && !shiftRangeRef(area.range, axis, at, count, direction)) {
      document.printAreas.splice(index, 1);
    }
  }
  const shift = { axis, at, count, op: direction === 1 ? 'insert' as const : 'delete' as const };
  const titleKey = axis === 'row' ? 'repeatRows' : 'repeatColumns';
  const titleSpan = document[titleKey];
  if (titleSpan) {
    const mapped = transformReferenceInterval(titleSpan.start, titleSpan.end, shift);
    if (mapped) document[titleKey] = { ...titleSpan, start: mapped.start, end: mapped.end };
    else delete document[titleKey];
  }
  const breakKey = axis === 'row' ? 'row' : 'column';
  for (let index = document.pageBreaks.length - 1; index >= 0; index -= 1) {
    const pageBreak = document.pageBreaks[index]!;
    const coordinate = pageBreak[breakKey];
    if (coordinate === undefined) continue;
    const mapped = shiftIndex(coordinate, at, count, direction);
    if (mapped === null) document.pageBreaks.splice(index, 1);
    else pageBreak[breakKey] = mapped;
  }
}

function shiftPrintDocumentCellRanges(
  document: PrintDocumentSnapshot | undefined,
  targetSheetId: string,
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  plan: CellShiftPlan,
): void {
  if (!document || document.sheetId !== targetSheetId) return;
  for (let index = document.printAreas.length - 1; index >= 0; index -= 1) {
    const area = document.printAreas[index]!;
    if (area.range.sheetId === targetSheetId && !shiftCellRangeReference(area.range, workbook, sheet, plan)) {
      document.printAreas.splice(index, 1);
    }
  }
}

function rewriteReferencesForMovedRegion(
  workbook: WorkbookModel,
  targetSheet: WorksheetModel,
  selection: RangeRef,
  destination: RangeRef,
  rowDelta: number,
  columnDelta: number,
  referenceOwners: StructuralReferenceOwnerIndex,
): MovedFormulaRewritePlan {
  const plan: MovedFormulaRewritePlan = { cells: [], names: [], rules: [], hyperlinks: [] };
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const transformMovedFormula = (formula: string, ownerSheetId: string): string => transformFormula(formula, (ast) => mapAstMovedReferences(ast, {
    selection,
    rowDelta,
    columnDelta,
    ownerSheetId,
    targetSheetId: targetSheet.id,
    targetSheetName: targetSheet.name,
    sheetOrder,
  }));

  const formulaOwners = new Map<string, StructuralReferenceOwnerAddress>();
  for (const owner of referenceOwners.getRangeDependents(targetSheet.id, selection)) {
    formulaOwners.set(structuralOwnerKey(owner), owner);
  }
  for (const owner of referenceOwners.getInvalidFormulaOwners()) {
    formulaOwners.set(structuralOwnerKey(owner), owner);
  }
  for (const formulaOwner of formulaOwners.values()) {
    const owner = workbook.getSheet(formulaOwner.sheetId);
    const cell = owner.cells.get(formulaOwner.row, formulaOwner.column);
    if (cell?.formula === undefined) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula owner ${formulaOwner.sheetId}!${formulaOwner.row}:${formulaOwner.column} is missing from the workbook`);
    }
    if (owner.id === targetSheet.id
      && (insideCell(selection, formulaOwner.row, formulaOwner.column) || insideCell(destination, formulaOwner.row, formulaOwner.column))) continue;
    const next = transformMovedFormula(cell.formula, owner.id);
    if (next !== cell.formula) plan.cells.push({ sheetId: owner.id, row: formulaOwner.row, column: formulaOwner.column, formula: next });
  }
  for (const owner of workbook.getSheets()) {
    for (const storedRule of [...owner.conditionalFormats, ...owner.dataValidations]) {
      const rule = storedRule as MoveFormulaRule;
      const ownerSheetId = rule.formulaAnchor?.sheetId ?? rule.sheetId;
      const addRuleFormula = (field: MoveRuleFormulaField, formula: string): void => {
        const next = transformMovedFormula(formula, ownerSheetId);
        if (next !== formula) plan.rules.push({ rule, field, formula: next });
      };
      if (rule.operator === 'formula' && typeof rule.value1 === 'string') addRuleFormula('value1', rule.value1);
      else {
        if (typeof rule.value1 === 'string' && rule.value1.trim().startsWith('=')) addRuleFormula('value1', rule.value1);
        if (typeof rule.value2 === 'string' && rule.value2.trim().startsWith('=')) addRuleFormula('value2', rule.value2);
      }
      if (rule.formula1 && (rule.formula1.trim().startsWith('=') || rule.operator === 'formula' || rule.type === 'custom')) {
        addRuleFormula('formula1', rule.formula1);
      }
      if (rule.formula2 && (rule.formula2.trim().startsWith('=') || rule.type === 'custom')) addRuleFormula('formula2', rule.formula2);
      if (rule.listSource?.kind === 'formula') addRuleFormula('listSource.formula', rule.listSource.formula);
    }
    for (const hyperlink of owner.hyperlinks.values()) {
      const target = hyperlink.target;
      if (target.kind !== 'sheet' || target.sheetId !== targetSheet.id) continue;
      const next = { ...target };
      if (next.row !== undefined && next.column !== undefined
        && insideCell(selection, next.row, next.column)) {
        next.row += rowDelta;
        next.column += columnDelta;
      }
      if (next.address !== undefined) next.address = transformMovedFormula(next.address, targetSheet.id);
      if (JSON.stringify(next) !== JSON.stringify(target)) plan.hyperlinks.push({ hyperlink, target: next });
    }
  }
  for (const entry of workbook.definedNameModels) {
    const ownerSheetId = entry.anchor?.sheetId ?? entry.sheetId ?? targetSheet.id;
    const formula = transformMovedFormula(entry.formula, ownerSheetId);
    let anchor = entry.anchor;
    if (entry.anchor?.sheetId === targetSheet.id
      && insideCell(selection, entry.anchor.row, entry.anchor.column)) {
      anchor = {
        ...entry.anchor,
        row: entry.anchor.row + rowDelta,
        column: entry.anchor.column + columnDelta,
      };
    }
    if (formula !== entry.formula || anchor !== entry.anchor) plan.names.push({ entry, formula, anchor });
  }
  return plan;
}

interface MoveFormulaRule {
  sheetId: string;
  formulaAnchor?: { sheetId: string; row: number; column: number };
  type?: string;
  operator?: string;
  value1?: string | number;
  value2?: string | number;
  formula1?: string;
  formula2?: string;
  listSource?: { kind: 'values'; values: string[] } | { kind: 'range'; range: RangeRef } | { kind: 'formula'; formula: string };
}

type MoveRuleFormulaField = 'value1' | 'value2' | 'formula1' | 'formula2' | 'listSource.formula';

interface MovedFormulaRewritePlan {
  cells: Array<{ sheetId: string; row: number; column: number; formula: string }>;
  names: Array<{
    entry: WorkbookModel['definedNameModels'][number];
    formula: string;
    anchor?: WorkbookModel['definedNameModels'][number]['anchor'];
  }>;
  rules: Array<{ rule: MoveFormulaRule; field: MoveRuleFormulaField; formula: string }>;
  hyperlinks: Array<{ hyperlink: CellHyperlink; target: CellHyperlink['target'] }>;
}

function applyMovedFormulaRewritePlan(workbook: WorkbookModel, plan: MovedFormulaRewritePlan): StructuralReferenceOwnerAddress[] {
  const rewrittenOwners: StructuralReferenceOwnerAddress[] = [];
  for (const change of plan.cells) {
    const sheet = workbook.getSheet(change.sheetId);
    const cell = sheet.cells.get(change.row, change.column);
    if (!cell) throw new Error(`STRUCTURAL_PATCH_INVARIANT: moved-range formula owner ${change.sheetId}!${change.row}:${change.column} disappeared`);
    sheet.cells.set(change.row, change.column, { ...cell, formula: change.formula });
    rewrittenOwners.push({ sheetId: change.sheetId, row: change.row, column: change.column });
  }
  for (const change of plan.names) {
    change.entry.formula = change.formula;
    change.entry.anchor = change.anchor;
  }
  for (const change of plan.rules) {
    if (change.field === 'listSource.formula') {
      if (change.rule.listSource?.kind !== 'formula') throw new Error('STRUCTURAL_PATCH_INVARIANT: validation formula source changed during move');
      change.rule.listSource.formula = change.formula;
    } else {
      switch (change.field) {
        case 'value1': change.rule.value1 = change.formula; break;
        case 'value2': change.rule.value2 = change.formula; break;
        case 'formula1': change.rule.formula1 = change.formula; break;
        case 'formula2': change.rule.formula2 = change.formula; break;
      }
    }
  }
  for (const change of plan.hyperlinks) change.hyperlink.target = change.target;
  return rewrittenOwners;
}

function transformFormula(formula: string, transform: (ast: ReturnType<typeof parseFormula>) => ReturnType<typeof parseFormula>): string {
  const hasFormulaPrefix = formula.trim().startsWith('=');
  try {
    const formatted = formatFormula(transform(parseFormula(hasFormulaPrefix ? formula : `=${formula}`)));
    return hasFormulaPrefix ? formatted : formatted.replace(/^=/, '');
  } catch (error) {
    if (error instanceof Error && /^(UNSUPPORTED_FEATURE|UNSUPPORTED_STRUCTURAL_REFERENCE|STRUCTURAL_PATCH_INVARIANT):/.test(error.message)) {
      throw error;
    }
    throw new Error(`Formula transformation failed: ${formula}`, { cause: error as Error });
  }
}

function shiftRangeRef(range: RangeRef, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): boolean {
  const shift: StructuralShift = { axis, at, count, op: direction === 1 ? 'insert' : 'delete' };
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  const interval = transformReferenceInterval(range[startKey], range[endKey], shift);
  if (!interval) return false;
  range[startKey] = interval.start;
  range[endKey] = interval.end;
  return true;
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

function shiftRuleRanges(workbook: WorkbookModel, rules: Array<{
  id: string;
  sheetId: string;
  ranges: RangeRef[];
  type?: string;
  formulaAnchor?: { sheetId: string; row: number; column: number };
  operator?: string;
  value1?: string | number;
  value2?: string | number;
  formula1?: string;
  formula2?: string;
  listSource?: { kind: 'values'; values: string[] } | { kind: 'range'; range: RangeRef } | { kind: 'formula'; formula: string };
}>, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1, sheetId: string): void {
  const shift: StructuralShift = { axis, at, count, op: direction === 1 ? 'insert' : 'delete' };
  for (const rule of rules) {
    const ownerSheetId = rule.formulaAnchor?.sheetId ?? rule.sheetId;
    if (rule.formulaAnchor?.sheetId === sheetId) {
      const shifted = shiftIndex(axis === 'row' ? rule.formulaAnchor.row : rule.formulaAnchor.column, at, count, direction);
      if (shifted === null) throw new Error(`Rule ${rule.id} formula anchor is removed by structural mutation`);
      rule.formulaAnchor = axis === 'row'
        ? { ...rule.formulaAnchor, row: shifted }
        : { ...rule.formulaAnchor, column: shifted };
    }
    rule.ranges = rule.ranges.filter((range) => range.sheetId !== sheetId || shiftRangeRef(range, axis, at, count, direction));
    if (rule.ranges.length === 0) throw new Error(`Rule ${rule.id} has no range after structural mutation`);
    if (rule.listSource?.kind === 'range' && rule.listSource.range.sheetId === sheetId
      && !shiftRangeRef(rule.listSource.range, axis, at, count, direction)) {
      throw new Error(`Data validation ${rule.id} list source is removed by structural mutation`);
    }
    const transformRuleFormula = (formula: string): string => transformFormula(formula, (ast) => mapAstStructuralReferences(ast, {
      shift,
      ownerSheetId,
      targetSheetId: sheetId,
      sheetOrder: workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name })),
    }));
    if (rule.operator === 'formula' && typeof rule.value1 === 'string') rule.value1 = transformRuleFormula(rule.value1);
    else {
      if (typeof rule.value1 === 'string' && rule.value1.trim().startsWith('=')) rule.value1 = transformRuleFormula(rule.value1);
      if (typeof rule.value2 === 'string' && rule.value2.trim().startsWith('=')) rule.value2 = transformRuleFormula(rule.value2);
    }
    if (rule.formula1 && (rule.formula1.trim().startsWith('=') || rule.operator === 'formula' || rule.type === 'custom')) rule.formula1 = transformRuleFormula(rule.formula1);
    if (rule.formula2 && (rule.formula2.trim().startsWith('=') || rule.type === 'custom')) rule.formula2 = transformRuleFormula(rule.formula2);
    if (rule.listSource?.kind === 'formula') rule.listSource.formula = transformRuleFormula(rule.listSource.formula);
  }
}

function shiftFilter(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.autoFilter) return;
  if (!shiftRangeRef(sheet.autoFilter.range, axis, at, count, direction)) {
    throw new Error('Structural mutation removes the worksheet AutoFilter range');
  }
  shiftAutoFilterSortState(sheet.autoFilter, axis, at, count, direction);
  if (axis === 'column') {
    const next: typeof sheet.autoFilter.columns = {};
    for (const [key, columnDefinition] of Object.entries(sheet.autoFilter.columns)) {
      const column = Number(key);
      const shifted = shiftIndex(column, at, count, direction);
      if (shifted == null) throw new Error(`Structural mutation removes AutoFilter column ${column}`);
      next[shifted] = { ...columnDefinition, column: shifted };
    }
    sheet.autoFilter.columns = next;
  }
}

function shiftTableAutoFilter(table: SheetTableModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const autoFilter = table.autoFilter;
  if (!autoFilter) return;
  if (!shiftRangeRef(autoFilter.range, axis, at, count, direction)) {
    throw new Error(`Structural mutation removes AutoFilter range for table ${table.id}`);
  }
  shiftAutoFilterSortState(autoFilter, axis, at, count, direction);
  if (axis !== 'column') return;
  const next: typeof autoFilter.columns = {};
  for (const [key, column] of Object.entries(autoFilter.columns)) {
    const shifted = shiftIndex(Number(key), at, count, direction);
    if (shifted == null) throw new Error(`Structural mutation removes AutoFilter column ${key} for table ${table.id}`);
    next[shifted] = { ...column, column: shifted };
  }
  autoFilter.columns = next;
}

function shiftAutoFilterSortState(autoFilter: NonNullable<WorksheetModel['autoFilter']>, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const sortState = autoFilter.sortState;
  if (!sortState) return;
  if (!shiftRangeRef(sortState.ref, axis, at, count, direction)) {
    throw new Error('Structural mutation removes AutoFilter sort reference');
  }
  for (const condition of sortState.conditions) {
    if (!shiftRangeRef(condition.ref, axis, at, count, direction)) {
      throw new Error('Structural mutation removes an AutoFilter sort condition');
    }
  }
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

function shiftSparklines(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
  targetSheetId: string = sheet.id,
): void {
  for (const sparkline of sheet.sparklines) {
    if (sparkline.sourceRange.sheetId === targetSheetId
      && !shiftRangeRef(sparkline.sourceRange, axis, at, count, direction)) {
      throw new Error(`Structural mutation removes sparkline source range ${sparkline.id}`);
    }
    if (sparkline.sheetId !== targetSheetId) continue;
    const position = axis === 'row' ? sparkline.anchor.row : sparkline.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) throw new Error(`Structural mutation removes sparkline anchor ${sparkline.id}`);
    if (axis === 'row') sparkline.anchor.row = shifted;
    else sparkline.anchor.column = shifted;
  }
}

function shiftPivots(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
  targetSheetId: string = sheet.id,
): void {
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range' && pivot.source.range.sheetId === targetSheetId
      && !shiftRangeRef(pivot.source.range, axis, at, count, direction)) {
      throw new Error(`Structural mutation removes pivot source range ${pivot.id}`);
    }
    if (pivot.source.kind === 'worksheet-ranges') {
      for (const sourceRange of pivot.source.ranges) {
        if (sourceRange.range.sheetId === targetSheetId
          && !shiftRangeRef(sourceRange.range, axis, at, count, direction)) {
          throw new Error(`Structural mutation removes a source range for pivot ${pivot.id}`);
        }
      }
      if (pivot.source.ranges.length === 0) throw new Error(`Pivot ${pivot.id} has no source ranges`);
    }
    if (pivot.target.sheetId === targetSheetId) {
      const position = axis === 'row' ? pivot.target.anchor.row : pivot.target.anchor.column;
      const shifted = shiftIndex(position, at, count, direction);
      if (shifted == null) throw new Error(`Structural mutation removes pivot target anchor ${pivot.id}`);
      if (axis === 'row') pivot.target.anchor.row = shifted;
      else pivot.target.anchor.column = shifted;
    }
  }
}

function shiftDrawingPayloadReferences(
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
  targetSheetId: string = sheet.id,
): void {
  for (const [payloadId, payload] of sheet.drawingPayloads) {
    if (payload.kind === 'camera' || payload.kind === 'screenshot') {
      if (payload.sourceRange.sheetId === targetSheetId
        && !shiftRangeRef(payload.sourceRange, axis, at, count, direction)) {
        throw new Error(`Structural mutation removes ${payload.kind} source range ${payloadId}`);
      }
    } else if (payload.kind === 'chart') {
      const requireRange = (range: RangeRef, label: string): void => {
        if (range.sheetId === targetSheetId && !shiftRangeRef(range, axis, at, count, direction)) {
          throw new Error(`Structural mutation removes ${label} for chart ${payloadId}`);
        }
      };
      if (payload.source.kind === 'worksheet-ranges') {
        for (const range of payload.source.ranges) requireRange(range, 'worksheet source range');
        if (payload.source.ranges.length === 0) throw new Error(`Chart ${payloadId} has no worksheet source ranges`);
      } else if (payload.source.kind === 'report-range') requireRange(payload.source.range, 'report binding');
      if (payload.categoryRange) requireRange(payload.categoryRange, 'category range');
      for (const series of payload.series ?? []) {
        requireRange(series.range, `series ${series.id} range`);
        for (const range of [series.xRange, series.yRange, series.sizeRange, series.categoryRange,
          series.stockRoles?.open, series.stockRoles?.high, series.stockRoles?.low, series.stockRoles?.close,
          series.stockRoles?.volume, series.dataLabels?.valuesFromCells, series.errorBars?.plusRange,
          series.errorBars?.minusRange]) {
          if (range) requireRange(range, `series ${series.id} data range`);
        }
      }
    } else if (payload.kind === 'form-control') {
      if ('cellLink' in payload && payload.cellLink?.sheetId === targetSheetId) {
        const row = axis === 'row' ? shiftIndex(payload.cellLink.row, at, count, direction) : payload.cellLink.row;
        const column = axis === 'column' ? shiftIndex(payload.cellLink.column, at, count, direction) : payload.cellLink.column;
        if (row === null || column === null) throw new Error(`Structural mutation removes form-control cell link ${payload.cellLink.sheetId}`);
        payload.cellLink = { ...payload.cellLink, row, column };
      }
      if ('inputRange' in payload && payload.inputRange.sheetId === targetSheetId
        && !shiftRangeRef(payload.inputRange, axis, at, count, direction)) {
        throw new Error('Structural mutation removes a form-control input range');
      }
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
    if (shifted == null) throw new Error(`Drawing ${drawing.id} anchor is removed by structural mutation`);
    if (axis === 'row') drawing.anchor.row = shifted;
    else drawing.anchor.column = shifted;
  }
  const end = axis === 'row' ? drawing.anchor.endRow : drawing.anchor.endColumn;
  if (end != null) {
    const shifted = shiftIndex(end, at, count, direction);
    if (shifted == null) throw new Error(`Drawing ${drawing.id} end anchor is removed by structural mutation`);
    if (axis === 'row') drawing.anchor.endRow = shifted;
    else drawing.anchor.endColumn = shifted;
  }
}

function shiftSheetTables(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const table of sheet.sheetTables) {
    if (!shiftRangeRef(table.range, axis, at, count, direction)) throw new Error(`Structural mutation removes sheet table ${table.id}`);
  }
  for (const table of sheet.sheetTables) shiftTableAutoFilter(table, axis, at, count, direction);
}

function shiftWorkbookTables(
  workbook: WorkbookModel,
  sheetId: string,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
  tables: Iterable<WorkbookTableModel> = workbook.dataModel.tables.values(),
): void {
  for (const table of tables) {
    if (table.sourceRange?.sheetId !== sheetId) continue;
    if (!shiftRangeRef(table.sourceRange, axis, at, count, direction)) {
      throw new Error(`Workbook table ${table.id} lost its source range`);
    }
  }
}

function shiftReview(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  sheet.review.validateRemapCoordinates((row, column) => {
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) return undefined;
    return { row: axis === 'row' ? shifted : row, column: axis === 'column' ? shifted : column };
  });
  sheet.review.remapCoordinates((row, column) => {
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) return undefined;
    return { row: axis === 'row' ? shifted : row, column: axis === 'column' ? shifted : column };
  });
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
    next.set(cellKey(nextRow, nextColumn), hyperlink);
  }
  sheet.hyperlinks.clear();
  for (const [key, hyperlink] of next) sheet.hyperlinks.set(key, hyperlink);
}

function shiftHyperlinkTargets(
  workbook: WorkbookModel,
  owners: readonly WorksheetModel[],
  targetSheetId: string,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  const targetSheet = workbook.getSheet(targetSheetId);
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  for (const owner of owners) for (const hyperlink of owner.hyperlinks.values()) {
    const target = hyperlink.target;
    if (target.kind !== 'sheet' || target.sheetId !== targetSheetId) continue;
    const next = { ...target };
    if (next.row !== undefined) {
      const row = axis === 'row' ? shiftIndex(next.row, at, count, direction) : next.row;
      if (row === null) throw new Error(`Structural mutation removes hyperlink ${hyperlink.id} target`);
      next.row = row;
    }
    if (next.column !== undefined) {
      const column = axis === 'column' ? shiftIndex(next.column, at, count, direction) : next.column;
      if (column === null) throw new Error(`Structural mutation removes hyperlink ${hyperlink.id} target`);
      next.column = column;
    }
    if (next.address !== undefined) {
      next.address = transformFormula(next.address, (ast) => mapAstStructuralReferences(ast, {
        shift: { axis, at, count, op: direction === 1 ? 'insert' : 'delete' },
        ownerSheetId: targetSheetId,
        targetSheetId,
        targetSheetName: targetSheet.name,
        sheetOrder,
      }));
    }
    hyperlink.target = next;
  }
}

function shiftCellBandHyperlinkTargets(
  workbook: WorkbookModel,
  owner: WorksheetModel,
  targetSheet: WorksheetModel,
  plan: CellShiftPlan,
  shift: StructuralShift,
  cellShift: CellShiftReferenceTransform,
): void {
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  for (const hyperlink of owner.hyperlinks.values()) {
    const target = hyperlink.target;
    if (target.kind !== 'sheet' || target.sheetId !== targetSheet.id) continue;
    const next = { ...target };
    if (next.row !== undefined && next.column !== undefined) {
      const mapped = insideCell(plan.band, next.row, next.column)
        ? mapCellShiftCoordinate(plan, next.row, next.column)
        : { row: next.row, column: next.column };
      if (!mapped) throw new Error(`Cell shift removes hyperlink ${hyperlink.id} target`);
      next.row = mapped.row;
      next.column = mapped.column;
    }
    if (next.address !== undefined) {
      next.address = transformFormula(next.address, (ast) => mapAstStructuralReferences(ast, {
        shift,
        cellShift,
        ownerSheetId: targetSheet.id,
        targetSheetId: targetSheet.id,
        targetSheetName: targetSheet.name,
        sheetOrder,
      }));
    }
    hyperlink.target = next;
  }
}

function shiftSpills(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const spill of sheet.spillRanges) {
    if (!shiftRangeRef(spill.range, axis, at, count, direction)) throw new Error('Structural mutation removes a spill range');
    const position = axis === 'row' ? spill.anchor.row : spill.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) throw new Error('Structural mutation removes a spill anchor');
    if (axis === 'row') spill.anchor.row = shifted;
    else spill.anchor.column = shifted;
  }
}

function shiftProtection(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const rule of sheet.protectionRules) {
    if (rule.range && !shiftRangeRef(rule.range, axis, at, count, direction)) {
      throw new Error(`Structural mutation removes protection range ${rule.id}`);
    }
  }
}

function shiftBanded(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.bandedRule) return;
  if (!shiftRangeRef(sheet.bandedRule.range, axis, at, count, direction)) {
    throw new Error('Structural mutation removes the banded range');
  }
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
    if (!shiftRangeRef(range, axis, at, count, direction)) {
      throw new Error(`Structural mutation removes outline group ${group.id}`);
    }
    next.push(axis === 'row'
      ? { ...group, start: range.startRow, end: range.endRow }
      : { ...group, start: range.startColumn, end: range.endColumn });
  }
  sheet.outline.groups = next;
}

interface FormulaRewritePlan {
  readonly cells: Array<{ sheetId: string; row: number; column: number; formula: string }>;
  readonly names: Array<{
    entry: WorkbookModel['definedNameModels'][number];
    formula: string;
    anchor?: WorkbookModel['definedNameModels'][number]['anchor'];
  }>;
}

function preflightFormulaRewrite(
  workbook: WorkbookModel,
  targetSheet: WorksheetModel,
  shift: StructuralShift,
  referenceOwners: StructuralReferenceOwnerIndex,
  cellShift?: CellShiftReferenceTransform,
): FormulaRewritePlan {
  const plan: FormulaRewritePlan = { cells: [], names: [] };
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const owners = new Map<string, StructuralReferenceOwnerAddress>();
  for (const owner of referenceOwners.getStructuralDependents(targetSheet.id, shift.axis, shift.at)) {
    owners.set(structuralOwnerKey(owner), owner);
  }
  for (const owner of referenceOwners.getInvalidFormulaOwners()) {
    owners.set(structuralOwnerKey(owner), owner);
  }
  for (const owner of owners.values()) {
    const sheet = workbook.getSheet(owner.sheetId);
    const cell = sheet.cells.get(owner.row, owner.column);
    if (cell?.formula === undefined) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula owner ${owner.sheetId}!${owner.row}:${owner.column} is missing from the workbook`);
    }
    const formula = transformFormula(cell.formula, (ast) => mapAstStructuralReferences(ast, {
      shift,
      cellShift,
      ownerSheetId: sheet.id,
      targetSheetId: targetSheet.id,
      targetSheetName: targetSheet.name,
      sheetOrder,
    }));
    if (formula !== cell.formula) plan.cells.push({ sheetId: sheet.id, row: owner.row, column: owner.column, formula });
  }
  for (const entry of workbook.definedNameModels) {
    const ownerSheetId = entry.anchor?.sheetId ?? entry.sheetId ?? targetSheet.id;
    let anchor = entry.anchor;
    if (anchor?.sheetId === targetSheet.id) {
      if (cellShift) {
        const moved = mapCellShiftCoordinateForOwner(cellShift, anchor.row, anchor.column);
        if (!moved) throw new Error(`Defined name ${entry.name} anchor is removed by structural mutation`);
        anchor = { ...anchor, row: moved.row, column: moved.column };
      } else {
        const position = shift.axis === 'row' ? anchor.row : anchor.column;
        const shifted = shiftIndex(position, shift.at, shift.count, shift.op === 'insert' ? 1 : -1);
        if (shifted === null) throw new Error(`Defined name ${entry.name} anchor is removed by structural mutation`);
        anchor = shift.axis === 'row'
          ? { ...anchor, row: shifted }
          : { ...anchor, column: shifted };
      }
    }
    const formula = transformFormula(entry.formula, (ast) => mapAstStructuralReferences(ast, {
      shift,
      cellShift,
      ownerSheetId,
      targetSheetId: targetSheet.id,
      targetSheetName: targetSheet.name,
      sheetOrder,
    }));
    if (formula !== entry.formula || anchor !== entry.anchor) plan.names.push({ entry, formula, anchor });
  }
  return plan;
}

function structuralOwnerKey(owner: StructuralReferenceOwnerAddress): string {
  return `${owner.sheetId}\u0000${owner.row}\u0000${owner.column}`;
}

function mapCellShiftCoordinateForOwner(
  transform: CellShiftReferenceTransform,
  row: number,
  column: number,
): { row: number; column: number } | null {
  const { selection } = transform;
  const inBand = transform.axis === 'row'
    ? row >= selection.startRow && column >= selection.startColumn && column <= selection.endColumn
    : column >= selection.startColumn && row >= selection.startRow && row <= selection.endRow;
  if (!inBand) return { row, column };
  if (transform.axis === 'row') {
    if (transform.direction < 0 && row <= selection.endRow) return null;
    return { row: row + transform.direction * (selection.endRow - selection.startRow + 1), column };
  }
  if (transform.direction < 0 && column <= selection.endColumn) return null;
  return { row, column: column + transform.direction * (selection.endColumn - selection.startColumn + 1) };
}

function applyFormulaRewritePlan(
  workbook: WorkbookModel,
  targetSheetId: string,
  shift: StructuralShift,
  cellShift: CellShiftReferenceTransform | undefined,
  plan: FormulaRewritePlan,
  cellShiftPlan?: CellShiftPlan,
): StructuralReferenceOwnerAddress[] {
  const rewrittenOwners: StructuralReferenceOwnerAddress[] = [];
  for (const change of plan.cells) {
    const sheet = workbook.getSheet(change.sheetId);
    let coordinate: { row: number; column: number } | null = { row: change.row, column: change.column };
    if (sheet.id === targetSheetId) {
      if (cellShift && cellShiftPlan) {
        coordinate = mapCellShiftCoordinateForOwner(cellShift, change.row, change.column);
      } else {
        const value = shift.axis === 'row' ? change.row : change.column;
        const shifted = shiftIndex(value, shift.at, shift.count, shift.op === 'insert' ? 1 : -1);
        coordinate = shifted === null ? null : shift.axis === 'row'
          ? { row: shifted, column: change.column }
          : { row: change.row, column: shifted };
      }
    }
    if (!coordinate) continue;
    const cell = sheet.cells.get(coordinate.row, coordinate.column);
    if (!cell) throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula owner ${sheet.id}!${coordinate.row}:${coordinate.column} was not preserved`);
    sheet.cells.set(coordinate.row, coordinate.column, { ...cell, formula: change.formula });
    rewrittenOwners.push({ sheetId: sheet.id, row: coordinate.row, column: coordinate.column });
  }
  for (const change of plan.names) {
    change.entry.formula = change.formula;
    change.entry.anchor = change.anchor;
  }
  return rewrittenOwners;
}

function applyMoveRange(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  source: RangeRef,
  targetOrigin: { row: Row; column: Column },
  referenceOwners: StructuralReferenceOwnerIndex,
): StructuralTransformResult {
  if (source.sheetId !== sheet.id) throw new Error('Move range source must belong to the target worksheet');
  if (![source.startRow, source.endRow, source.startColumn, source.endColumn, targetOrigin.row, targetOrigin.column]
    .every(Number.isSafeInteger)) throw new Error('Move range coordinates must be safe integers');
  const normalizedSource = normalizeRange(source);
  if (normalizedSource.startRow < 0 || normalizedSource.startColumn < 0
    || normalizedSource.endRow >= 1_048_576 || normalizedSource.endColumn >= 16_384) {
    throw new Error('Move range source is outside worksheet bounds');
  }
  const height = normalizedSource.endRow - normalizedSource.startRow + 1;
  const width = normalizedSource.endColumn - normalizedSource.startColumn + 1;
  const targetEndRow = targetOrigin.row + height - 1;
  const targetEndColumn = targetOrigin.column + width - 1;
  if (targetOrigin.row < 0 || targetOrigin.column < 0
    || !Number.isSafeInteger(targetEndRow) || !Number.isSafeInteger(targetEndColumn)
    || targetEndRow >= 1_048_576 || targetEndColumn >= 16_384) {
    throw new Error('Move range target is outside worksheet bounds');
  }
  const target: RangeRef = {
    sheetId: sheet.id,
    startRow: targetOrigin.row,
    endRow: targetEndRow,
    startColumn: targetOrigin.column,
    endColumn: targetEndColumn,
  };
  if (rangesIntersect(normalizedSource, target)) {
    throw new Error('Move range cannot overlap its source range');
  }
  validateMoveMetadataPreservation(workbook, sheet, normalizedSource, target);
  validateDataRegionMovePreservation(sheet, normalizedSource, target);

  const rowDelta = target.startRow - normalizedSource.startRow;
  const colDelta = target.startColumn - normalizedSource.startColumn;
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const cellsToMove = sheet.cells.getRegion(
    normalizedSource.startRow,
    normalizedSource.endRow,
    normalizedSource.startColumn,
    normalizedSource.endColumn,
  ).map((entry) => ({
    ...entry,
    cell: entry.cell.formula
      ? {
        ...entry.cell,
        formula: transformFormula(entry.cell.formula, (ast) => mapAstMovedReferences(ast, {
          selection: normalizedSource,
          rowDelta,
          columnDelta: colDelta,
          ownerSheetId: sheet.id,
          targetSheetId: sheet.id,
          targetSheetName: sheet.name,
          sheetOrder,
        })),
      }
      : entry.cell,
  }));
  const formulaRewrite = rewriteReferencesForMovedRegion(workbook, sheet, normalizedSource, target, rowDelta, colDelta, referenceOwners);
  sheet.ensureRangeExtent(target.startRow, target.endRow, target.startColumn, target.endColumn);
  sheet.cells.extractRegion(
    normalizedSource.startRow,
    normalizedSource.endRow,
    normalizedSource.startColumn,
    normalizedSource.endColumn,
  );
  // Moving a range replaces every destination coordinate, including cells
  // that were empty in the source. This prevents stale target values.
  const overwritten = sheet.cells.extractRegion(target.startRow, target.endRow, target.startColumn, target.endColumn);
  for (const item of cellsToMove) {
    sheet.cells.set(item.row + rowDelta, item.column + colDelta, item.cell);
  }

  const relocate = (range: RangeRef): void => {
    if (!rangeContains(normalizedSource, range)) return;
    range.startRow += rowDelta;
    range.endRow += rowDelta;
    range.startColumn += colDelta;
    range.endColumn += colDelta;
  };
  const relocateAutoFilter = (filter: NonNullable<WorksheetModel['autoFilter']>): void => {
    const rangeMoved = rangeContains(normalizedSource, filter.range);
    if (rangeMoved) relocate(filter.range);
    if (filter.sortState) {
      relocate(filter.sortState.ref);
      for (const condition of filter.sortState.conditions) relocate(condition.ref);
    }
    if (!rangeMoved || colDelta === 0) return;
    const columns: typeof filter.columns = {};
    for (const [key, definition] of Object.entries(filter.columns)) {
      const column = Number(key);
      const shifted = column >= normalizedSource.startColumn && column <= normalizedSource.endColumn
        ? column + colDelta
        : column;
      if (Object.prototype.hasOwnProperty.call(columns, String(shifted))) {
        throw new Error('STRUCTURAL_PATCH_INVARIANT: move collides AutoFilter column criteria');
      }
      columns[shifted] = { ...definition, column: shifted };
    }
    filter.columns = columns;
  };
  for (const merge of sheet.merges) {
    if (rangeContains(normalizedSource, merge.range)) {
      relocate(merge.range);
      merge.anchor.row += rowDelta;
      merge.anchor.column += colDelta;
    }
  }
  relocateDataRegions(sheet, normalizedSource, rowDelta, colDelta);
  const printDocument = workbook.printDocuments.get(sheet.id);
  for (const area of printDocument?.printAreas ?? []) relocate(area.range);
  for (const sourceManifest of workbook.dataModel.sources.values()) {
    if (sourceManifest.sourceRange?.sheetId === sheet.id) relocate(sourceManifest.sourceRange);
  }
  for (const owner of workbook.getSheets()) {
    for (const rule of [...owner.conditionalFormats, ...owner.dataValidations]) {
      for (const range of rule.ranges) relocate(range);
      if ('listSource' in rule && rule.listSource?.kind === 'range') relocate(rule.listSource.range);
      if (rule.formulaAnchor?.sheetId === sheet.id && insideCell(normalizedSource, rule.formulaAnchor.row, rule.formulaAnchor.column)) {
        rule.formulaAnchor = { ...rule.formulaAnchor, row: rule.formulaAnchor.row + rowDelta, column: rule.formulaAnchor.column + colDelta };
      }
    }
  }
  if (sheet.autoFilter) relocateAutoFilter(sheet.autoFilter);
  for (const table of sheet.sheetTables) {
    relocate(table.range);
    if (table.autoFilter) relocateAutoFilter(table.autoFilter);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) relocate(table.sourceRange);
  }
  for (const owner of workbook.getSheets()) for (const payload of owner.drawingPayloads.values()) {
    if (payload.kind === 'camera' || payload.kind === 'screenshot') {
      relocate(payload.sourceRange);
    } else if (payload.kind === 'chart') {
      if (payload.source.kind === 'worksheet-ranges') for (const range of payload.source.ranges) relocate(range);
      else if (payload.source.kind === 'report-range') relocate(payload.source.range);
      if (payload.categoryRange) relocate(payload.categoryRange);
      for (const series of payload.series ?? []) {
        relocate(series.range);
        if (series.xRange) relocate(series.xRange);
        if (series.yRange) relocate(series.yRange);
        if (series.sizeRange) relocate(series.sizeRange);
        if (series.categoryRange) relocate(series.categoryRange);
        if (series.stockRoles?.open) relocate(series.stockRoles.open);
        if (series.stockRoles?.high) relocate(series.stockRoles.high);
        if (series.stockRoles?.low) relocate(series.stockRoles.low);
        if (series.stockRoles?.close) relocate(series.stockRoles.close);
        if (series.stockRoles?.volume) relocate(series.stockRoles.volume);
        if (series.dataLabels?.valuesFromCells) relocate(series.dataLabels.valuesFromCells);
        if (series.errorBars?.plusRange) relocate(series.errorBars.plusRange);
        if (series.errorBars?.minusRange) relocate(series.errorBars.minusRange);
      }
    } else if (payload.kind === 'form-control') {
      if ('cellLink' in payload && payload.cellLink?.sheetId === sheet.id && insideCell(normalizedSource, payload.cellLink.row, payload.cellLink.column)) {
        payload.cellLink = {
          ...payload.cellLink,
          row: payload.cellLink.row + rowDelta,
          column: payload.cellLink.column + colDelta,
        };
      }
      if ('inputRange' in payload) relocate(payload.inputRange);
    }
  }
  for (const owner of workbook.getSheets()) for (const pivot of owner.pivots) {
    if (pivot.source.kind === 'worksheet-range') relocate(pivot.source.range);
    if (pivot.source.kind === 'worksheet-ranges') for (const sourceRange of pivot.source.ranges) relocate(sourceRange.range);
    if (pivot.target.sheetId === sheet.id && insideCell(normalizedSource, pivot.target.anchor.row, pivot.target.anchor.column)) {
      pivot.target.anchor.row += rowDelta;
      pivot.target.anchor.column += colDelta;
    }
  }
  for (const owner of workbook.getSheets()) for (const sparkline of owner.sparklines) {
    relocate(sparkline.sourceRange);
    if (sparkline.sheetId === sheet.id && insideCell(normalizedSource, sparkline.anchor.row, sparkline.anchor.column)) {
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
  sheet.review.remapCoordinates((row, column) => insideCell(normalizedSource, row, column)
    ? { row: row + rowDelta, column: column + colDelta }
    : { row, column });
  const nextHyperlinks = new Map<string, CellHyperlink>();
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (insideCell(normalizedSource, row, column)) nextHyperlinks.set(cellKey(row + rowDelta, column + colDelta), hyperlink);
    else nextHyperlinks.set(key, hyperlink);
  }
  sheet.hyperlinks.clear();
  for (const [key, hyperlink] of nextHyperlinks) sheet.hyperlinks.set(key, hyperlink);
  const rewrittenFormulaOwners = applyMovedFormulaRewritePlan(workbook, formulaRewrite);
  return {
    kind: 'structural-transform',
    removedCells: overwritten,
    clearInputRanges: [structuredClone(normalizedSource), structuredClone(target)],
    populateInputRanges: [structuredClone(normalizedSource), structuredClone(target)],
    rewrittenFormulaOwners,
  };
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
  const validateAutoFilter = (filter: NonNullable<WorksheetModel['autoFilter']>, label: string): void => {
    validateRange(filter.range, `${label} range`);
    const rangeMoves = rangeContains(source, filter.range);
    const columnDelta = target.startColumn - source.startColumn;
    const shiftedColumns = new Set<number>();
    for (const [key, definition] of Object.entries(filter.columns)) {
      const column = Number(key);
      if (!Number.isSafeInteger(definition.column) || key !== String(definition.column)
        || definition.column < filter.range.startColumn || definition.column > filter.range.endColumn
        || definition.column !== column) {
        throw new Error(`Cannot move range: ${label} has an invalid column identity`);
      }
      const shifted = rangeMoves && columnDelta !== 0
        && column >= source.startColumn && column <= source.endColumn
        ? column + columnDelta
        : column;
      if (shiftedColumns.has(shifted)) throw new Error(`Cannot move range: ${label} column criteria collide`);
      shiftedColumns.add(shifted);
    }
    if (!filter.sortState) return;
    validateRange(filter.sortState.ref, `${label} sort reference`);
    for (const condition of filter.sortState.conditions) validateRange(condition.ref, `${label} sort condition`);
  };
  for (const merge of sheet.merges) validateRange(merge.range, `merge ${merge.range.sheetId}`);
  for (const owner of workbook.getSheets()) {
    for (const rule of [...owner.conditionalFormats, ...owner.dataValidations]) {
      for (const range of rule.ranges) validateRange(range, `rule ${rule.id} range`);
      if ('listSource' in rule && rule.listSource?.kind === 'range') validateRange(rule.listSource.range, `validation ${rule.id} list source`);
    }
    for (const pivot of owner.pivots) {
      if (pivot.source.kind === 'worksheet-range') validateRange(pivot.source.range, `pivot ${pivot.id} source`);
      if (pivot.source.kind === 'worksheet-ranges') {
        for (const sourceRange of pivot.source.ranges) validateRange(sourceRange.range, `pivot ${pivot.id} source`);
      }
      if (pivot.target.sheetId === sheet.id && insideCell(target, pivot.target.anchor.row, pivot.target.anchor.column)
        && !insideCell(source, pivot.target.anchor.row, pivot.target.anchor.column)) {
        throw new Error(`Cannot move range: pivot ${pivot.id} output would be overwritten`);
      }
    }
    for (const sparkline of owner.sparklines) {
      validateRange(sparkline.sourceRange, `sparkline ${sparkline.id} source`);
      if (sparkline.sheetId === sheet.id && insideCell(target, sparkline.anchor.row, sparkline.anchor.column)
        && !insideCell(source, sparkline.anchor.row, sparkline.anchor.column)) {
        throw new Error(`Cannot move range: sparkline ${sparkline.id} would be overwritten`);
      }
    }
    for (const payload of owner.drawingPayloads.values()) {
      if (payload.kind === 'camera' || payload.kind === 'screenshot') {
        validateRange(payload.sourceRange, `${payload.kind} source`);
      } else if (payload.kind === 'chart') {
        if (payload.source.kind === 'worksheet-ranges') for (const range of payload.source.ranges) validateRange(range, 'chart source');
        else if (payload.source.kind === 'report-range') validateRange(payload.source.range, 'chart report binding');
        if (payload.categoryRange) validateRange(payload.categoryRange, 'chart category source');
        for (const series of payload.series ?? []) {
          validateRange(series.range, `chart series ${series.id ?? ''} source`);
          for (const range of [series.xRange, series.yRange, series.sizeRange, series.categoryRange,
            series.stockRoles?.open, series.stockRoles?.high, series.stockRoles?.low, series.stockRoles?.close,
            series.stockRoles?.volume, series.dataLabels?.valuesFromCells, series.errorBars?.plusRange,
            series.errorBars?.minusRange]) {
            if (range) validateRange(range, `chart series ${series.id ?? ''} source`);
          }
        }
      } else if (payload.kind === 'form-control') {
        if ('inputRange' in payload) validateRange(payload.inputRange, 'form-control input range');
        if ('cellLink' in payload && payload.cellLink?.sheetId === sheet.id
          && insideCell(target, payload.cellLink.row, payload.cellLink.column)
          && !insideCell(source, payload.cellLink.row, payload.cellLink.column)) {
          throw new Error('Cannot move range: form-control cell link would be overwritten');
        }
      }
    }
  }
  if (sheet.autoFilter) validateAutoFilter(sheet.autoFilter, 'filter');
  for (const table of sheet.sheetTables) {
    validateRange(table.range, `table ${table.id}`);
    if (table.autoFilter) validateAutoFilter(table.autoFilter, `table ${table.id} autoFilter`);
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) validateRange(table.sourceRange, `workbook table ${table.id}`);
  }
  for (const sourceManifest of workbook.dataModel.sources.values()) {
    if (sourceManifest.sourceRange?.sheetId === sheet.id) validateRange(sourceManifest.sourceRange, `data source ${sourceManifest.id}`);
  }
  for (const spill of sheet.spillRanges) validateRange(spill.range, 'spill range');
  for (const rule of sheet.protectionRules) if (rule.range) validateRange(rule.range, `protection ${rule.id}`);
  if (sheet.bandedRule) validateRange(sheet.bandedRule.range, 'banded rule');
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
  for (const { key, row, column } of sheet.review.noteEntries()) {
    if (insideCell(target, row!, column!) && !insideCell(source, row!, column!)) throw new Error(`Cannot move range: note ${key} would be overwritten`);
  }
  for (const thread of sheet.review.threadEntries()) {
    if (insideCell(target, thread.row, thread.column) && !insideCell(source, thread.row, thread.column)) {
      throw new Error(`Cannot move range: comment ${thread.id} would be overwritten`);
    }
  }
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (Number.isSafeInteger(row) && Number.isSafeInteger(column)
      && insideCell(target, row, column) && !insideCell(source, row, column)) {
      throw new Error(`Cannot move range: hyperlink ${key} would be overwritten at the target`);
    }
    if (hyperlink.target.kind === 'sheet' && hyperlink.target.sheetId === sheet.id
      && hyperlink.target.row !== undefined && hyperlink.target.column !== undefined
      && insideCell(target, hyperlink.target.row, hyperlink.target.column)
      && !insideCell(source, hyperlink.target.row, hyperlink.target.column)) {
      throw new Error(`Cannot move range: hyperlink ${key} target would be overwritten`);
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
