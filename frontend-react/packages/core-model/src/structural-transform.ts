import type { CellAddress, CellData, RangeRef, Row, Column } from './index';
import type { WorkbookCalculationContextEffect } from './calculation-context-effect';
import type { CellHyperlink, ChartTextFormulaField, DrawingObject, StructuralTransformParams, SheetTableModel, SpillRange, ProtectionRule, OutlineGroup, CellShiftSpec } from './domain';
import type { WorkbookTableModel } from './data-model';
import type { DataSourceManifest } from './data-source';
import type { PrintDocumentSnapshot } from './workbook-state';
import { mapReportSheetCoordinates } from './report-sheet-transform';
import { chartTextFormulaEntries, readChartTextFormula, writeChartTextFormula } from './chart-text-reference';
import { structuralRuleFormulaFields, type StructuralFormulaRule, type StructuralFormulaRuleField } from './structural-formula-owner';
import { WorkbookModel, WorksheetModel, cellKey, hasFormulaGroupMetadata, worksheetPaneValidationError, type WorksheetPane } from './index';
import {
  formatFormula,
  MAX_COLUMN_INDEX,
  MAX_ROW_INDEX,
  mapAstMovedReferences,
  mapAstStructuralReferences,
  parseFormula,
  rewriteFormulaTableReferences,
  ReferenceTransformDomain,
  type CellShiftReferenceTransform,
  type StructuralShift,
  type DefinedNameReferenceOwnerIdentity,
  type FormulaReferenceNode,
  type FormulaRuleReferenceFailure,
  type FormulaRuleReferenceFailureReason,
  type FormulaRuleReferenceOwnerIdentity,
  type FormulaDefinedName,
} from '@react-sheets/formula-engine';

function shiftIndex(position: number, at: number, count: number, direction: 1 | -1, axis: 'row' | 'column'): number | null {
  const maximum = axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  const mapped = ReferenceTransformDomain.mapPoint(position, at, count, direction, maximum);
  if (mapped.kind === 'out-of-bounds') {
    throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${axis} coordinate ${mapped.position} exceeds worksheet bounds`);
  }
  return mapped.kind === 'mapped' ? mapped.position : null;
}

export interface StructuralTransformResult {
  readonly kind: 'structural-transform';
  readonly removedCells: Array<{ row: Row; column: Column; cell: CellData }>;
  /** Sparse calculation inputs to clear and repopulate after this applied patch. */
  readonly clearInputRanges: readonly RangeRef[];
  readonly populateInputRanges: readonly RangeRef[];
  /** Formula-reference owners rewritten outside the cell ranges above. */
  readonly rewrittenFormulaOwners: readonly StructuralReferenceOwnerAddress[];
  /** Reversible pre/post states for formula owners rewritten by this transform. */
  readonly formulaOwnerDeltas?: readonly StructuralFormulaOwnerDelta[];
  /** Exact defined-name before/after values; present when name owners were fully indexed. */
  readonly definedNameOwnerDeltas?: readonly StructuralDefinedNameOwnerDelta[];
  /** A caller must rebuild calculation context when incremental owner updates cannot resolve the new identity. */
  readonly calculationContextEffect?: WorkbookCalculationContextEffect;
}

export interface StructuralReferenceOwnerAddress {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

export interface StructuralFormulaCellOwnerDelta {
  readonly kind: 'formula-cell';
  readonly beforeAddress: StructuralReferenceOwnerAddress;
  readonly afterAddress: StructuralReferenceOwnerAddress;
  readonly before: StructuralFormulaOwnerState;
  readonly after: StructuralFormulaOwnerState;
}

export interface StructuralFormulaRuleOwnerDelta {
  readonly kind: 'formula-rule';
  readonly sheetId: string;
  readonly ruleKind: 'conditional-format' | 'data-validation';
  readonly ruleId: string;
  readonly field: StructuralFormulaRuleField;
  readonly beforeFormula: string;
  readonly afterFormula: string;
  readonly beforeRanges: readonly RangeRef[];
  readonly afterRanges: readonly RangeRef[];
}

export type StructuralFormulaObjectOwnerDelta =
  | {
    readonly kind: 'formula-object';
    readonly ownerKind: 'chart-text';
    readonly sheetId: string;
    readonly payloadId: string;
    readonly field: ChartTextFormulaField;
    readonly beforeFormula: string;
    readonly afterFormula: string;
  }
  | {
    readonly kind: 'formula-object';
    readonly ownerKind: 'shape-property';
    readonly sheetId: string;
    readonly payloadId: string;
    readonly beforeFormula: string;
    readonly afterFormula: string;
  }
  | {
    readonly kind: 'formula-object';
    readonly ownerKind: 'table-sheet-column';
    readonly sheetId: string;
    readonly fieldId: string;
    readonly beforeFormula: string;
    readonly afterFormula: string;
  }
  | {
    readonly kind: 'formula-object';
    readonly ownerKind: 'data-view-field';
    readonly viewId: string;
    readonly fieldId: string;
    readonly beforeFormula: string;
    readonly afterFormula: string;
  }
  | {
    readonly kind: 'formula-object';
    readonly ownerKind: 'cell-style-template';
    readonly templateId: string;
    readonly field: 'formula1' | 'formula2' | 'listSource.formula';
    readonly beforeFormula: string;
    readonly afterFormula: string;
  };

export interface StructuralDefinedNameOwnerDelta {
  readonly owner: DefinedNameReferenceOwnerIdentity;
  readonly before: FormulaDefinedName;
  readonly after: FormulaDefinedName;
}

export type StructuralFormulaOwnerDelta = StructuralFormulaCellOwnerDelta | StructuralFormulaRuleOwnerDelta | StructuralFormulaObjectOwnerDelta;

export interface StructuralFormulaOwnerState {
  readonly formula: string | null;
  readonly sourceFormula: string | null;
  readonly barcodeFormula: string | null;
}

/** Formula-reference-owner queries supplied by the canonical formula runtime. */
export interface StructuralReferenceOwnerIndex {
  getStructuralDependents(sheetId: string, axis: 'row' | 'column', at: number): readonly StructuralReferenceOwnerAddress[];
  getRangeDependents(sheetId: string, range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): readonly StructuralReferenceOwnerAddress[];
  getInvalidFormulaOwners(): readonly StructuralReferenceOwnerAddress[];
  getStructuralDefinedNameDependents(sheetId: string, axis: 'row' | 'column', at: number): readonly DefinedNameReferenceOwnerIdentity[];
  getRangeDefinedNameDependents(sheetId: string, range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): readonly DefinedNameReferenceOwnerIdentity[];
  getDefinedNamesAnchoredInRange(sheetId: string, range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): readonly DefinedNameReferenceOwnerIdentity[];
  getDefinedNamesAnchoredAtOrAfter(sheetId: string, axis: 'row' | 'column', at: number): readonly DefinedNameReferenceOwnerIdentity[];
  getDefinedNameReferenceFailures(): readonly { readonly owner: DefinedNameReferenceOwnerIdentity; readonly reason: string }[];
  getStructuralFormulaRuleDependents(sheetId: string, axis: 'row' | 'column', at: number): readonly FormulaRuleReferenceOwnerIdentity[];
  getRangeFormulaRuleDependents(sheetId: string, range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): readonly FormulaRuleReferenceOwnerIdentity[];
  getFormulaRuleReferenceFailures(): readonly FormulaRuleReferenceFailure[];
  setFormulaRuleReference(owner: FormulaRuleReferenceOwnerIdentity, references: readonly FormulaReferenceNode[], context: CellAddress, failure?: FormulaRuleReferenceFailureReason): void;
  removeFormulaRuleReference(owner: FormulaRuleReferenceOwnerIdentity): boolean;
  hasFormulaRuleReference(owner: FormulaRuleReferenceOwnerIdentity): boolean;
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
  for (const table of sheet.sheetTables) {
    const tableWidth = table.range.endColumn - table.range.startColumn + 1;
    if (table.columns.length !== tableWidth) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: Sheet Table ${table.id} columns do not match its range width`);
    }
    if (axis === 'column' && direction === 1
      && at > table.range.startColumn && at <= table.range.endColumn) {
      throw new Error(`UNSUPPORTED_FEATURE: inserting a worksheet column inside Sheet Table ${table.id} requires a table-column structural patch`);
    }
  }
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

type StructuralFormulaOwnerLocator =
  | { readonly kind: 'table-sheet-column'; readonly sheetId: string; readonly columnIndex: number; readonly fieldId: string }
  | { readonly kind: 'shape-drawing-payload'; readonly sheetId: string; readonly payloadId: string }
  | { readonly kind: 'chart-text-drawing-payload'; readonly sheetId: string; readonly payloadId: string; readonly field: ChartTextFormulaField }
  | { readonly kind: 'data-view-field'; readonly viewId: string; readonly fieldIndex: number; readonly fieldId: string }
  | { readonly kind: 'cell-style-template-formula'; readonly templateId: string; readonly field: 'formula1' | 'formula2' | 'listSource.formula' };

type StagedStructuralFormulaChange =
  | {
    readonly kind: 'formula';
    readonly participant: string;
    readonly owner: StructuralFormulaOwnerLocator;
    readonly before: string;
    readonly after: string;
  }
  | {
    readonly kind: 'cell-style-template-anchor';
    readonly participant: string;
    readonly templateId: string;
    readonly before: Readonly<CellAddress>;
    readonly after: Readonly<CellAddress>;
  };

function structuralFormulaObjectDelta(change: StagedStructuralFormulaChange): StructuralFormulaObjectOwnerDelta | undefined {
  if (change.kind !== 'formula') return undefined;
  switch (change.owner.kind) {
    case 'chart-text-drawing-payload':
      return { kind: 'formula-object', ownerKind: 'chart-text', sheetId: change.owner.sheetId,
        payloadId: change.owner.payloadId, field: change.owner.field,
        beforeFormula: change.before, afterFormula: change.after };
    case 'shape-drawing-payload':
      return { kind: 'formula-object', ownerKind: 'shape-property', sheetId: change.owner.sheetId,
        payloadId: change.owner.payloadId, beforeFormula: change.before, afterFormula: change.after };
    case 'table-sheet-column':
      return { kind: 'formula-object', ownerKind: 'table-sheet-column', sheetId: change.owner.sheetId,
        fieldId: change.owner.fieldId, beforeFormula: change.before, afterFormula: change.after };
    case 'data-view-field':
      return { kind: 'formula-object', ownerKind: 'data-view-field', viewId: change.owner.viewId,
        fieldId: change.owner.fieldId, beforeFormula: change.before, afterFormula: change.after };
    case 'cell-style-template-formula':
      return { kind: 'formula-object', ownerKind: 'cell-style-template', templateId: change.owner.templateId,
        field: change.owner.field, beforeFormula: change.before, afterFormula: change.after };
  }
}

// Workbook-level formulas have no worksheet-relative origin unless they persist an explicit anchor.
const UNANCHORED_WORKBOOK_FORMULA_OWNER = '';

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
  if (count <= 0) return { kind: 'structural-transform', removedCells: [], clearInputRanges: [], populateInputRanges: [], rewrittenFormulaOwners: [], definedNameOwnerDeltas: [] };
  const reportSheetAfter = sheet.reportSheet
    ? mapReportSheetCoordinates(
      sheet.reportSheet,
      (cell) => {
        const position = axis === 'row' ? cell.row : cell.column;
        const mapped = shiftIndex(position, at, count, direction, axis);
        return mapped === null
          ? null
          : axis === 'row' ? { ...cell, row: mapped } : { ...cell, column: mapped };
      },
      axis === 'row' ? (row) => shiftIndex(row, at, count, direction, axis) : undefined,
      `${direction === 1 ? 'insert' : 'delete'}-${axis}s`,
    )
    : undefined;
  const calculationRanges = structuralAxisInputRanges(sheet, axis, at, count, direction);
  const shift: StructuralShift = {
    axis,
    at,
    count,
    op: direction === 1 ? 'insert' : 'delete',
  };
  const formulaRewrite = preflightFormulaRewrite(workbook, sheet, shift, referenceOwners);
  rejectFormulaGroupMetadataInRange(sheet, axis === 'row'
    ? { sheetId: sheet.id, startRow: at, endRow: sheet.rowCount - 1, startColumn: 0, endColumn: Math.max(sheet.columnCount - 1, 0) }
    : { sheetId: sheet.id, startRow: 0, endRow: Math.max(sheet.rowCount - 1, 0), startColumn: at, endColumn: sheet.columnCount - 1 }, 'axis shift');
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
    shiftRuleRanges(owner.conditionalFormats, axis, at, count, direction, sheet.id);
    shiftRuleRanges(owner.dataValidations, axis, at, count, direction, sheet.id);
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
  if (reportSheetAfter) sheet.reportSheet = reportSheetAfter;
  const formulaRewriteResult = applyFormulaRewritePlan(workbook, sheet.id, shift, undefined, formulaRewrite);
  return {
    kind: 'structural-transform',
    removedCells: removed,
    clearInputRanges: calculationRanges.clearInputRanges,
    populateInputRanges: calculationRanges.populateInputRanges,
    rewrittenFormulaOwners: formulaRewriteResult.owners,
    formulaOwnerDeltas: [...formulaRewriteResult.deltas, ...formulaRewriteResult.formulaRuleDeltas],
    definedNameOwnerDeltas: formulaRewriteResult.definedNameDeltas,
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
  rejectFormulaGroupMetadataInRange(sheet, plan.band, 'cell shift');
  preflightCellShiftMetadata(workbook, sheet, plan);
  const reportSheetAfter = sheet.reportSheet
    ? mapReportSheetCoordinates(
      sheet.reportSheet,
      (cell) => mapCellShiftCoordinateForOwner(referenceShift, cell.row, cell.column),
      undefined,
      'cell-shift',
    )
    : undefined;
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
  if (reportSheetAfter) sheet.reportSheet = reportSheetAfter;
  const formulaRewriteResult = applyFormulaRewritePlan(workbook, sheet.id, shift, referenceShift, formulaRewrite, plan);
  return {
    kind: 'structural-transform',
    removedCells,
    clearInputRanges: [structuredClone(plan.band)],
    populateInputRanges: [structuredClone(plan.band)],
    rewrittenFormulaOwners: formulaRewriteResult.owners,
    formulaOwnerDeltas: [...formulaRewriteResult.deltas, ...formulaRewriteResult.formulaRuleDeltas],
    definedNameOwnerDeltas: formulaRewriteResult.definedNameDeltas,
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
  staged.reportSheet = sheet.reportSheet ? structuredClone(sheet.reportSheet) : undefined;
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

/** Clone only workbook-wide reference owners for a non-target worksheet. */
function cloneStructuralReferenceOwnerSheet(sheet: WorksheetModel): WorksheetModel {
  const staged = new WorksheetModel(sheet.id, sheet.name, sheet.rowCount, sheet.columnCount);
  staged.pivots.push(...structuredClone(sheet.pivots));
  staged.sparklines.push(...structuredClone(sheet.sparklines));
  staged.conditionalFormats.push(...structuredClone(sheet.conditionalFormats));
  staged.dataValidations.push(...structuredClone(sheet.dataValidations));
  for (const [key, payload] of sheet.drawingPayloads) staged.drawingPayloads.set(key, structuredClone(payload));
  for (const [key, hyperlink] of sheet.hyperlinks) staged.hyperlinks.set(key, structuredClone(hyperlink));
  return staged;
}

function cloneStructuralPreflightSheets(workbook: WorkbookModel, targetSheet: WorksheetModel): WorksheetModel[] {
  return workbook.getSheets().map((sheet) => sheet.id === targetSheet.id
    ? cloneStructuralMetadataSheet(sheet)
    : cloneStructuralReferenceOwnerSheet(sheet));
}

function preflightCellShiftMetadata(workbook: WorkbookModel, sheet: WorksheetModel, plan: CellShiftPlan): void {
  const stagedSheets = cloneStructuralPreflightSheets(workbook, sheet);
  const staged = stagedSheets.find((candidate) => candidate.id === sheet.id);
  if (!staged) throw new Error(`STRUCTURAL_PATCH_INVARIANT: worksheet ${sheet.id} is absent from metadata preflight`);
  const tables: WorkbookTableModel[] = [];
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) tables.push(structuredClone(table));
  }
  const sources = new Map<string, DataSourceManifest>();
  for (const [id, source] of workbook.dataModel.sources) {
    if (source.sourceRange?.sheetId === sheet.id) sources.set(id, structuredClone(source));
  }
  const printDocument = workbook.printDocuments.get(sheet.id);
  shiftCellBandMetadata(workbook, staged, plan, tables, stagedSheets, sources,
    printDocument ? structuredClone(printDocument) : undefined);
  if (staged.reportSheet) {
    staged.reportSheet = mapReportSheetCoordinates(
      staged.reportSheet,
      (cell) => mapCellShiftCoordinateForOwner(
        { axis: plan.spec.axis, selection: plan.selection, direction: plan.direction },
        cell.row,
        cell.column,
      ),
      undefined,
      'cell-shift',
    );
  }
}

function preflightAxisMetadata(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): void {
  const stagedSheets = cloneStructuralPreflightSheets(workbook, sheet);
  const staged = stagedSheets.find((candidate) => candidate.id === sheet.id);
  if (!staged) throw new Error(`STRUCTURAL_PATCH_INVARIANT: worksheet ${sheet.id} is absent from metadata preflight`);
  const tables: WorkbookTableModel[] = [];
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceRange?.sheetId === sheet.id) tables.push(structuredClone(table));
  }
  const sources = new Map<string, DataSourceManifest>();
  for (const [id, source] of workbook.dataModel.sources) {
    if (source.sourceRange?.sheetId === sheet.id) sources.set(id, structuredClone(source));
  }
  shiftDataRegionAxis(workbook, staged, axis, at, count, direction, sources);
  shiftMerges(staged, axis, at, count, direction);
  for (const owner of stagedSheets) {
    shiftRuleRanges(owner.conditionalFormats, axis, at, count, direction, staged.id);
    shiftRuleRanges(owner.dataValidations, axis, at, count, direction, staged.id);
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
      if ('listSource' in rule && rule.listSource?.kind === 'range' && rule.listSource.range.sheetId === sheet.id
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
  for (const owner of ownerSheets) for (const [payloadId, payload] of owner.drawingPayloads) {
    if (payload.kind === 'camera' || payload.kind === 'screenshot') {
      if (!shiftRange(payload.sourceRange)) throw new Error(`Cell shift would remove ${payload.kind} source range`);
    } else if (payload.kind === 'chart') {
      if (payload.source.kind === 'worksheet-ranges') {
        for (const range of payload.source.ranges) {
          if (!shiftRange(range)) throw new Error(`Cell shift removes a worksheet source range for chart ${payloadId}`);
        }
        if (payload.source.ranges.length === 0) throw new Error(`Chart ${payloadId} has no worksheet source ranges`);
      }
      else if (payload.source.kind === 'report-range' && !shiftRange(payload.source.range)) throw new Error('Cell shift would remove Chart report binding');
      if (payload.categoryRange && !shiftRange(payload.categoryRange)) throw new Error(`Cell shift removes chart category range ${payloadId}`);
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
      if ('cellLink' in payload && payload.cellLink?.sheetId === sheet.id) {
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
    const mapped = ReferenceTransformDomain.mapInterval(titleSpan.start, titleSpan.end, shift);
    if (mapped.kind === 'out-of-bounds') {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: print title ${axis} interval exceeds worksheet bounds`);
    }
    if (mapped.kind === 'mapped') document[titleKey] = { ...titleSpan, start: mapped.start, end: mapped.end };
    else delete document[titleKey];
  }
  const breakKey = axis === 'row' ? 'row' : 'column';
  for (let index = document.pageBreaks.length - 1; index >= 0; index -= 1) {
    const pageBreak = document.pageBreaks[index]!;
    const coordinate = pageBreak[breakKey];
    if (coordinate === undefined) continue;
    const mapped = shiftIndex(coordinate, at, count, direction, axis);
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
  const plan: MovedFormulaRewritePlan = { cells: [], names: [], rules: [], hyperlinks: [], participantChanges: [] };
  assertDefinedNameReferenceIndexUsable(referenceOwners);
  assertFormulaRuleReferenceIndexUsable(referenceOwners);
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
  const inverseTransformMovedFormula = (formula: string, ownerSheetId: string): string => transformFormula(formula, (ast) => mapAstMovedReferences(ast, {
    selection: destination,
    rowDelta: -rowDelta,
    columnDelta: -columnDelta,
    ownerSheetId,
    targetSheetId: targetSheet.id,
    targetSheetName: targetSheet.name,
    sheetOrder,
  }));

  const referenceOwnersByAddress = new Map<string, StructuralReferenceOwnerAddress>();
  for (const owner of referenceOwners.getRangeDependents(targetSheet.id, selection)) {
    referenceOwnersByAddress.set(structuralOwnerKey(owner), owner);
  }
  for (const owner of referenceOwners.getInvalidFormulaOwners()) {
    referenceOwnersByAddress.set(structuralOwnerKey(owner), owner);
  }
  for (const referenceOwner of referenceOwnersByAddress.values()) {
    const owner = workbook.getSheet(referenceOwner.sheetId);
    const cell = owner.cells.get(referenceOwner.row, referenceOwner.column);
    if (!cell) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula reference owner ${referenceOwner.sheetId}!${referenceOwner.row}:${referenceOwner.column} is missing from the workbook`);
    }
    if (owner.id === targetSheet.id
      && (insideCell(selection, referenceOwner.row, referenceOwner.column) || insideCell(destination, referenceOwner.row, referenceOwner.column))) continue;
    const formula = cell.formula === undefined ? undefined : transformMovedFormula(cell.formula, owner.id);
    const formulaChanged = formula !== undefined && formula !== cell.formula;
    if (formulaChanged && formula !== undefined && cell.formula !== undefined) {
      assertStructuralFormulaRoundTrip(`${owner.id}!${referenceOwner.row}:${referenceOwner.column}.formula`, cell.formula, formula,
        (value) => inverseTransformMovedFormula(value, owner.id));
    }
    const sourceFormula = cell.formulaMetadata?.sourceFormula !== undefined
      ? transformMovedFormula(cell.formulaMetadata.sourceFormula, owner.id)
      : undefined;
    const sourceFormulaChanged = sourceFormula !== undefined && sourceFormula !== cell.formulaMetadata?.sourceFormula;
    if (sourceFormulaChanged && sourceFormula !== undefined && cell.formulaMetadata?.sourceFormula !== undefined) {
      assertStructuralFormulaRoundTrip(`${owner.id}!${referenceOwner.row}:${referenceOwner.column}.sourceFormula`, cell.formulaMetadata.sourceFormula, sourceFormula,
        (value) => inverseTransformMovedFormula(value, owner.id));
    }
    if ((formulaChanged || sourceFormulaChanged)
      && hasFormulaGroupMetadata(cell)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: formula group at ${owner.id}!${referenceOwner.row}:${referenceOwner.column} requires an explicit formula-group transform before a moved range`);
    }
    const barcode = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
      ? transformMovedFormula(cell.presentation.source.formula, owner.id)
      : undefined;
    const barcodeFormulaChanged = barcode !== undefined
      && cell.presentation?.kind === 'barcode'
      && cell.presentation.source.kind === 'formula'
      && barcode !== cell.presentation.source.formula;
    if (barcodeFormulaChanged && barcode !== undefined
      && cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula') {
      assertStructuralFormulaRoundTrip(`${owner.id}!${referenceOwner.row}:${referenceOwner.column}.barcodeFormula`,
        cell.presentation.source.formula, barcode, (value) => inverseTransformMovedFormula(value, owner.id));
    }
    if (formulaChanged || sourceFormulaChanged || barcodeFormulaChanged) {
      plan.cells.push({
        sheetId: owner.id,
        row: referenceOwner.row,
        column: referenceOwner.column,
        before: formulaOwnerState(cell),
        ...(formulaChanged && formula !== undefined ? { formula } : {}),
        ...(sourceFormulaChanged && sourceFormula !== undefined ? { sourceFormula } : {}),
        ...(barcodeFormulaChanged && barcode !== undefined ? { barcodeFormula: barcode } : {}),
      });
    }
  }
  for (const identity of referenceOwners.getRangeFormulaRuleDependents(targetSheet.id, selection)) {
    const rule = getStructuralFormulaRule(workbook, identity);
    const field = identity.field as StructuralFormulaRuleField;
    const formula = structuralRuleFormulaFields(rule).get(field);
    if (formula === undefined) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula rule owner ${identity.sheetId}:${identity.ruleId}.${identity.field} is missing`);
    }
    const ownerSheetId = rule.formulaAnchor?.sheetId ?? rule.sheetId;
    const next = transformMovedFormula(formula, ownerSheetId);
    assertStructuralFormulaRoundTrip(`${identity.sheetId}:${identity.ruleId}.${field}`, formula, next,
      (value) => inverseTransformMovedFormula(value, ownerSheetId));
    if (next !== formula) {
      plan.rules.push({
        owner: identity,
        beforeFormula: formula,
        formula: next,
        beforeRanges: structuredClone(rule.ranges),
      });
    }
  }
  for (const owner of workbook.getSheets()) {
    for (const hyperlink of owner.hyperlinks.values()) {
      const target = hyperlink.target;
      if (target.kind !== 'sheet' || target.sheetId !== targetSheet.id) continue;
      const next = { ...target };
      if (next.row !== undefined && next.column !== undefined
        && insideCell(selection, next.row, next.column)) {
        next.row += rowDelta;
        next.column += columnDelta;
      }
      if (next.address !== undefined) {
        const address = transformMovedFormula(next.address, targetSheet.id);
        assertStructuralFormulaRoundTrip(`hyperlink:${hyperlink.id}.address`, next.address, address,
          (value) => inverseTransformMovedFormula(value, targetSheet.id));
        next.address = address;
      }
      if (JSON.stringify(next) !== JSON.stringify(target)) plan.hyperlinks.push({ hyperlink, target: next });
    }
  }
  const nameOwners = new Map<string, DefinedNameReferenceOwnerIdentity>();
  for (const owner of referenceOwners.getRangeDefinedNameDependents(targetSheet.id, selection)) {
    nameOwners.set(definedNameReferenceOwnerKey(owner), owner);
  }
  for (const owner of referenceOwners.getDefinedNamesAnchoredInRange(targetSheet.id, selection)) {
    nameOwners.set(definedNameReferenceOwnerKey(owner), owner);
  }
  for (const identity of nameOwners.values()) {
    const entry = getIndexedDefinedName(workbook, identity);
    const ownerSheetId = entry.anchor?.sheetId ?? entry.sheetId ?? targetSheet.id;
    const formula = transformMovedFormula(entry.formula, ownerSheetId);
    assertStructuralFormulaRoundTrip(`defined-name:${entry.scope}:${entry.sheetId ?? '*'}:${entry.name}`, entry.formula, formula,
      (value) => inverseTransformMovedFormula(value, ownerSheetId));
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
  plan.participantChanges = preflightWorkbookFormulaOwners(
    workbook,
    targetSheet,
    transformMovedFormula,
    inverseTransformMovedFormula,
    (anchor) => insideCell(selection, anchor.row, anchor.column)
      ? { ...anchor, row: anchor.row + rowDelta, column: anchor.column + columnDelta }
      : anchor,
  );
  return plan;
}

interface MovedFormulaRewritePlan {
  cells: Array<{
    sheetId: string;
    row: number;
    column: number;
    before: StructuralFormulaOwnerState;
    formula?: string;
    sourceFormula?: string;
    barcodeFormula?: string;
  }>;
  names: Array<{
    entry: WorkbookModel['definedNameModels'][number];
    formula: string;
    anchor?: WorkbookModel['definedNameModels'][number]['anchor'];
  }>;
  rules: Array<{
    owner: FormulaRuleReferenceOwnerIdentity;
    beforeFormula: string;
    formula: string;
    beforeRanges: RangeRef[];
  }>;
  hyperlinks: Array<{ hyperlink: CellHyperlink; target: CellHyperlink['target'] }>;
  participantChanges: StagedStructuralFormulaChange[];
}

function applyMovedFormulaRewritePlan(workbook: WorkbookModel, plan: MovedFormulaRewritePlan): FormulaRewriteApplication {
  const deltas: StructuralFormulaOwnerDelta[] = plan.participantChanges.flatMap((change) => {
    const delta = change.kind === 'formula' && change.owner.kind === 'chart-text-drawing-payload'
      ? structuralFormulaObjectDelta(change)
      : undefined;
    return delta ? [delta] : [];
  });
  applyStagedStructuralFormulaChanges(workbook, plan.participantChanges);
  const rewrittenOwners: StructuralReferenceOwnerAddress[] = [];
  const formulaRuleDeltas: StructuralFormulaOwnerDelta[] = [];
  const definedNameDeltas: StructuralDefinedNameOwnerDelta[] = [];
  for (const change of plan.cells) {
    const sheet = workbook.getSheet(change.sheetId);
    const cell = sheet.cells.get(change.row, change.column);
    if (!cell) throw new Error(`STRUCTURAL_PATCH_INVARIANT: moved-range reference owner ${change.sheetId}!${change.row}:${change.column} disappeared`);
    const next: CellData = {
      ...cell,
      ...(change.formula === undefined ? {} : { formula: change.formula }),
    };
    if (change.sourceFormula !== undefined) {
      if (!cell.formulaMetadata || cell.formulaMetadata.preservedOnly) {
        throw new Error('STRUCTURAL_PATCH_INVARIANT: moved formula provenance changed owner type during apply');
      }
      next.formulaMetadata = { ...cell.formulaMetadata, sourceFormula: change.sourceFormula };
    }
    if (change.barcodeFormula !== undefined) {
      if (cell.presentation?.kind !== 'barcode' || cell.presentation.source.kind !== 'formula') {
        throw new Error('STRUCTURAL_PATCH_INVARIANT: moved-range barcode formula changed owner type during apply');
      }
      next.presentation = { ...cell.presentation, source: { ...cell.presentation.source, formula: change.barcodeFormula } };
    }
    sheet.cells.set(change.row, change.column, next);
    rewrittenOwners.push({ sheetId: change.sheetId, row: change.row, column: change.column });
    deltas.push({
      kind: 'formula-cell',
      beforeAddress: { sheetId: change.sheetId, row: change.row, column: change.column },
      afterAddress: { sheetId: change.sheetId, row: change.row, column: change.column },
      before: structuredClone(change.before),
      after: formulaOwnerState(next),
    });
  }
  for (const change of plan.names) {
    workbook.setDefinedName({ ...change.entry, formula: change.formula, anchor: change.anchor });
    definedNameDeltas.push(createDefinedNameOwnerDelta(change.entry, change.formula, change.anchor));
  }
  for (const change of plan.rules) {
    const rule = getStructuralFormulaRule(workbook, change.owner);
    const field = change.owner.field as StructuralFormulaRuleField;
    if (structuralRuleFormulaFields(rule).get(field) !== change.beforeFormula) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${field} changed during move preflight`);
    }
    if (!writeStructuralFormulaRule(rule, field, change.formula)) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${field} changed owner type during move`);
    }
    if (change.beforeRanges.length === 0 || rule.ranges.length === 0) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId} has no range after move`);
    }
    formulaRuleDeltas.push({
      kind: 'formula-rule',
      sheetId: change.owner.sheetId,
      ruleKind: change.owner.ruleKind,
      ruleId: change.owner.ruleId,
      field,
      beforeFormula: change.beforeFormula,
      afterFormula: change.formula,
      beforeRanges: structuredClone(change.beforeRanges),
      afterRanges: structuredClone(rule.ranges),
    });
  }
  for (const change of plan.hyperlinks) change.hyperlink.target = change.target;
  return { owners: rewrittenOwners, deltas, formulaRuleDeltas, definedNameDeltas };
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
  const interval = ReferenceTransformDomain.mapInterval(range[startKey], range[endKey], shift);
  if (interval.kind === 'out-of-bounds') {
    throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${axis} interval exceeds worksheet bounds`);
  }
  if (interval.kind === 'deleted') return false;
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

function shiftRuleRanges(rules: Array<{
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
  for (const rule of rules) {
    if (rule.formulaAnchor?.sheetId === sheetId) {
      const shifted = shiftIndex(axis === 'row' ? rule.formulaAnchor.row : rule.formulaAnchor.column, at, count, direction, axis);
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
      const shifted = shiftIndex(column, at, count, direction, axis);
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
    const shifted = shiftIndex(Number(key), at, count, direction, axis);
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
  assertStructuralPane(sheet.pane);
  if (sheet.pane.kind === 'none') return;
  const maximum = axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  const remap = (position: number, limit: number, label: string): number => {
    if (!Number.isSafeInteger(position) || position < 0 || position > limit) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: pane ${label} is outside worksheet bounds`);
    }
    let next = position;
    if (direction === 1 && position >= at) next += count;
    if (direction === -1 && position > at) next = position >= at + count ? position - count : at;
    if (!Number.isSafeInteger(next) || next < 0 || next > limit) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: pane ${label} exceeds worksheet bounds after structural transform`);
    }
    return next;
  };
  if (axis === 'row') {
    if (sheet.pane.kind === 'frozen') sheet.pane.ySplit = remap(sheet.pane.ySplit, MAX_ROW_INDEX + 1, 'ySplit');
    sheet.pane.startRow = remap(sheet.pane.startRow, maximum, 'startRow');
    assertStructuralPane(sheet.pane);
    return;
  }
  if (sheet.pane.kind === 'frozen') sheet.pane.xSplit = remap(sheet.pane.xSplit, MAX_COLUMN_INDEX + 1, 'xSplit');
  sheet.pane.startColumn = remap(sheet.pane.startColumn, maximum, 'startColumn');
  assertStructuralPane(sheet.pane);
}

function assertStructuralPane(pane: WorksheetPane): void {
  const invalidField = worksheetPaneValidationError(pane);
  if (invalidField !== undefined) throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: pane ${invalidField} is invalid`);
}

function shiftHiddenAndSizes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    remapIndexSet(sheet.hiddenRows, at, count, direction, axis);
    remapSizeMap(sheet.rowHeightsPx, at, count, direction, axis);
    return;
  }
  remapIndexSet(sheet.hiddenColumns, at, count, direction, axis);
  remapSizeMap(sheet.columnWidthsPx, at, count, direction, axis);
}

function remapIndexSet(set: Set<number>, at: number, count: number, direction: 1 | -1, axis: 'row' | 'column'): void {
  const next = new Set<number>();
  for (const value of set) {
    const shifted = shiftIndex(value, at, count, direction, axis);
    if (shifted != null) next.add(shifted);
  }
  set.clear();
  for (const value of next) set.add(value);
}

function remapSizeMap(map: Record<number, number>, at: number, count: number, direction: 1 | -1, axis: 'row' | 'column'): void {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const shifted = shiftIndex(Number(key), at, count, direction, axis);
    if (shifted != null) next[shifted] = value;
  }
  for (const key of Object.keys(map)) delete map[Number(key)];
  Object.assign(map, next);
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
    const shifted = shiftIndex(position, at, count, direction, axis);
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
      const shifted = shiftIndex(position, at, count, direction, axis);
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
        const row = axis === 'row' ? shiftIndex(payload.cellLink.row, at, count, direction, axis) : payload.cellLink.row;
        const column = axis === 'column' ? shiftIndex(payload.cellLink.column, at, count, direction, axis) : payload.cellLink.column;
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
    const shifted = shiftIndex(start, at, count, direction, axis);
    if (shifted == null) throw new Error(`Drawing ${drawing.id} anchor is removed by structural mutation`);
    if (axis === 'row') drawing.anchor.row = shifted;
    else drawing.anchor.column = shifted;
  }
  const end = axis === 'row' ? drawing.anchor.endRow : drawing.anchor.endColumn;
  if (end != null) {
    const shifted = shiftIndex(end, at, count, direction, axis);
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
    const shifted = shiftIndex(position, at, count, direction, axis);
    if (shifted == null) return undefined;
    return { row: axis === 'row' ? shifted : row, column: axis === 'column' ? shifted : column };
  });
  sheet.review.remapCoordinates((row, column) => {
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction, axis);
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
    const shifted = shiftIndex(position, at, count, direction, axis);
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
      const row = axis === 'row' ? shiftIndex(next.row, at, count, direction, axis) : next.row;
      if (row === null) throw new Error(`Structural mutation removes hyperlink ${hyperlink.id} target`);
      next.row = row;
    }
    if (next.column !== undefined) {
      const column = axis === 'column' ? shiftIndex(next.column, at, count, direction, axis) : next.column;
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
    const shifted = shiftIndex(position, at, count, direction, axis);
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
  readonly cells: Array<{
    sheetId: string;
    row: number;
    column: number;
    before: StructuralFormulaOwnerState;
    formula?: string;
    sourceFormula?: string;
    barcodeFormula?: string;
  }>;
  readonly names: Array<{
    entry: WorkbookModel['definedNameModels'][number];
    formula: string;
    anchor?: WorkbookModel['definedNameModels'][number]['anchor'];
  }>;
  readonly formulaRules: Array<{
    owner: FormulaRuleReferenceOwnerIdentity;
    beforeFormula: string;
    afterFormula: string;
    beforeRanges: RangeRef[];
  }>;
  readonly participantChanges: StagedStructuralFormulaChange[];
}

function formulaOwnerState(cell: CellData): StructuralFormulaOwnerState {
  return {
    formula: cell.formula ?? null,
    sourceFormula: cell.formulaMetadata?.sourceFormula ?? null,
    barcodeFormula: cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
      ? cell.presentation.source.formula
      : null,
  };
}

function preflightFormulaRewrite(
  workbook: WorkbookModel,
  targetSheet: WorksheetModel,
  shift: StructuralShift,
  referenceOwners: StructuralReferenceOwnerIndex,
  cellShift?: CellShiftReferenceTransform,
): FormulaRewritePlan {
  const plan: FormulaRewritePlan = { cells: [], names: [], formulaRules: [], participantChanges: [] };
  assertDefinedNameReferenceIndexUsable(referenceOwners);
  assertFormulaRuleReferenceIndexUsable(referenceOwners);
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const mapFormula = (formula: string, ownerSheetId: string): string => transformFormula(formula, (ast) => mapAstStructuralReferences(ast, {
    shift,
    cellShift,
    ownerSheetId,
    targetSheetId: targetSheet.id,
    targetSheetName: targetSheet.name,
    sheetOrder,
  }));
  const inverseShift: StructuralShift = { ...shift, op: shift.op === 'insert' ? 'delete' : 'insert' };
  const inverseCellShift = cellShift ? { ...cellShift, direction: cellShift.direction === 1 ? -1 as const : 1 as const } : undefined;
  const mapFormulaInverse = (formula: string, ownerSheetId: string): string => transformFormula(formula, (ast) => mapAstStructuralReferences(ast, {
    shift: inverseShift,
    cellShift: inverseCellShift,
    ownerSheetId,
    targetSheetId: targetSheet.id,
    targetSheetName: targetSheet.name,
    sheetOrder,
  }));
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
    if (!cell) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula reference owner ${owner.sheetId}!${owner.row}:${owner.column} is missing from the workbook`);
    }
    const formula = cell.formula === undefined ? undefined : mapFormula(cell.formula, sheet.id);
    const sourceFormula = cell.formulaMetadata?.sourceFormula !== undefined
      ? mapFormula(cell.formulaMetadata.sourceFormula, sheet.id)
      : undefined;
    const formulaChanged = formula !== undefined && formula !== cell.formula;
    const sourceFormulaChanged = sourceFormula !== undefined && sourceFormula !== cell.formulaMetadata?.sourceFormula;
    if ((formulaChanged || sourceFormulaChanged)
      && hasFormulaGroupMetadata(cell)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: formula group at ${sheet.id}!${owner.row}:${owner.column} requires an explicit formula-group transform`);
    }
    const barcode = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
      ? mapFormula(cell.presentation.source.formula, sheet.id)
      : undefined;
    const barcodeFormulaChanged = barcode !== undefined
      && cell.presentation?.kind === 'barcode'
      && cell.presentation.source.kind === 'formula'
      && barcode !== cell.presentation.source.formula;
    if (formulaChanged || sourceFormulaChanged || barcodeFormulaChanged) {
      plan.cells.push({
        sheetId: sheet.id,
        row: owner.row,
        column: owner.column,
        before: formulaOwnerState(cell),
        ...(formulaChanged && formula !== undefined ? { formula } : {}),
        ...(sourceFormulaChanged && sourceFormula !== undefined ? { sourceFormula } : {}),
        ...(barcodeFormulaChanged && barcode !== undefined ? { barcodeFormula: barcode } : {}),
      });
    }
  }
  const formulaRuleOwners = new Map<string, FormulaRuleReferenceOwnerIdentity>();
  const affectedFormulaRules = referenceOwners.getStructuralFormulaRuleDependents(targetSheet.id, shift.axis, shift.at);
  for (const owner of affectedFormulaRules) formulaRuleOwners.set(formulaRuleReferenceOwnerKey(owner), owner);
  for (const identity of formulaRuleOwners.values()) {
    const rule = getStructuralFormulaRule(workbook, identity);
    const field = identity.field as StructuralFormulaRuleField;
    const beforeFormula = structuralRuleFormulaFields(rule).get(field);
    if (beforeFormula === undefined) {
      throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula rule owner ${identity.sheetId}:${identity.ruleId}.${identity.field} is missing`);
    }
    const ownerSheetId = rule.formulaAnchor?.sheetId ?? rule.sheetId;
    const afterFormula = mapFormula(beforeFormula, ownerSheetId);
    assertStructuralFormulaRoundTrip(`formula-rule:${identity.ruleKind}:${identity.sheetId}:${identity.ruleId}.${identity.field}`,
      beforeFormula, afterFormula, (value) => mapFormulaInverse(value, ownerSheetId));
    if (afterFormula !== beforeFormula) {
      plan.formulaRules.push({ owner: identity, beforeFormula, afterFormula, beforeRanges: structuredClone(rule.ranges) });
    }
  }
  const nameOwners = new Map<string, DefinedNameReferenceOwnerIdentity>();
  for (const owner of referenceOwners.getStructuralDefinedNameDependents(targetSheet.id, shift.axis, shift.at)) {
    nameOwners.set(definedNameReferenceOwnerKey(owner), owner);
  }
  for (const owner of referenceOwners.getDefinedNamesAnchoredAtOrAfter(targetSheet.id, shift.axis, shift.at)) {
    nameOwners.set(definedNameReferenceOwnerKey(owner), owner);
  }
  for (const identity of nameOwners.values()) {
    const entry = getIndexedDefinedName(workbook, identity);
    const ownerSheetId = entry.anchor?.sheetId ?? entry.sheetId ?? targetSheet.id;
    let anchor = entry.anchor;
    if (anchor?.sheetId === targetSheet.id) {
      if (cellShift) {
        const moved = mapCellShiftCoordinateForOwner(cellShift, anchor.row, anchor.column);
        if (!moved) throw new Error(`Defined name ${entry.name} anchor is removed by structural mutation`);
        anchor = { ...anchor, row: moved.row, column: moved.column };
      } else {
        const position = shift.axis === 'row' ? anchor.row : anchor.column;
        const shifted = shiftIndex(position, shift.at, shift.count, shift.op === 'insert' ? 1 : -1, shift.axis);
        if (shifted === null) throw new Error(`Defined name ${entry.name} anchor is removed by structural mutation`);
        anchor = shift.axis === 'row'
          ? { ...anchor, row: shifted }
          : { ...anchor, column: shifted };
      }
    }
    const formula = mapFormula(entry.formula, ownerSheetId);
    assertStructuralFormulaRoundTrip(`defined-name:${entry.scope}:${entry.sheetId ?? '*'}:${entry.name}`, entry.formula, formula,
      (value) => mapFormulaInverse(value, ownerSheetId));
    if (formula !== entry.formula || anchor !== entry.anchor) plan.names.push({ entry, formula, anchor });
  }
  const mapAddress = (address: CellAddress): CellAddress | null => {
    if (address.sheetId !== targetSheet.id) return address;
    if (cellShift) {
      const moved = mapCellShiftCoordinateForOwner(cellShift, address.row, address.column);
      if (!moved) return null;
      return moved.row === address.row && moved.column === address.column ? address : { ...address, ...moved };
    }
    const position = shift.axis === 'row' ? address.row : address.column;
    const shifted = shiftIndex(position, shift.at, shift.count, shift.op === 'insert' ? 1 : -1, shift.axis);
    if (shifted === null) return null;
    if (shifted === position) return address;
    return shift.axis === 'row' ? { ...address, row: shifted } : { ...address, column: shifted };
  };
  plan.participantChanges.push(...preflightWorkbookFormulaOwners(workbook, targetSheet, mapFormula, mapFormulaInverse, mapAddress));
  return plan;
}

function assertStructuralFormulaRoundTrip(
  participant: string,
  before: string,
  after: string,
  inverse: (formula: string) => string,
): void {
  if (after !== before && !sameStructuralFormula(inverse(after), before)) {
    throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${participant} cannot be restored by the inverse structural operation`);
  }
}

function sameStructuralFormula(left: string, right: string): boolean {
  const canonicalFormula = (formula: string): string => transformFormula(formula, (ast) => ast);
  return canonicalFormula(left) === canonicalFormula(right);
}

/** Formula-bearing workbook/sheet owners outside the cell, rule and name indexes. */
function preflightWorkbookFormulaOwners(
  workbook: WorkbookModel,
  targetSheet: WorksheetModel,
  mapFormula: (formula: string, ownerSheetId: string) => string,
  mapFormulaInverse: (formula: string, ownerSheetId: string) => string,
  mapAddress: (address: CellAddress) => CellAddress | null,
): StagedStructuralFormulaChange[] {
  const changes: StagedStructuralFormulaChange[] = [];
  const sheetIds = new Set(workbook.sheetOrder);
  const stage = (
    owner: StructuralFormulaOwnerLocator,
    participant: string,
    before: string | undefined,
    ownerSheetId: string,
  ): void => {
    if (before === undefined || before.length === 0) return;
    const after = mapFormula(before, ownerSheetId);
    assertStructuralFormulaRoundTrip(participant, before, after, (value) => mapFormulaInverse(value, ownerSheetId));
    if (after === before) return;
    changes.push({ kind: 'formula', owner, participant, before, after });
  };
  const stageAnchor = (
    templateId: string,
    participant: string,
    before: CellAddress | undefined,
  ): void => {
    if (!before || before.sheetId !== targetSheet.id) return;
    const after = mapAddress(before);
    if (!after) throw new Error(`Structural mutation removes ${participant} formula anchor`);
    if (after === before) return;
    changes.push({
      kind: 'cell-style-template-anchor',
      templateId,
      participant,
      before: { ...before },
      after: { ...after },
    });
  };

  for (const sheet of workbook.getSheets()) {
    for (const [columnIndex, column] of (sheet.tableSheet?.columns ?? []).entries()) {
      stage(
        { kind: 'table-sheet-column', sheetId: sheet.id, columnIndex, fieldId: column.fieldId },
        `table-sheet:${sheet.id}.${column.fieldId}`,
        column.formula,
        sheet.id,
      );
    }
    for (const [payloadId, payload] of sheet.drawingPayloads) {
      if (payload.kind === 'shape') {
        stage(
          { kind: 'shape-drawing-payload', sheetId: sheet.id, payloadId },
          `drawing:${payloadId}.propertyFormula`,
          payload.propertyFormula,
          sheet.id,
        );
      } else if (payload.kind === 'chart') {
        for (const { field, formula } of chartTextFormulaEntries(payload)) {
          stage(
            { kind: 'chart-text-drawing-payload', sheetId: sheet.id, payloadId, field },
            `drawing:${payloadId}.${field}`,
            formula,
            sheet.id,
          );
        }
      }
    }
  }
  for (const view of workbook.dataModel.views.values()) {
    for (const [fieldIndex, field] of view.fields.entries()) {
      stage(
        { kind: 'data-view-field', viewId: view.id, fieldIndex, fieldId: field.fieldId },
        `data-view:${view.id}.${field.fieldId}`,
        field.formula,
        UNANCHORED_WORKBOOK_FORMULA_OWNER,
      );
    }
  }
  for (const template of workbook.cellStyleTemplates.values()) {
    const validation = template.dataValidation;
    if (!validation) continue;
    const formulaAnchor: unknown = validation.formulaAnchor;
    if (formulaAnchor !== undefined) {
      if (formulaAnchor === null || typeof formulaAnchor !== 'object' || Array.isArray(formulaAnchor)) {
        throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: cell-style-template:${template.id} formula anchor is invalid`);
      }
      const anchor = formulaAnchor as CellAddress;
      if (!sheetIds.has(anchor.sheetId)
        || !Number.isSafeInteger(anchor.row) || anchor.row < 0 || anchor.row > MAX_ROW_INDEX
        || !Number.isSafeInteger(anchor.column) || anchor.column < 0 || anchor.column > MAX_COLUMN_INDEX) {
        throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: cell-style-template:${template.id} formula anchor is invalid`);
      }
    }
    stageAnchor(template.id, `cell-style-template:${template.id}`, validation.formulaAnchor);
    const formulaOwnerSheetId = validation.formulaAnchor?.sheetId ?? UNANCHORED_WORKBOOK_FORMULA_OWNER;
    if (validation.formula1 && (validation.formula1.trim().startsWith('=') || validation.type === 'custom')) {
      stage(
        { kind: 'cell-style-template-formula', templateId: template.id, field: 'formula1' },
        `cell-style-template:${template.id}.formula1`,
        validation.formula1,
        formulaOwnerSheetId,
      );
    }
    if (validation.formula2 && (validation.formula2.trim().startsWith('=') || validation.type === 'custom')) {
      stage(
        { kind: 'cell-style-template-formula', templateId: template.id, field: 'formula2' },
        `cell-style-template:${template.id}.formula2`,
        validation.formula2,
        formulaOwnerSheetId,
      );
    }
    if (validation.listSource?.kind === 'formula') {
      stage(
        { kind: 'cell-style-template-formula', templateId: template.id, field: 'listSource.formula' },
        `cell-style-template:${template.id}.listSource`,
        validation.listSource.formula,
        formulaOwnerSheetId,
      );
    }
  }
  return changes;
}

function readStructuralFormulaOwner(
  workbook: WorkbookModel,
  owner: StructuralFormulaOwnerLocator,
): string | undefined {
  switch (owner.kind) {
    case 'table-sheet-column': {
      const column = workbook.getSheet(owner.sheetId).tableSheet?.columns[owner.columnIndex];
      return column?.fieldId === owner.fieldId ? column.formula : undefined;
    }
    case 'shape-drawing-payload': {
      const payload = workbook.getSheet(owner.sheetId).drawingPayloads.get(owner.payloadId);
      return payload?.kind === 'shape' ? payload.propertyFormula : undefined;
    }
    case 'chart-text-drawing-payload': {
      const payload = workbook.getSheet(owner.sheetId).drawingPayloads.get(owner.payloadId);
      return payload?.kind === 'chart' ? readChartTextFormula(payload, owner.field) : undefined;
    }
    case 'data-view-field': {
      const field = workbook.dataModel.views.get(owner.viewId)?.fields[owner.fieldIndex];
      return field?.fieldId === owner.fieldId ? field.formula : undefined;
    }
    case 'cell-style-template-formula': {
      const validation = workbook.cellStyleTemplates.get(owner.templateId)?.dataValidation;
      if (owner.field === 'formula1') return validation?.formula1;
      if (owner.field === 'formula2') return validation?.formula2;
      return validation?.listSource?.kind === 'formula' ? validation.listSource.formula : undefined;
    }
  }
}

function readStructuralTemplateAnchor(workbook: WorkbookModel, templateId: string): CellAddress | undefined {
  return workbook.cellStyleTemplates.get(templateId)?.dataValidation?.formulaAnchor;
}

function writeStructuralFormulaOwner(
  workbook: WorkbookModel,
  owner: StructuralFormulaOwnerLocator,
  formula: string,
): boolean {
  switch (owner.kind) {
    case 'table-sheet-column': {
      const column = workbook.getSheet(owner.sheetId).tableSheet?.columns[owner.columnIndex];
      if (!column || column.fieldId !== owner.fieldId) return false;
      column.formula = formula;
      return true;
    }
    case 'shape-drawing-payload': {
      const payload = workbook.getSheet(owner.sheetId).drawingPayloads.get(owner.payloadId);
      if (!payload || payload.kind !== 'shape') return false;
      payload.propertyFormula = formula;
      return true;
    }
    case 'chart-text-drawing-payload': {
      const payload = workbook.getSheet(owner.sheetId).drawingPayloads.get(owner.payloadId);
      if (!payload || payload.kind !== 'chart') return false;
      writeChartTextFormula(payload, owner.field, formula);
      return true;
    }
    case 'data-view-field': {
      const field = workbook.dataModel.views.get(owner.viewId)?.fields[owner.fieldIndex];
      if (!field || field.fieldId !== owner.fieldId) return false;
      field.formula = formula;
      return true;
    }
    case 'cell-style-template-formula': {
      const validation = workbook.cellStyleTemplates.get(owner.templateId)?.dataValidation;
      if (!validation) return false;
      if (owner.field === 'formula1') validation.formula1 = formula;
      else if (owner.field === 'formula2') validation.formula2 = formula;
      else if (validation.listSource?.kind === 'formula') validation.listSource.formula = formula;
      else return false;
      return true;
    }
  }
}

function sameCellAddress(left: Readonly<CellAddress> | undefined, right: Readonly<CellAddress>): boolean {
  return left !== undefined
    && left.sheetId === right.sheetId
    && left.row === right.row
    && left.column === right.column;
}

function applyStagedStructuralFormulaChanges(
  workbook: WorkbookModel,
  changes: readonly StagedStructuralFormulaChange[],
): void {
  for (const change of changes) {
    if (change.kind === 'formula') {
      if (readStructuralFormulaOwner(workbook, change.owner) !== change.before) {
        throw new Error(`STRUCTURAL_PATCH_INVARIANT: ${change.participant} changed during formula preflight`);
      }
    } else if (!sameCellAddress(readStructuralTemplateAnchor(workbook, change.templateId), change.before)) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: ${change.participant} anchor changed during formula preflight`);
    }
  }
  for (const change of changes) {
    if (change.kind === 'formula') {
      if (!writeStructuralFormulaOwner(workbook, change.owner, change.after)) {
        throw new Error(`STRUCTURAL_PATCH_INVARIANT: ${change.participant} disappeared during formula apply`);
      }
      continue;
    }
    const validation = workbook.cellStyleTemplates.get(change.templateId)?.dataValidation;
    if (!validation) throw new Error(`STRUCTURAL_PATCH_INVARIANT: ${change.participant} disappeared during formula apply`);
    validation.formulaAnchor = { ...change.after };
  }
}

function structuralOwnerKey(owner: StructuralReferenceOwnerAddress): string {
  return `${owner.sheetId}\u0000${owner.row}\u0000${owner.column}`;
}

function formulaRuleReferenceOwnerKey(owner: FormulaRuleReferenceOwnerIdentity): string {
  return JSON.stringify([owner.sheetId, owner.ruleKind, owner.ruleId, owner.field]);
}

function getStructuralFormulaRule(
  workbook: WorkbookModel,
  owner: FormulaRuleReferenceOwnerIdentity,
): StructuralFormulaRule {
  const sheet = workbook.getSheet(owner.sheetId);
  const rules = owner.ruleKind === 'conditional-format' ? sheet.conditionalFormats : sheet.dataValidations;
  const matches = rules.filter((rule) => rule.id === owner.ruleId);
  if (matches.length !== 1 || matches[0]!.sheetId !== owner.sheetId) {
    throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: formula rule owner ${owner.sheetId}:${owner.ruleId} is missing or ambiguous`);
  }
  return matches[0] as StructuralFormulaRule;
}

function definedNameReferenceOwnerKey(owner: DefinedNameReferenceOwnerIdentity): string {
  return JSON.stringify([owner.scope, owner.scope === 'sheet' ? owner.sheetId : null, owner.name.trim().toUpperCase()]);
}

function getIndexedDefinedName(workbook: WorkbookModel, owner: DefinedNameReferenceOwnerIdentity): WorkbookModel['definedNameModels'][number] {
  const entry = workbook.getDefinedNameExact(owner.name, owner.scope, owner.sheetId);
  if (!entry) {
    throw new Error(`STRUCTURAL_REFERENCE_INDEX_INVARIANT: defined-name owner ${owner.scope}:${owner.sheetId ?? '*'}:${owner.name} is missing from the workbook`);
  }
  return entry;
}

function assertDefinedNameReferenceIndexUsable(referenceOwners: StructuralReferenceOwnerIndex): void {
  const failure = referenceOwners.getDefinedNameReferenceFailures()[0];
  if (!failure) return;
  const identity = `${failure.owner.scope}:${failure.owner.sheetId ?? '*'}:${failure.owner.name}`;
  const reason = failure.reason === 'unresolved-context'
    ? 'does not have a stable worksheet context'
    : failure.reason === 'invalid-formula'
      ? 'contains a formula that cannot be parsed'
      : 'contains a reference that cannot be resolved';
  throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: defined name ${identity} ${reason}`);
}

function assertFormulaRuleReferenceIndexUsable(referenceOwners: StructuralReferenceOwnerIndex): void {
  const failure = referenceOwners.getFormulaRuleReferenceFailures()[0];
  if (!failure) return;
  const owner = failure.owner;
  const reason = failure.reason === 'invalid-formula'
    ? 'contains a formula that cannot be parsed'
    : failure.reason === 'unresolved-context'
      ? 'does not have a stable worksheet context'
      : failure.reason === 'invalid-range'
        ? 'has an invalid or empty applies-to range'
        : failure.reason === 'invalid-owner'
          ? 'has a missing or ambiguous identity'
          : 'contains a reference that cannot be resolved';
  throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${owner.ruleKind} ${owner.sheetId}:${owner.ruleId}.${owner.field} ${reason}`);
}

function rejectFormulaGroupMetadataInRange(sheet: WorksheetModel, range: RangeRef, operation: string): void {
  sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (cell, row, column) => {
    if (!hasFormulaGroupMetadata(cell)) return;
    throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${cell.formulaMetadata?.kind} formula metadata at ${sheet.id}!${row}:${column} requires an explicit formula-group operation before ${operation}`);
  });
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

interface FormulaRewriteApplication {
  readonly owners: StructuralReferenceOwnerAddress[];
  readonly deltas: StructuralFormulaOwnerDelta[];
  readonly formulaRuleDeltas: StructuralFormulaOwnerDelta[];
  readonly definedNameDeltas: StructuralDefinedNameOwnerDelta[];
}

function applyFormulaRewritePlan(
  workbook: WorkbookModel,
  targetSheetId: string,
  shift: StructuralShift,
  cellShift: CellShiftReferenceTransform | undefined,
  plan: FormulaRewritePlan,
  cellShiftPlan?: CellShiftPlan,
): FormulaRewriteApplication {
  const deltas: StructuralFormulaOwnerDelta[] = plan.participantChanges.flatMap((change) => {
    const delta = change.kind === 'formula' && change.owner.kind === 'chart-text-drawing-payload'
      ? structuralFormulaObjectDelta(change)
      : undefined;
    return delta ? [delta] : [];
  });
  const formulaRuleTargets = plan.formulaRules.map((change) => {
    const rule = getStructuralFormulaRule(workbook, change.owner);
    const beforeFormula = structuralRuleFormulaFields(rule).get(change.owner.field as StructuralFormulaRuleField);
    if (beforeFormula !== change.beforeFormula) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${change.owner.field} changed during preflight`);
    }
    return { change, rule };
  });
  applyStagedStructuralFormulaChanges(workbook, plan.participantChanges);
  const rewrittenOwners: StructuralReferenceOwnerAddress[] = [];
  const formulaRuleDeltas: StructuralFormulaOwnerDelta[] = [];
  const definedNameDeltas: StructuralDefinedNameOwnerDelta[] = [];
  for (const { change, rule } of formulaRuleTargets) {
    if (!writeStructuralFormulaRule(rule, change.owner.field, change.afterFormula)) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${change.owner.field} disappeared during apply`);
    }
    if (change.beforeRanges.length === 0 || rule.ranges.length === 0) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId} has no range after structural transform`);
    }
    formulaRuleDeltas.push({
      kind: 'formula-rule',
      sheetId: change.owner.sheetId,
      ruleKind: change.owner.ruleKind,
      ruleId: change.owner.ruleId,
      field: change.owner.field as StructuralFormulaRuleField,
      beforeFormula: change.beforeFormula,
      afterFormula: change.afterFormula,
      beforeRanges: structuredClone(change.beforeRanges),
      afterRanges: structuredClone(rule.ranges),
    });
  }
  for (const change of plan.cells) {
    const sheet = workbook.getSheet(change.sheetId);
    let coordinate: { row: number; column: number } | null = { row: change.row, column: change.column };
    if (sheet.id === targetSheetId) {
      if (cellShift && cellShiftPlan) {
        coordinate = mapCellShiftCoordinateForOwner(cellShift, change.row, change.column);
      } else {
        const value = shift.axis === 'row' ? change.row : change.column;
        const shifted = shiftIndex(value, shift.at, shift.count, shift.op === 'insert' ? 1 : -1, shift.axis);
        coordinate = shifted === null ? null : shift.axis === 'row'
          ? { row: shifted, column: change.column }
          : { row: change.row, column: shifted };
      }
    }
    if (!coordinate) continue;
    const cell = sheet.cells.get(coordinate.row, coordinate.column);
    if (!cell) throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula reference owner ${sheet.id}!${coordinate.row}:${coordinate.column} was not preserved`);
    const next: CellData = {
      ...cell,
      ...(change.formula === undefined ? {} : { formula: change.formula }),
    };
    if (change.sourceFormula !== undefined) {
      if (!cell.formulaMetadata || cell.formulaMetadata.preservedOnly) {
        throw new Error('STRUCTURAL_PATCH_INVARIANT: formula provenance changed owner type during apply');
      }
      next.formulaMetadata = { ...cell.formulaMetadata, sourceFormula: change.sourceFormula };
    }
    if (change.barcodeFormula !== undefined) {
      if (cell.presentation?.kind !== 'barcode' || cell.presentation.source.kind !== 'formula') {
        throw new Error('STRUCTURAL_PATCH_INVARIANT: barcode formula changed owner type during apply');
      }
      next.presentation = { ...cell.presentation, source: { ...cell.presentation.source, formula: change.barcodeFormula } };
    }
    sheet.cells.set(coordinate.row, coordinate.column, next);
    rewrittenOwners.push({ sheetId: sheet.id, row: coordinate.row, column: coordinate.column });
    deltas.push({
      kind: 'formula-cell',
      beforeAddress: { sheetId: change.sheetId, row: change.row, column: change.column },
      afterAddress: { sheetId: sheet.id, row: coordinate.row, column: coordinate.column },
      before: structuredClone(change.before),
      after: formulaOwnerState(next),
    });
  }
  for (const change of plan.names) {
    workbook.setDefinedName({ ...change.entry, formula: change.formula, anchor: change.anchor });
    definedNameDeltas.push(createDefinedNameOwnerDelta(change.entry, change.formula, change.anchor));
  }
  return { owners: rewrittenOwners, deltas, formulaRuleDeltas, definedNameDeltas };
}

function writeStructuralFormulaRule(rule: StructuralFormulaRule, field: string, formula: string): boolean {
  switch (field as StructuralFormulaRuleField) {
    case 'value1':
      if (typeof rule.value1 !== 'string') return false;
      rule.value1 = formula;
      return true;
    case 'value2':
      if (typeof rule.value2 !== 'string') return false;
      rule.value2 = formula;
      return true;
    case 'formula1':
      if (typeof rule.formula1 !== 'string') return false;
      rule.formula1 = formula;
      return true;
    case 'formula2':
      if (typeof rule.formula2 !== 'string') return false;
      rule.formula2 = formula;
      return true;
    case 'listSource.formula':
      if (rule.listSource?.kind !== 'formula') return false;
      rule.listSource = { ...rule.listSource, formula };
      return true;
    default:
      return false;
  }
}

function createDefinedNameOwnerDelta(
  entry: WorkbookModel['definedNameModels'][number],
  formula: string,
  anchor: WorkbookModel['definedNameModels'][number]['anchor'],
): StructuralDefinedNameOwnerDelta {
  const before: FormulaDefinedName = {
    name: entry.name,
    formula: entry.formula,
    scope: entry.scope,
    ...(entry.sheetId ? { sheetId: entry.sheetId } : {}),
    ...(entry.anchor ? { anchor: structuredClone(entry.anchor) } : {}),
  };
  const after: FormulaDefinedName = {
    ...before,
    formula,
    ...(anchor ? { anchor: structuredClone(anchor) } : { anchor: undefined }),
  };
  return {
    owner: { scope: entry.scope, name: entry.name, ...(entry.sheetId ? { sheetId: entry.sheetId } : {}) },
    before,
    after,
  };
}

export interface SheetTableRenamePlan {
  apply(): StructuralTransformResult;
}

export function planSheetTableRename(
  workbook: WorkbookModel,
  tableId: string,
  nextName: string,
): SheetTableRenamePlan {
  const matches = workbook.getSheets().flatMap((sheet) => sheet.sheetTables
    .filter((table) => table.id === tableId)
    .map((table) => ({ sheet, table })));
  if (matches.length !== 1) throw new Error(`Sheet Table identity must resolve exactly once: ${tableId}`);
  const { sheet: tableSheet, table } = matches[0]!;
  const previousName = table.name;
  const changesIdentity = previousName.trim().toUpperCase() !== nextName.trim().toUpperCase();
  const mapFormula = (formula: string): string => changesIdentity
    ? rewriteFormulaTableReferences(formula, previousName, nextName)
    : formula;
  const inverseFormula = (formula: string): string => changesIdentity
    ? rewriteFormulaTableReferences(formula, nextName, previousName)
    : formula;
  const cellChanges: Array<{
    readonly address: StructuralReferenceOwnerAddress;
    readonly before: StructuralFormulaOwnerState;
    readonly after: StructuralFormulaOwnerState;
  }> = [];
  for (const cellSheet of workbook.getSheets()) {
    cellSheet.cells.forEachFormulaOwner((cell, row, column) => {
      const owner = { sheetId: cellSheet.id, row, column };
      const before = formulaOwnerState(cell);
      const formula = cell.formula === undefined ? undefined : mapFormula(cell.formula);
      const sourceFormula = cell.formulaMetadata?.sourceFormula === undefined
        ? undefined
        : mapFormula(cell.formulaMetadata.sourceFormula);
      const currentBarcodeFormula = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
        ? cell.presentation.source.formula
        : undefined;
      const barcodeFormula = currentBarcodeFormula === undefined ? undefined : mapFormula(currentBarcodeFormula);
      if (formula === cell.formula && sourceFormula === cell.formulaMetadata?.sourceFormula
        && barcodeFormula === currentBarcodeFormula) return;
      const rewritesOnlyPreservedDataTableSource = cell.formulaMetadata?.kind === 'dataTable'
        && cell.formulaMetadata.preservedOnly === true
        && cell.formula === undefined
        && sourceFormula !== undefined
        && sourceFormula !== cell.formulaMetadata.sourceFormula
        && formula === undefined
        && barcodeFormula === currentBarcodeFormula;
      if (hasFormulaGroupMetadata(cell) && !rewritesOnlyPreservedDataTableSource) {
        throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: formula group at ${owner.sheetId}!${owner.row}:${owner.column} requires an explicit table-reference transform`);
      }
      cellChanges.push({
        address: owner,
        before,
        after: {
          formula: formula ?? null,
          sourceFormula: sourceFormula ?? null,
          barcodeFormula: barcodeFormula ?? null,
        },
      });
    });
  }

  const ruleChanges: Array<{
    readonly owner: FormulaRuleReferenceOwnerIdentity;
    readonly beforeFormula: string;
    readonly afterFormula: string;
    readonly beforeRanges: readonly RangeRef[];
  }> = [];
  for (const ownerSheet of workbook.getSheets()) {
    for (const [ruleKind, rules] of [
      ['conditional-format', ownerSheet.conditionalFormats],
      ['data-validation', ownerSheet.dataValidations],
    ] as const) {
      for (const rule of rules) {
        for (const [field, beforeFormula] of structuralRuleFormulaFields(rule)) {
          const afterFormula = mapFormula(beforeFormula);
          if (afterFormula === beforeFormula) continue;
          ruleChanges.push({
            owner: { sheetId: ownerSheet.id, ruleKind, ruleId: rule.id, field },
            beforeFormula,
            afterFormula,
            beforeRanges: structuredClone(rule.ranges),
          });
        }
      }
    }
  }

  const nameChanges = workbook.definedNameModels.flatMap((entry) => {
    const formula = mapFormula(entry.formula);
    return formula === entry.formula ? [] : [{ entry, formula }];
  });
  const participantChanges = preflightWorkbookFormulaOwners(
    workbook,
    tableSheet,
    (formula) => mapFormula(formula),
    (formula) => inverseFormula(formula),
    (address) => address,
  );
  const formulaOwnerDeltas: StructuralFormulaOwnerDelta[] = [
    ...cellChanges.map(({ address, before, after }): StructuralFormulaCellOwnerDelta => ({
      kind: 'formula-cell',
      beforeAddress: { ...address },
      afterAddress: { ...address },
      before: structuredClone(before),
      after: structuredClone(after),
    })),
    ...ruleChanges.map(({ owner, beforeFormula, afterFormula, beforeRanges }): StructuralFormulaRuleOwnerDelta => ({
      kind: 'formula-rule',
      sheetId: owner.sheetId,
      ruleKind: owner.ruleKind,
      ruleId: owner.ruleId,
      field: owner.field as StructuralFormulaRuleField,
      beforeFormula,
      afterFormula,
      beforeRanges: structuredClone(beforeRanges),
      afterRanges: structuredClone(beforeRanges),
    })),
    ...participantChanges.flatMap((change) => {
      const delta = structuralFormulaObjectDelta(change);
      return delta ? [delta] : [];
    }),
  ];
  const definedNameOwnerDeltas = nameChanges.map(({ entry, formula }) => createDefinedNameOwnerDelta(entry, formula, entry.anchor));

  return {
    apply: () => {
      if (table.name !== previousName) throw new Error(`STRUCTURAL_PATCH_PRECONDITION: Sheet Table ${tableId} changed during rename preflight`);
      for (const change of cellChanges) {
        const cell = workbook.getSheet(change.address.sheetId).cells.getFormulaOwnerWithoutHydration(change.address.row, change.address.column);
        if (!cell || JSON.stringify(formulaOwnerState(cell)) !== JSON.stringify(change.before)) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula owner ${change.address.sheetId}!${change.address.row}:${change.address.column} changed during table rename preflight`);
        }
      }
      for (const change of ruleChanges) {
        const rule = getStructuralFormulaRule(workbook, change.owner);
        if (structuralRuleFormulaFields(rule).get(change.owner.field as StructuralFormulaRuleField) !== change.beforeFormula
          || JSON.stringify(rule.ranges) !== JSON.stringify(change.beforeRanges)) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${change.owner.field} changed during table rename preflight`);
        }
      }
      for (const change of nameChanges) {
        const current = workbook.getDefinedNameExact(change.entry.name, change.entry.scope, change.entry.sheetId);
        if (!current || JSON.stringify(current) !== JSON.stringify(change.entry)) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: defined name ${change.entry.name} changed during table rename preflight`);
        }
      }
      for (const change of participantChanges) {
        if (change.kind === 'formula' && readStructuralFormulaOwner(workbook, change.owner) !== change.before) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${change.participant} changed during table rename preflight`);
        }
      }

      for (const change of cellChanges) {
        const cellSheet = workbook.getSheet(change.address.sheetId);
        const cell = cellSheet.cells.getFormulaOwnerWithoutHydration(change.address.row, change.address.column)!;
        const next: CellData = { ...cell };
        if (change.after.formula === null) delete next.formula;
        else next.formula = change.after.formula;
        if (change.after.formula !== change.before.formula) delete next.formulaValue;
        if (change.after.sourceFormula !== null) {
          if (!next.formulaMetadata) throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula provenance owner ${change.address.sheetId}!${change.address.row}:${change.address.column} disappeared`);
          next.formulaMetadata = { ...next.formulaMetadata, sourceFormula: change.after.sourceFormula };
        }
        if (change.after.barcodeFormula !== null) {
          if (next.presentation?.kind !== 'barcode' || next.presentation.source.kind !== 'formula') {
            throw new Error(`STRUCTURAL_PATCH_INVARIANT: barcode formula owner ${change.address.sheetId}!${change.address.row}:${change.address.column} changed type`);
          }
          next.presentation = { ...next.presentation, source: { ...next.presentation.source, formula: change.after.barcodeFormula } };
        }
        if (!cellSheet.cells.replaceFormulaOwnerWithoutHydration(change.address.row, change.address.column, next)) {
          throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula owner ${change.address.sheetId}!${change.address.row}:${change.address.column} disappeared during table rename`);
        }
      }
      for (const change of ruleChanges) {
        const rule = getStructuralFormulaRule(workbook, change.owner);
        if (!writeStructuralFormulaRule(rule, change.owner.field, change.afterFormula)) {
          throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula rule ${change.owner.sheetId}:${change.owner.ruleId}.${change.owner.field} changed type`);
        }
      }
      for (const change of nameChanges) workbook.setDefinedName({ ...change.entry, formula: change.formula });
      applyStagedStructuralFormulaChanges(workbook, participantChanges);
      return {
        kind: 'structural-transform',
        removedCells: [],
        clearInputRanges: [],
        populateInputRanges: [],
        rewrittenFormulaOwners: cellChanges.map(({ address }) => ({ ...address })),
        ...(formulaOwnerDeltas.length > 0 ? { formulaOwnerDeltas } : {}),
        ...(definedNameOwnerDeltas.length > 0 ? { definedNameOwnerDeltas } : {}),
      };
    },
  };
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
  rejectFormulaGroupMetadataInRange(sheet, normalizedSource, 'move-range source');
  rejectFormulaGroupMetadataInRange(sheet, target, 'move-range destination');
  validateMoveMetadataPreservation(workbook, sheet, normalizedSource, target);
  validateDataRegionMovePreservation(sheet, normalizedSource, target);

  const rowDelta = target.startRow - normalizedSource.startRow;
  const colDelta = target.startColumn - normalizedSource.startColumn;
  const reportSheetAfter = sheet.reportSheet
    ? mapReportSheetCoordinates(
      sheet.reportSheet,
      (cell) => {
        if (insideCell(target, cell.row, cell.column) && !insideCell(normalizedSource, cell.row, cell.column)) {
          throw new Error(`Cannot move range: report binding at ${cell.row}:${cell.column} would be overwritten`);
        }
        return insideCell(normalizedSource, cell.row, cell.column)
          ? { row: cell.row + rowDelta, column: cell.column + colDelta }
          : { ...cell };
      },
      undefined,
      'move-range',
    )
    : undefined;
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const cellsToMove = sheet.cells.getRegion(
    normalizedSource.startRow,
    normalizedSource.endRow,
    normalizedSource.startColumn,
    normalizedSource.endColumn,
  ).map((entry) => {
    const mapMovedReferences = (ast: ReturnType<typeof parseFormula>) => mapAstMovedReferences(ast, {
      selection: normalizedSource,
      rowDelta,
      columnDelta: colDelta,
      ownerSheetId: sheet.id,
      targetSheetId: sheet.id,
      targetSheetName: sheet.name,
      sheetOrder,
    });
    const inverseMapMovedReferences = (ast: ReturnType<typeof parseFormula>) => mapAstMovedReferences(ast, {
      selection: target,
      rowDelta: -rowDelta,
      columnDelta: -colDelta,
      ownerSheetId: sheet.id,
      targetSheetId: sheet.id,
      targetSheetName: sheet.name,
      sheetOrder,
    });
    const formula = entry.cell.formula === undefined
      ? undefined
      : transformFormula(entry.cell.formula, mapMovedReferences);
    if (formula !== undefined && entry.cell.formula !== undefined) {
      assertStructuralFormulaRoundTrip(`${sheet.id}!${entry.row}:${entry.column}.formula`, entry.cell.formula, formula,
        (value) => transformFormula(value, inverseMapMovedReferences));
    }
    if (formula !== entry.cell.formula && entry.cell.formulaMetadata?.preservedOnly) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: preserved-only formula at ${sheet.id}!${entry.row}:${entry.column} cannot be rewritten for a moved range`);
    }
    if (entry.cell.formulaMetadata?.preservedOnly && entry.cell.formulaMetadata.sourceFormula) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: preserved-only formula provenance at ${sheet.id}!${entry.row}:${entry.column} cannot move safely`);
    }
    let cell = formula === undefined || formula === entry.cell.formula
      ? entry.cell
      : { ...entry.cell, formula };
    const sourceFormula = entry.cell.formulaMetadata?.sourceFormula !== undefined
      ? transformFormula(entry.cell.formulaMetadata.sourceFormula, mapMovedReferences)
      : undefined;
    if (sourceFormula !== undefined && entry.cell.formulaMetadata?.sourceFormula !== undefined) {
      assertStructuralFormulaRoundTrip(`${sheet.id}!${entry.row}:${entry.column}.sourceFormula`,
        entry.cell.formulaMetadata.sourceFormula, sourceFormula,
        (value) => transformFormula(value, inverseMapMovedReferences));
    }
    if (sourceFormula !== undefined && sourceFormula !== entry.cell.formulaMetadata?.sourceFormula) {
      if (!entry.cell.formulaMetadata) throw new Error('STRUCTURAL_PATCH_INVARIANT: formula provenance disappeared during move preflight');
      cell = { ...cell, formulaMetadata: { ...entry.cell.formulaMetadata, sourceFormula } };
    }
    if (entry.cell.presentation?.kind === 'barcode' && entry.cell.presentation.source.kind === 'formula') {
      const barcodeFormula = transformFormula(entry.cell.presentation.source.formula, mapMovedReferences);
      assertStructuralFormulaRoundTrip(`${sheet.id}!${entry.row}:${entry.column}.barcodeFormula`,
        entry.cell.presentation.source.formula, barcodeFormula,
        (value) => transformFormula(value, inverseMapMovedReferences));
      if (barcodeFormula !== entry.cell.presentation.source.formula) {
        cell = {
          ...cell,
          presentation: { ...entry.cell.presentation, source: { ...entry.cell.presentation.source, formula: barcodeFormula } },
        };
      }
    }
    return { ...entry, before: formulaOwnerState(entry.cell), cell };
  });
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
  const formulaRewriteResult = applyMovedFormulaRewritePlan(workbook, formulaRewrite);
  const movedFormulaDeltas: StructuralFormulaOwnerDelta[] = [];
  for (const item of cellsToMove) {
    const after = formulaOwnerState(item.cell);
    if (item.before.formula === after.formula
      && item.before.sourceFormula === after.sourceFormula
      && item.before.barcodeFormula === after.barcodeFormula) continue;
    movedFormulaDeltas.push({
      kind: 'formula-cell',
      beforeAddress: { sheetId: sheet.id, row: item.row, column: item.column },
      afterAddress: { sheetId: sheet.id, row: item.row + rowDelta, column: item.column + colDelta },
      before: structuredClone(item.before),
      after,
    });
  }
  if (reportSheetAfter) sheet.reportSheet = reportSheetAfter;
  return {
    kind: 'structural-transform',
    removedCells: overwritten,
    clearInputRanges: [structuredClone(normalizedSource), structuredClone(target)],
    populateInputRanges: [structuredClone(normalizedSource), structuredClone(target)],
    rewrittenFormulaOwners: formulaRewriteResult.owners,
    formulaOwnerDeltas: [...movedFormulaDeltas, ...formulaRewriteResult.deltas, ...formulaRewriteResult.formulaRuleDeltas],
    definedNameOwnerDeltas: formulaRewriteResult.definedNameDeltas,
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
