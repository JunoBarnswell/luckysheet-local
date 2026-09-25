import {
  formatFormula,
  parseFormula,
  renameAstSheetReferences,
} from '@react-sheets/formula-engine';
import type {
  CellStyleTemplate,
  RangeRef,
  SheetId,
  WorksheetModel,
  WorkbookModel,
  AutoFilterModel,
  ConditionalFormatRule,
  DataValidationRule,
} from './index';
import { CALCULATION_CONTEXT_EFFECTS } from './calculation-context-effect';
import type {
  DrawingPayload,
  DefinedNameModel,
  HyperlinkTarget,
} from './domain';
import { chartTextFormulaEntries, writeChartTextFormula } from './chart-text-reference';
import type { PivotModel, PivotSource } from './pivot';
import type { StructuralFormulaObjectOwnerDelta, StructuralTransformResult } from './structural-transform';

export type SheetIdentityTransformKind = 'rename' | 'duplicate' | 'delete';

/**
 * The only input accepted by the worksheet lifecycle reference graph.
 * `sourceName` is required even though identity is SheetId because authored
 * formulas are still stored as text and must be transformed losslessly.
 */
export interface SheetIdentityTransformSpec {
  kind: SheetIdentityTransformKind;
  sourceSheetId: SheetId;
  sourceName: string;
  targetSheetId?: SheetId;
  targetName?: string;
}

export interface SheetReferenceInvalidation {
  participant: string;
  ownerSheetId?: SheetId;
  reference: string;
  reason: 'deleted-sheet-reference' | 'unsupported-formula';
}

export class SheetIdentityTransformError extends Error {
  readonly code = 'SHEET_IDENTITY_TRANSFORM_REJECTED';
  constructor(
    message: string,
    readonly invalidations: readonly SheetReferenceInvalidation[] = [],
  ) {
    super(message);
    this.name = 'SheetIdentityTransformError';
  }
}

export class SheetIdentityTransformInvariantError extends Error {
  readonly code = 'SHEET_IDENTITY_TRANSFORM_INVARIANT';

  constructor(message: string) {
    super(message);
    this.name = 'SheetIdentityTransformInvariantError';
  }
}

export interface SheetIdentityTransformPlan {
  readonly spec: Readonly<SheetIdentityTransformSpec>;
  readonly invalidations: readonly SheetReferenceInvalidation[];
  apply(): StructuralTransformResult | undefined;
}

type FormulaChange = {
  sheetId: SheetId;
  row: number;
  column: number;
  formula?: { before: string; after: string };
  sourceFormula?: { before: string; after: string };
  barcodeFormula?: { before: string; after: string };
};

function mapSheetId(sheetId: SheetId, sourceSheetId: SheetId, targetSheetId: SheetId): SheetId {
  return sheetId === sourceSheetId ? targetSheetId : sheetId;
}

function mapRange(range: RangeRef, sourceSheetId: SheetId, targetSheetId: SheetId): RangeRef {
  return { ...range, sheetId: mapSheetId(range.sheetId, sourceSheetId, targetSheetId) };
}

function mapFormula(formula: string, oldName: string, newName: string, participant: string): string {
  try {
    return formatFormula(renameAstSheetReferences(parseFormula(formula), oldName, newName));
  } catch (error) {
    throw new SheetIdentityTransformError(
      `${participant} contains an unsupported formula: ${formula}`,
      [{ participant, reference: formula, reason: 'unsupported-formula' }],
    );
  }
}

function formulaReferencesSheet(formula: string, sheetName: string, participant: string, ownerSheetId?: SheetId): boolean {
  try {
    const ast = parseFormula(formula);
    return formatFormula(renameAstSheetReferences(ast, sheetName, `${sheetName}__deleted__`)) !== formatFormula(ast);
  } catch (error) {
    throw new SheetIdentityTransformError(
      `${participant} contains an unsupported formula: ${formula}`,
      [{ participant, ownerSheetId, reference: formula, reason: 'unsupported-formula' }],
    );
  }
}

function mapFormulaReference(
  formula: string,
  oldName: string,
  newName: string,
  participant: string,
  ownerSheetId?: SheetId,
  preservedOnly = false,
): string {
  if (!formulaReferencesSheet(formula, oldName, participant, ownerSheetId)) return formula;
  if (preservedOnly) {
    throw new SheetIdentityTransformError(`${participant} preserves a formula that cannot be rewritten safely`, [
      { participant, ownerSheetId, reference: formula, reason: 'unsupported-formula' },
    ]);
  }
  return mapFormula(formula, oldName, newName, participant);
}

function collectFormulaChanges(
  workbook: WorkbookModel,
  oldName: string,
  newName: string,
): { changes: FormulaChange[]; requiresCalculationContextRebuild: boolean } {
  const changes: FormulaChange[] = [];
  let requiresCalculationContextRebuild = false;
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const hasBarcodeFormula = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula';
      if (!cell.formula && !cell.formulaMetadata?.sourceFormula && !hasBarcodeFormula) return;
      const participant = `cell:${sheet.id}!${row},${column}`;
      const change: FormulaChange = { sheetId: sheet.id, row, column };
      if (cell.formula) {
        if (formulaReferencesSheet(cell.formula, newName, participant, sheet.id)) {
          requiresCalculationContextRebuild = true;
        }
        const formula = mapFormula(cell.formula, oldName, newName, participant);
        if (formula !== cell.formula) change.formula = { before: cell.formula, after: formula };
      }
      if (cell.formulaMetadata?.sourceFormula) {
        if (formulaReferencesSheet(cell.formulaMetadata.sourceFormula, newName, `${participant}.sourceFormula`, sheet.id)) {
          requiresCalculationContextRebuild = true;
        }
        const sourceFormula = mapFormulaReference(cell.formulaMetadata.sourceFormula, oldName, newName, `${participant}.sourceFormula`, sheet.id, cell.formulaMetadata.preservedOnly);
        if (sourceFormula !== cell.formulaMetadata.sourceFormula) change.sourceFormula = { before: cell.formulaMetadata.sourceFormula, after: sourceFormula };
      }
      if (cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula') {
        if (formulaReferencesSheet(cell.presentation.source.formula, newName, `${participant}.barcode`, sheet.id)) {
          requiresCalculationContextRebuild = true;
        }
        const formula = mapFormulaReference(cell.presentation.source.formula, oldName, newName, `${participant}.barcode`, sheet.id);
        if (formula !== cell.presentation.source.formula) change.barcodeFormula = { before: cell.presentation.source.formula, after: formula };
      }
      if (change.formula || change.sourceFormula || change.barcodeFormula) changes.push(change);
    });
  }
  return { changes, requiresCalculationContextRebuild };
}

function transformDefinedNames(workbook: WorkbookModel, oldName: string, newName: string): DefinedNameModel[] {
  return workbook.definedNameModels.map((entry) => ({
    ...entry,
    formula: mapFormula(entry.formula, oldName, newName, `defined-name:${entry.name}`),
    anchor: entry.anchor ? { ...entry.anchor } : undefined,
  }));
}

function rewriteRuleFormulas<T extends ConditionalFormatRule | DataValidationRule>(rule: T, oldName: string, newName: string): T {
  const next = structuredClone(rule);
  if ('value1' in next && typeof next.value1 === 'string') next.value1 = mapFormula(next.value1, oldName, newName, `${next.id}.value1`);
  if ('value2' in next && typeof next.value2 === 'string') next.value2 = mapFormula(next.value2, oldName, newName, `${next.id}.value2`);
  if ('formula1' in next && next.formula1) next.formula1 = mapFormula(next.formula1, oldName, newName, `${next.id}.formula1`);
  if ('formula2' in next && next.formula2) next.formula2 = mapFormula(next.formula2, oldName, newName, `${next.id}.formula2`);
  if ('listSource' in next && next.listSource?.kind === 'formula') {
    next.listSource = { ...next.listSource, formula: mapFormula(next.listSource.formula, oldName, newName, `${next.id}.listSource`) };
  }
  return next as T;
}

function rewriteCellStyleTemplateFormulas(template: CellStyleTemplate, oldName: string, newName: string): CellStyleTemplate {
  const next = structuredClone(template);
  const validation = next.dataValidation;
  if (!validation) return next;
  if (validation.formula1) validation.formula1 = mapFormulaReference(validation.formula1, oldName, newName, `cell-style-template:${template.id}.formula1`);
  if (validation.formula2) validation.formula2 = mapFormulaReference(validation.formula2, oldName, newName, `cell-style-template:${template.id}.formula2`);
  if (validation.listSource?.kind === 'formula') {
    validation.listSource.formula = mapFormulaReference(validation.listSource.formula, oldName, newName, `cell-style-template:${template.id}.listSource`);
  }
  return next;
}

function remapHyperlinkTarget(target: HyperlinkTarget, sourceSheetId: SheetId, targetSheetId: SheetId): HyperlinkTarget {
  return target.kind === 'sheet' && target.sheetId === sourceSheetId
    ? { ...target, sheetId: targetSheetId }
    : structuredClone(target);
}

function remapPivotSource(source: PivotSource, sourceSheetId: SheetId, targetSheetId: SheetId, tableIds: ReadonlyMap<string, string>): PivotSource {
  if (source.kind === 'worksheet-range') return { ...source, range: mapRange(source.range, sourceSheetId, targetSheetId) };
  if (source.kind === 'worksheet-ranges') {
    return {
      ...source,
      ranges: source.ranges.map((item) => ({ ...item, range: mapRange(item.range, sourceSheetId, targetSheetId) })),
    };
  }
  if (source.kind === 'named-range') {
    return source.sheetId === sourceSheetId ? { ...source, sheetId: targetSheetId } : { ...source };
  }
  if (source.kind === 'table') return { ...source, tableId: tableIds.get(source.tableId) ?? source.tableId };
  return { ...source };
}

function remapPivot(pivot: PivotModel, sourceSheetId: SheetId, targetSheetId: SheetId, pivotId: string, tableIds: ReadonlyMap<string, string>): PivotModel {
  const next = structuredClone(pivot);
  next.id = pivotId;
  next.source = remapPivotSource(next.source, sourceSheetId, targetSheetId, tableIds);
  next.target = { ...next.target, sheetId: mapSheetId(next.target.sheetId, sourceSheetId, targetSheetId) };
  return next;
}

function remapDrawingPayload(
  payload: DrawingPayload,
  sourceSheetId: SheetId,
  targetSheetId: SheetId,
  drawingIds: ReadonlyMap<string, string>,
  pivotIds: ReadonlyMap<string, string>,
  tableIds: ReadonlyMap<string, string>,
  sourceName: string,
  targetName: string,
): DrawingPayload {
  const next = structuredClone(payload);
  switch (next.kind) {
    case 'connector':
      next.start = { ...next.start, drawingId: drawingIds.get(next.start.drawingId) ?? next.start.drawingId };
      next.end = { ...next.end, drawingId: drawingIds.get(next.end.drawingId) ?? next.end.drawingId };
      break;
    case 'camera':
      next.sourceRange = mapRange(next.sourceRange, sourceSheetId, targetSheetId);
      break;
    case 'screenshot':
      next.sourceRange = mapRange(next.sourceRange, sourceSheetId, targetSheetId);
      break;
    case 'form-control':
      if ('cellLink' in next && next.cellLink) next.cellLink = { ...next.cellLink, sheetId: mapSheetId(next.cellLink.sheetId, sourceSheetId, targetSheetId) };
      if ('inputRange' in next) next.inputRange = mapRange(next.inputRange, sourceSheetId, targetSheetId);
      break;
    case 'slicer':
    case 'timeline':
      next.pivotId = pivotIds.get(next.pivotId) ?? next.pivotId;
      next.connections = next.connections?.map((connection) => ({ ...connection, pivotId: pivotIds.get(connection.pivotId) ?? connection.pivotId }));
      break;
    case 'shape':
      if (next.hyperlink) next.hyperlink = remapHyperlinkTarget(next.hyperlink, sourceSheetId, targetSheetId);
      if (next.propertyFormula) next.propertyFormula = mapFormulaReference(next.propertyFormula, sourceName, targetName, `duplicate-shape:${sourceSheetId}.propertyFormula`);
      break;
    case 'chart':
      next.source = next.source.kind === 'worksheet-ranges'
        ? { ...next.source, ranges: next.source.ranges.map((range) => mapRange(range, sourceSheetId, targetSheetId)) }
        : next.source.kind === 'report-range'
          ? { ...next.source, range: mapRange(next.source.range, sourceSheetId, targetSheetId) }
          : next.source.kind === 'table'
            ? { ...next.source, tableId: tableIds.get(next.source.tableId) ?? next.source.tableId }
            : { ...next.source, pivotId: pivotIds.get(next.source.pivotId) ?? next.source.pivotId };
      if (next.categoryRange) next.categoryRange = mapRange(next.categoryRange, sourceSheetId, targetSheetId);
      next.series = next.series?.map((series) => ({
        ...series,
        range: mapRange(series.range, sourceSheetId, targetSheetId),
        xRange: series.xRange ? mapRange(series.xRange, sourceSheetId, targetSheetId) : undefined,
        yRange: series.yRange ? mapRange(series.yRange, sourceSheetId, targetSheetId) : undefined,
        sizeRange: series.sizeRange ? mapRange(series.sizeRange, sourceSheetId, targetSheetId) : undefined,
        categoryRange: series.categoryRange ? mapRange(series.categoryRange, sourceSheetId, targetSheetId) : undefined,
        stockRoles: series.stockRoles ? {
          ...series.stockRoles,
          open: series.stockRoles.open ? mapRange(series.stockRoles.open, sourceSheetId, targetSheetId) : undefined,
          high: mapRange(series.stockRoles.high, sourceSheetId, targetSheetId),
          low: mapRange(series.stockRoles.low, sourceSheetId, targetSheetId),
          close: mapRange(series.stockRoles.close, sourceSheetId, targetSheetId),
          volume: series.stockRoles.volume ? mapRange(series.stockRoles.volume, sourceSheetId, targetSheetId) : undefined,
        } : undefined,
        errorBars: series.errorBars ? {
          ...series.errorBars,
          plusRange: series.errorBars.plusRange ? mapRange(series.errorBars.plusRange, sourceSheetId, targetSheetId) : undefined,
          minusRange: series.errorBars.minusRange ? mapRange(series.errorBars.minusRange, sourceSheetId, targetSheetId) : undefined,
        } : undefined,
        dataLabels: series.dataLabels?.valuesFromCells ? { ...series.dataLabels, valuesFromCells: mapRange(series.dataLabels.valuesFromCells, sourceSheetId, targetSheetId) } : series.dataLabels,
        trendlines: series.trendlines ? structuredClone(series.trendlines) : undefined,
      }));
      for (const { field, formula } of chartTextFormulaEntries(next)) {
        writeChartTextFormula(next, field, mapFormulaReference(
          formula,
          sourceName,
          targetName,
          `duplicate-chart:${sourceSheetId}.${field}`,
          sourceSheetId,
        ));
      }
      break;
    default:
      break;
  }
  return next;
}

function allocateId(existing: ReadonlySet<string>, sourceId: string, targetSheetId: string): string {
  const stem = `${sourceId}::${targetSheetId}`;
  let candidate = stem;
  let suffix = 2;
  while (existing.has(candidate)) candidate = `${stem}::${suffix++}`;
  return candidate;
}

function cloneWorksheetWithIdentity(workbook: WorkbookModel, source: WorksheetModel, targetSheetId: SheetId, targetName: string): WorksheetModel {
  const copy = source.cloneWithIdentity(targetSheetId, targetName);
  const allSheets = workbook.getSheets();
  const ids = (selector: (sheet: WorksheetModel) => string[]): Set<string> => new Set(allSheets.flatMap(selector));
  const tableIds = new Map<string, string>();
  for (const table of copy.sheetTables) tableIds.set(table.id, allocateId(ids((sheet) => sheet.sheetTables.map((entry) => entry.id)), table.id, targetSheetId));
  const pivotIds = new Map<string, string>();
  for (const pivot of copy.pivots) pivotIds.set(pivot.id, allocateId(ids((sheet) => sheet.pivots.map((entry) => entry.id)), pivot.id, targetSheetId));
  const drawingIds = new Map<string, string>();
  for (const drawing of copy.drawings) drawingIds.set(drawing.id, allocateId(ids((sheet) => sheet.drawings.map((entry) => entry.id)), drawing.id, targetSheetId));
  const payloadIds = new Map<string, string>();
  for (const payloadId of copy.drawingPayloads.keys()) payloadIds.set(payloadId, allocateId(ids((sheet) => [...sheet.drawingPayloads.keys()]), payloadId, targetSheetId));
  const groupIds = new Map<string, string>();
  for (const group of copy.drawingGroups) groupIds.set(group.id, allocateId(ids((sheet) => sheet.drawingGroups.map((entry) => entry.id)), group.id, targetSheetId));
  const sparklineIds = new Map<string, string>();
  for (const sparkline of copy.sparklines) sparklineIds.set(sparkline.id, allocateId(ids((sheet) => sheet.sparklines.map((entry) => entry.id)), sparkline.id, targetSheetId));
  const sparklineGroupIds = new Map<string, string>();
  for (const group of copy.sparklineGroups) sparklineGroupIds.set(group.id, allocateId(ids((sheet) => sheet.sparklineGroups.map((entry) => entry.id)), group.id, targetSheetId));
  const reviewIds = ids((sheet) => [
    ...sheet.review.noteEntries().map((entry) => entry.note.id),
    ...sheet.review.threadEntries().flatMap((thread) => [thread.id, ...thread.replies.map((reply) => reply.id)]),
  ]);
  const allocateReviewId = (sourceId: string): string => {
    const allocated = allocateId(reviewIds, sourceId, targetSheetId);
    reviewIds.add(allocated);
    return allocated;
  };

  copy.replaceDataRegions(copy.dataRegions.map((region) => ({ ...region, range: mapRange(region.range, source.id, targetSheetId) })));
  copy.merges.splice(0, copy.merges.length, ...copy.merges.map((merge) => ({ ...merge, range: mapRange(merge.range, source.id, targetSheetId) })));
  copy.pivots.splice(0, copy.pivots.length, ...copy.pivots.map((pivot) => remapPivot(pivot, source.id, targetSheetId, pivotIds.get(pivot.id)!, tableIds)));
  copy.sparklines.splice(0, copy.sparklines.length, ...copy.sparklines.map((sparkline) => ({ ...sparkline, id: sparklineIds.get(sparkline.id)!, sheetId: targetSheetId, sourceRange: mapRange(sparkline.sourceRange, source.id, targetSheetId), groupId: sparkline.groupId ? sparklineGroupIds.get(sparkline.groupId) : undefined })));
  copy.sparklineGroups.splice(0, copy.sparklineGroups.length, ...copy.sparklineGroups.map((group) => ({ ...group, id: sparklineGroupIds.get(group.id)!, sheetId: targetSheetId, sparklineIds: group.sparklineIds.map((id) => sparklineIds.get(id) ?? id) })));
  copy.conditionalFormats.splice(0, copy.conditionalFormats.length, ...copy.conditionalFormats.map((rule) => ({ ...structuredClone(rule), id: allocateId(ids((sheet) => sheet.conditionalFormats.map((entry) => entry.id)), rule.id, targetSheetId), sheetId: targetSheetId, ranges: rule.ranges.map((range) => mapRange(range, source.id, targetSheetId)), formulaAnchor: rule.formulaAnchor ? { ...rule.formulaAnchor, sheetId: targetSheetId } : undefined })));
  copy.dataValidations.splice(0, copy.dataValidations.length, ...copy.dataValidations.map((rule) => ({ ...structuredClone(rule), id: allocateId(ids((sheet) => sheet.dataValidations.map((entry) => entry.id)), rule.id, targetSheetId), sheetId: targetSheetId, ranges: rule.ranges.map((range) => mapRange(range, source.id, targetSheetId)), formulaAnchor: rule.formulaAnchor ? { ...rule.formulaAnchor, sheetId: targetSheetId } : undefined, listSource: rule.listSource?.kind === 'range' ? { ...rule.listSource, range: mapRange(rule.listSource.range, source.id, targetSheetId) } : rule.listSource })));
  copy.sheetTables.splice(0, copy.sheetTables.length, ...copy.sheetTables.map((table) => ({ ...structuredClone(table), id: tableIds.get(table.id)!, sheetId: targetSheetId, range: mapRange(table.range, source.id, targetSheetId), autoFilter: table.autoFilter ? { ...table.autoFilter, sheetId: targetSheetId, range: mapRange(table.autoFilter.range, source.id, targetSheetId) } : undefined })));
  copy.drawings.splice(0, copy.drawings.length, ...copy.drawings.map((drawing) => ({ ...drawing, id: drawingIds.get(drawing.id)!, sheetId: targetSheetId, payloadId: payloadIds.get(drawing.payloadId) ?? drawing.payloadId })));
  const payloads = [...copy.drawingPayloads.entries()].map(([id, payload]) => [payloadIds.get(id) ?? id, remapDrawingPayload(payload, source.id, targetSheetId, drawingIds, pivotIds, tableIds, source.name, targetName)] as const);
  copy.drawingPayloads.clear();
  for (const [id, payload] of payloads) copy.drawingPayloads.set(id, payload);
  copy.drawingGroups.splice(0, copy.drawingGroups.length, ...copy.drawingGroups.map((group) => ({ ...group, id: groupIds.get(group.id)!, sheetId: targetSheetId, memberDrawingIds: group.memberDrawingIds.map((id) => drawingIds.get(id) ?? id) })));
  copy.hyperlinks.forEach((hyperlink, key) => copy.hyperlinks.set(key, { ...hyperlink, target: remapHyperlinkTarget(hyperlink.target, source.id, targetSheetId) }));
  copy.review.reallocateIdentities(targetSheetId, allocateReviewId);
  copy.spillRanges.splice(0, copy.spillRanges.length, ...copy.spillRanges.map((spill) => ({ ...spill, sheetId: targetSheetId, range: mapRange(spill.range, source.id, targetSheetId) })));
  copy.protectionRules.splice(0, copy.protectionRules.length, ...copy.protectionRules.map((rule) => ({ ...rule, sheetId: rule.sheetId ? targetSheetId : undefined, range: rule.range ? mapRange(rule.range, source.id, targetSheetId) : undefined })));
  if (copy.autoFilter) copy.autoFilter = { ...copy.autoFilter, sheetId: targetSheetId, range: mapRange(copy.autoFilter.range, source.id, targetSheetId) } as AutoFilterModel;
  if (copy.bandedRule) copy.bandedRule = { ...copy.bandedRule, range: mapRange(copy.bandedRule.range, source.id, targetSheetId) };
  if (copy.reportSheet) copy.reportSheet = { ...copy.reportSheet, templateSheetId: mapSheetId(copy.reportSheet.templateSheetId, source.id, targetSheetId), tableId: copy.reportSheet.tableId ? tableIds.get(copy.reportSheet.tableId) ?? copy.reportSheet.tableId : undefined };
  if (copy.tableSheet) copy.tableSheet = {
    ...copy.tableSheet,
    columns: copy.tableSheet.columns.map((column) => {
      if (!column.formula) return column;
      return {
        ...column,
        formula: mapFormulaReference(column.formula, source.name, targetName, `duplicate-table-sheet:${source.id}.${column.fieldId}`),
      };
    }),
  };
  copy.cells.forEach((cell, row, column) => {
    const hasBarcodeFormula = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula';
    if (!cell.formula && !cell.formulaMetadata?.sourceFormula && !hasBarcodeFormula) return;
    const next = structuredClone(cell);
    let changed = false;
    if (next.formula) {
      const formula = mapFormula(next.formula, source.name, targetName, `duplicate-cell:${source.id}!${row},${column}`);
      changed ||= formula !== next.formula;
      next.formula = formula;
    }
    if (next.formulaMetadata?.sourceFormula) {
      const sourceFormula = mapFormulaReference(next.formulaMetadata.sourceFormula, source.name, targetName, `duplicate-cell:${source.id}!${row},${column}.sourceFormula`, source.id, next.formulaMetadata.preservedOnly);
      changed ||= sourceFormula !== next.formulaMetadata.sourceFormula;
      next.formulaMetadata.sourceFormula = sourceFormula;
    }
    if (next.presentation?.kind === 'barcode' && next.presentation.source.kind === 'formula') {
      const formula = mapFormulaReference(next.presentation.source.formula, source.name, targetName, `duplicate-cell:${source.id}!${row},${column}.barcode`, source.id);
      changed ||= formula !== next.presentation.source.formula;
      next.presentation.source.formula = formula;
    }
    if (changed) copy.cells.set(row, column, next);
  });
  copy.conditionalFormats.splice(0, copy.conditionalFormats.length, ...copy.conditionalFormats.map((rule) => rewriteRuleFormulas(rule, source.name, targetName)));
  copy.dataValidations.splice(0, copy.dataValidations.length, ...copy.dataValidations.map((rule) => rewriteRuleFormulas(rule, source.name, targetName)));
  return copy;
}

function collectDeletedSheetReferences(workbook: WorkbookModel, sourceSheetId: SheetId, sourceName: string): SheetReferenceInvalidation[] {
  const invalidations: SheetReferenceInvalidation[] = [];
  const deletedSheet = workbook.getSheet(sourceSheetId);
  const deletedPivotIds = new Set(deletedSheet.pivots.map((pivot) => pivot.id));
  const deletedSheetTableIds = new Set(deletedSheet.sheetTables.map((table) => table.id));
  const deletedDrawingIds = new Set(deletedSheet.drawings.map((drawing) => drawing.id));
  const invalidate = (participant: string, ownerSheetId: SheetId | undefined, reference: string): void => {
    invalidations.push({ participant, ownerSheetId, reference, reason: 'deleted-sheet-reference' });
  };
  const inspectRange = (participant: string, ownerSheetId: SheetId | undefined, range: RangeRef): void => {
    if (range.sheetId === sourceSheetId) invalidate(participant, ownerSheetId, JSON.stringify(range));
  };
  const inspectFormulaAnchor = (participant: string, ownerSheetId: SheetId | undefined, anchor: { sheetId: SheetId; row: number; column: number } | undefined): void => {
    if (anchor?.sheetId === sourceSheetId) invalidate(participant, ownerSheetId, JSON.stringify(anchor));
  };
  for (const sheet of workbook.getSheets()) {
    if (sheet.id === sourceSheetId) continue;
    sheet.cells.forEach((cell) => {
      if (cell.formula && formulaReferencesSheet(cell.formula, sourceName, 'cell-formula', sheet.id)) invalidations.push({ participant: 'cell-formula', ownerSheetId: sheet.id, reference: cell.formula, reason: 'deleted-sheet-reference' });
      const barcodeFormula = cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
        ? cell.presentation.source.formula
        : undefined;
      if (barcodeFormula && formulaReferencesSheet(barcodeFormula, sourceName, 'barcode-formula', sheet.id)) {
        invalidate('barcode-formula', sheet.id, barcodeFormula);
      }
      const sourceFormula = cell.formulaMetadata?.sourceFormula;
      if (sourceFormula && formulaReferencesSheet(sourceFormula, sourceName, 'preserved-cell-formula', sheet.id)) {
        invalidate('preserved-cell-formula', sheet.id, sourceFormula);
      }
    });
    for (const column of sheet.tableSheet?.columns ?? []) {
      if (column.formula && formulaReferencesSheet(column.formula, sourceName, `table-sheet:${sheet.id}.${column.fieldId}`, sheet.id)) {
        invalidate('table-sheet-column-formula', sheet.id, `${column.fieldId}:${column.formula}`);
      }
    }
    for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
      inspectFormulaAnchor(`${rule.id}.formulaAnchor`, sheet.id, rule.formulaAnchor);
      if ('listSource' in rule && rule.listSource?.kind === 'range') {
        inspectRange(`${rule.id}.listSource`, sheet.id, rule.listSource.range);
      }
      for (const formula of [
        'value1' in rule && typeof rule.value1 === 'string' ? rule.value1 : undefined,
        'value2' in rule && typeof rule.value2 === 'string' ? rule.value2 : undefined,
        'formula1' in rule ? rule.formula1 : undefined,
        'formula2' in rule ? rule.formula2 : undefined,
        'listSource' in rule && rule.listSource?.kind === 'formula' ? rule.listSource.formula : undefined,
      ]) if (formula && formulaReferencesSheet(formula, sourceName, `${rule.id}.formula`, sheet.id)) invalidations.push({ participant: 'range-rule-formula', ownerSheetId: sheet.id, reference: formula, reason: 'deleted-sheet-reference' });
    }
    const ranges: Array<{ participant: string; range: RangeRef }> = [
      ...sheet.merges.map((entry) => ({ participant: 'merge', range: entry.range })),
      ...sheet.conditionalFormats.flatMap((entry) => entry.ranges.map((range) => ({ participant: 'conditional-format', range }))),
      ...sheet.dataValidations.flatMap((entry) => entry.ranges.map((range) => ({ participant: 'data-validation', range }))),
      ...sheet.dataRegions.map((entry) => ({ participant: 'data-region', range: entry.range })),
      ...sheet.sheetTables.map((entry) => ({ participant: 'table', range: entry.range })),
      ...sheet.spillRanges.map((entry) => ({ participant: 'spill-range', range: entry.range })),
      ...sheet.sparklines.map((entry) => ({ participant: 'sparkline-source', range: entry.sourceRange })),
      ...(sheet.bandedRule ? [{ participant: 'banded-rule', range: sheet.bandedRule.range }] : []),
      ...(sheet.autoFilter ? [{ participant: 'auto-filter', range: sheet.autoFilter.range }] : []),
      ...sheet.protectionRules.flatMap((entry) => entry.range ? [{ participant: 'protection', range: entry.range }] : []),
    ];
    for (const entry of ranges) inspectRange(entry.participant, sheet.id, entry.range);
    for (const drawing of sheet.drawings) {
      if (drawing.sheetId === sourceSheetId) invalidate('drawing-owner', sheet.id, drawing.id);
    }
    for (const group of sheet.drawingGroups) {
      if (group.sheetId === sourceSheetId) invalidate('drawing-group-owner', sheet.id, group.id);
      for (const drawingId of group.memberDrawingIds) {
        if (deletedDrawingIds.has(drawingId)) invalidate('drawing-group-member', sheet.id, `${group.id}:${drawingId}`);
      }
    }
    for (const [payloadId, payload] of sheet.drawingPayloads) {
      const participant = `drawing-payload:${payloadId}`;
      switch (payload.kind) {
        case 'camera':
        case 'screenshot':
          inspectRange(participant, sheet.id, payload.sourceRange);
          break;
        case 'form-control':
          if ('cellLink' in payload && payload.cellLink?.sheetId === sourceSheetId) {
            invalidate(participant, sheet.id, `cellLink:${JSON.stringify(payload.cellLink)}`);
          }
          if ('inputRange' in payload) inspectRange(participant, sheet.id, payload.inputRange);
          break;
        case 'connector':
          for (const drawingId of [payload.start.drawingId, payload.end.drawingId]) {
            if (deletedDrawingIds.has(drawingId)) invalidate(participant, sheet.id, `drawing:${drawingId}`);
          }
          break;
        case 'shape':
          if (payload.hyperlink?.kind === 'sheet' && payload.hyperlink.sheetId === sourceSheetId) {
            invalidate(participant, sheet.id, `hyperlink:${JSON.stringify(payload.hyperlink)}`);
          }
          if (payload.propertyFormula && formulaReferencesSheet(payload.propertyFormula, sourceName, `${participant}.propertyFormula`, sheet.id)) {
            invalidate(`${participant}.propertyFormula`, sheet.id, payload.propertyFormula);
          }
          break;
        case 'chart':
          for (const { field, formula } of chartTextFormulaEntries(payload)) {
            if (formulaReferencesSheet(formula, sourceName, `${participant}.${field}`, sheet.id)) {
              invalidate(`${participant}.${field}`, sheet.id, formula);
            }
          }
          if (payload.source.kind === 'worksheet-ranges') {
            for (const range of payload.source.ranges) inspectRange(`${participant}.source`, sheet.id, range);
          } else if (payload.source.kind === 'report-range') {
            inspectRange(`${participant}.source`, sheet.id, payload.source.range);
          } else if (payload.source.kind === 'pivot' && deletedPivotIds.has(payload.source.pivotId)) {
            invalidate(participant, sheet.id, `pivot:${payload.source.pivotId}`);
          } else if (payload.source.kind === 'table' && deletedSheetTableIds.has(payload.source.tableId)) {
            invalidate(participant, sheet.id, `table:${payload.source.tableId}`);
          }
          if (payload.categoryRange) inspectRange(`${participant}.category`, sheet.id, payload.categoryRange);
          for (const series of payload.series ?? []) {
            for (const [key, range] of [
              ['range', series.range], ['xRange', series.xRange], ['yRange', series.yRange],
              ['sizeRange', series.sizeRange], ['categoryRange', series.categoryRange],
              ['stockOpen', series.stockRoles?.open], ['stockHigh', series.stockRoles?.high],
              ['stockLow', series.stockRoles?.low], ['stockClose', series.stockRoles?.close],
              ['stockVolume', series.stockRoles?.volume], ['errorPlus', series.errorBars?.plusRange],
              ['errorMinus', series.errorBars?.minusRange], ['labelValues', series.dataLabels?.valuesFromCells],
            ] as const) {
              if (range) inspectRange(`${participant}.series.${key}`, sheet.id, range);
            }
          }
          break;
        case 'slicer':
        case 'timeline':
          if (deletedPivotIds.has(payload.pivotId)) invalidate(participant, sheet.id, `pivot:${payload.pivotId}`);
          for (const connection of payload.connections ?? []) {
            if (deletedPivotIds.has(connection.pivotId)) invalidate(participant, sheet.id, `pivot:${connection.pivotId}`);
          }
          break;
        default:
          break;
      }
    }
    if (sheet.reportSheet?.templateSheetId === sourceSheetId) {
      invalidate('report-template-sheet', sheet.id, sheet.reportSheet.templateSheetId);
    }
    for (const pivot of sheet.pivots) {
      if (pivot.target.sheetId === sourceSheetId) invalidate('pivot-target', sheet.id, pivot.id);
      if (pivot.source.kind === 'worksheet-range') inspectRange('pivot-source', sheet.id, pivot.source.range);
      if (pivot.source.kind === 'worksheet-ranges') {
        for (const entry of pivot.source.ranges) inspectRange('pivot-source', sheet.id, entry.range);
      }
      if (pivot.source.kind === 'named-range' && pivot.source.sheetId === sourceSheetId) invalidate('pivot-named-range', sheet.id, pivot.id);
      if (pivot.source.kind === 'table' && deletedSheetTableIds.has(pivot.source.tableId)) invalidate('pivot-table-source', sheet.id, pivot.id);
    }
    for (const hyperlink of sheet.hyperlinks.values()) if (hyperlink.target.kind === 'sheet' && hyperlink.target.sheetId === sourceSheetId) invalidate('hyperlink', sheet.id, hyperlink.id);
  }
  for (const name of workbook.definedNameModels) {
    if (name.scope === 'sheet' && name.sheetId === sourceSheetId) continue;
    inspectFormulaAnchor(`defined-name:${name.name}.anchor`, name.sheetId, name.anchor);
    if (formulaReferencesSheet(name.formula, sourceName, `defined-name:${name.name}`, name.sheetId)) invalidate('defined-name', name.sheetId, name.name);
  }
  for (const source of workbook.dataModel.sources.values()) {
    if (source.sourceSheetId === sourceSheetId || source.sourceRange?.sheetId === sourceSheetId) {
      invalidate('workbook-data-source', undefined, source.id);
    }
  }
  for (const table of workbook.dataModel.tables.values()) {
    if (table.sourceSheetId === sourceSheetId || table.sourceRange?.sheetId === sourceSheetId) {
      invalidate('workbook-table-source', undefined, table.id);
    }
  }
  for (const view of workbook.dataModel.views.values()) {
    for (const field of view.fields) {
      if (field.formula && formulaReferencesSheet(field.formula, sourceName, `data-view:${view.id}.${field.fieldId}`)) {
        invalidate('workbook-data-view-formula', undefined, `${view.id}:${field.fieldId}`);
      }
    }
  }
  for (const query of workbook.queryDefinitions.values()) {
    const target = query.lastTarget;
    if (target?.sheetId === sourceSheetId) invalidate('query-load-target', undefined, query.id);
    if (target?.kind === 'pivot-source' && target.pivotId && deletedPivotIds.has(target.pivotId)) {
      invalidate('query-pivot-target', undefined, query.id);
    }
    if (target?.kind === 'sheet-table' && target.tableId && deletedSheetTableIds.has(target.tableId)) {
      invalidate('query-sheet-table-target', undefined, query.id);
    }
  }
  for (const template of workbook.cellStyleTemplates.values()) {
    const validation = template.dataValidation;
    if (!validation) continue;
    inspectFormulaAnchor(`cell-style-template:${template.id}.formulaAnchor`, undefined, validation.formulaAnchor);
    if (validation.listSource?.kind === 'range') inspectRange(`cell-style-template:${template.id}.listSource`, undefined, validation.listSource.range);
    for (const formula of [
      validation.formula1,
      validation.formula2,
      validation.listSource?.kind === 'formula' ? validation.listSource.formula : undefined,
    ]) {
      if (formula && formulaReferencesSheet(formula, sourceName, `cell-style-template:${template.id}`)) {
        invalidate('cell-style-template-formula', undefined, template.id);
      }
    }
  }
  for (const document of workbook.printDocuments.values()) {
    if (document.sheetId === sourceSheetId) continue;
    if (document.printAreas.some((area) => area.sheetId === sourceSheetId || area.range.sheetId === sourceSheetId)) {
      invalidate('print-document-area', document.sheetId, document.sheetId);
    }
    for (const pageBreak of document.pageBreaks) {
      if (pageBreak.sheetId === sourceSheetId) invalidate('print-document-page-break', document.sheetId, JSON.stringify(pageBreak));
    }
  }
  return invalidations;
}

export function planSheetIdentityTransform(workbook: WorkbookModel, input: SheetIdentityTransformSpec): SheetIdentityTransformPlan {
  const spec = { ...input };
  const source = workbook.getSheet(spec.sourceSheetId);
  if (source.name !== spec.sourceName) throw new SheetIdentityTransformError(`Sheet identity changed before ${spec.kind}: ${spec.sourceSheetId}`);
  if (spec.kind === 'rename') {
    const targetName = spec.targetName?.trim();
    if (!targetName) throw new SheetIdentityTransformError('Sheet rename requires a non-empty targetName');
    const sourceName = source.name;
    const formulaChangePlan = targetName === sourceName
      ? { changes: [], requiresCalculationContextRebuild: false }
      : collectFormulaChanges(workbook, sourceName, targetName);
    const formulaChanges = formulaChangePlan.changes;
    const definedNameResolvesToRenamedSheet = targetName !== sourceName
      && workbook.definedNameModels.some((entry) =>
        formulaReferencesSheet(
          entry.formula,
          targetName,
          `defined-name:${entry.name}`,
          entry.sheetId,
        ));
    const definedNames = targetName === sourceName ? workbook.definedNameModels.map((entry) => structuredClone(entry)) : transformDefinedNames(workbook, sourceName, targetName);
    const conditionalFormatChanges = new Map(workbook.getSheets().map((sheet) => [sheet.id, targetName === sourceName ? structuredClone(sheet.conditionalFormats) : sheet.conditionalFormats.map((rule) => rewriteRuleFormulas(rule, sourceName, targetName))] as const));
    const dataValidationChanges = new Map(workbook.getSheets().map((sheet) => [sheet.id, targetName === sourceName ? structuredClone(sheet.dataValidations) : sheet.dataValidations.map((rule) => rewriteRuleFormulas(rule, sourceName, targetName))] as const));
    const tableSheetChanges = targetName === sourceName ? [] : workbook.getSheets().flatMap((sheet) => sheet.tableSheet ? [{
      sheetId: sheet.id,
      definition: {
        ...structuredClone(sheet.tableSheet),
        columns: sheet.tableSheet.columns.map((column) => ({
          ...column,
          ...(column.formula ? { formula: mapFormulaReference(column.formula, sourceName, targetName, `table-sheet:${sheet.id}.${column.fieldId}`, sheet.id) } : {}),
        })),
      },
    }] : []);
    const dataViewChanges = targetName === sourceName ? [] : [...workbook.dataModel.views.values()].map((view) => ({
      id: view.id,
      view: {
        ...structuredClone(view),
        fields: view.fields.map((field) => ({
          ...field,
          ...(field.formula ? { formula: mapFormulaReference(field.formula, sourceName, targetName, `data-view:${view.id}.${field.fieldId}`) } : {}),
        })),
      },
    }));
    const cellStyleTemplateChanges = targetName === sourceName ? [] : [...workbook.cellStyleTemplates.values()]
      .map((template) => rewriteCellStyleTemplateFormulas(template, sourceName, targetName));
    const drawingPayloadChanges = targetName === sourceName ? [] : workbook.getSheets().flatMap((sheet) => [...sheet.drawingPayloads.entries()].flatMap(([payloadId, payload]) => {
      if (payload.kind === 'shape' && payload.propertyFormula) {
        const propertyFormula = mapFormulaReference(payload.propertyFormula, sourceName, targetName, `drawing:${payloadId}.propertyFormula`, sheet.id);
        return propertyFormula === payload.propertyFormula ? [] : [{ sheetId: sheet.id, payloadId, payload: { ...structuredClone(payload), propertyFormula }, formulaOwnerDeltas: [] as StructuralFormulaObjectOwnerDelta[] }];
      }
      if (payload.kind !== 'chart') return [];
      const next = structuredClone(payload);
      const formulaOwnerDeltas: StructuralFormulaObjectOwnerDelta[] = [];
      for (const { field, formula } of chartTextFormulaEntries(payload)) {
        const afterFormula = mapFormulaReference(formula, sourceName, targetName, `drawing:${payloadId}.${field}`, sheet.id);
        if (afterFormula === formula) continue;
        writeChartTextFormula(next, field, afterFormula);
        formulaOwnerDeltas.push({
          kind: 'formula-object',
          ownerKind: 'chart-text',
          sheetId: sheet.id,
          payloadId,
          field,
          beforeFormula: formula,
          afterFormula,
        });
      }
      return formulaOwnerDeltas.length === 0 ? [] : [{ sheetId: sheet.id, payloadId, payload: next, formulaOwnerDeltas }];
    }));
    const drawingFormulaOwnerDeltas = drawingPayloadChanges.flatMap((change) => change.formulaOwnerDeltas);
    return {
      spec: { ...spec, targetName },
      invalidations: [],
      apply: () => {
        const formulaOwners = formulaChanges.map((change) => {
          const sheet = workbook.getSheet(change.sheetId);
          const cell = sheet.cells.get(change.row, change.column);
          if (!cell) throw new SheetIdentityTransformError(`Formula owner disappeared: ${change.sheetId}!${change.row},${change.column}`);
          if ((change.formula && cell.formula !== change.formula.before)
            || (change.sourceFormula && cell.formulaMetadata?.sourceFormula !== change.sourceFormula.before)
            || (change.barcodeFormula && (cell.presentation?.kind !== 'barcode' || cell.presentation.source.kind !== 'formula' || cell.presentation.source.formula !== change.barcodeFormula.before))) {
            throw new SheetIdentityTransformError(`Formula owner changed before rename: ${change.sheetId}!${change.row},${change.column}`);
          }
          return { change, sheet, cell };
        });
        source.name = targetName;
        for (const { change, sheet, cell } of formulaOwners) {
          const next = { ...cell };
          if (change.formula) next.formula = change.formula.after;
          if (change.sourceFormula && cell.formulaMetadata) {
            next.formulaMetadata = { ...cell.formulaMetadata, sourceFormula: change.sourceFormula.after };
          }
          if (change.barcodeFormula && cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula') {
            next.presentation = { ...cell.presentation, source: { ...cell.presentation.source, formula: change.barcodeFormula.after } };
          }
          sheet.cells.set(change.row, change.column, next);
        }
        workbook.definedNameModels.splice(0, workbook.definedNameModels.length, ...definedNames);
        for (const sheet of workbook.getSheets()) {
          sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...(conditionalFormatChanges.get(sheet.id) ?? []));
          sheet.dataValidations.splice(0, sheet.dataValidations.length, ...(dataValidationChanges.get(sheet.id) ?? []));
        }
        for (const change of tableSheetChanges) workbook.getSheet(change.sheetId).tableSheet = change.definition;
        for (const change of dataViewChanges) workbook.dataModel.views.set(change.id, change.view);
        for (const template of cellStyleTemplateChanges) workbook.cellStyleTemplates.set(template.id, template);
        for (const change of drawingPayloadChanges) workbook.getSheet(change.sheetId).drawingPayloads.set(change.payloadId, change.payload);
        return {
          kind: 'structural-transform',
          removedCells: [],
          clearInputRanges: [],
          populateInputRanges: [],
          rewrittenFormulaOwners: formulaOwners.map(({ change }) => ({
            sheetId: change.sheetId,
            row: change.row,
            column: change.column,
          })),
          ...(drawingFormulaOwnerDeltas.length > 0 ? { formulaOwnerDeltas: drawingFormulaOwnerDeltas } : {}),
          ...(formulaChangePlan.requiresCalculationContextRebuild || definedNameResolvesToRenamedSheet
            ? { calculationContextEffect: CALCULATION_CONTEXT_EFFECTS.rebuild }
            : {}),
        };
      },
    };
  }
  if (spec.kind === 'duplicate') {
    const targetSheetId = spec.targetSheetId?.trim();
    const targetName = spec.targetName?.trim();
    if (!targetSheetId || !targetName) throw new SheetIdentityTransformError('Sheet duplicate requires targetSheetId and targetName');
    if (workbook.sheets.has(targetSheetId)) throw new SheetIdentityTransformError(`Duplicate sheet identity already exists: ${targetSheetId}`);
    return {
      spec: { ...spec, targetSheetId, targetName },
      invalidations: [],
      apply: () => {
        const copy = cloneWorksheetWithIdentity(workbook, source, targetSheetId, targetName);
        workbook.sheets.set(targetSheetId, copy);
        const scopedNames = workbook.definedNameModels.filter((entry) => entry.scope === 'sheet' && entry.sheetId === source.id).map((entry) => ({
          ...structuredClone(entry),
          sheetId: targetSheetId,
          formula: mapFormula(entry.formula, source.name, targetName, `defined-name:${entry.name}`),
          anchor: entry.anchor ? { ...entry.anchor, sheetId: mapSheetId(entry.anchor.sheetId, source.id, targetSheetId) } : undefined,
        }));
        workbook.definedNameModels.push(...scopedNames);
        const printDocument = workbook.printDocuments.get(source.id);
        if (printDocument) workbook.printDocuments.set(targetSheetId, { ...structuredClone(printDocument), sheetId: targetSheetId, printAreas: printDocument.printAreas.map((area) => ({ ...area, sheetId: targetSheetId, range: mapRange(area.range, source.id, targetSheetId) })), pageBreaks: printDocument.pageBreaks.map((item) => ({ ...item, sheetId: targetSheetId })) });
        const sourceIndex = workbook.sheetOrder.indexOf(source.id);
        workbook.sheetOrder.splice(sourceIndex + 1, 0, targetSheetId);
      },
    };
  }
  const invalidations = collectDeletedSheetReferences(workbook, source.id, source.name);
  if (invalidations.length > 0) throw new SheetIdentityTransformError(`Cannot delete sheet ${source.id}; external references must be resolved first`, invalidations);
  return {
    spec,
    invalidations: [],
    apply: () => {
      if (workbook.sheets.size <= 1) throw new SheetIdentityTransformError('A workbook must keep at least one worksheet');
      workbook.sheets.delete(source.id);
      workbook.sheetOrder = workbook.sheetOrder.filter((id) => id !== source.id);
      workbook.printDocuments.delete(source.id);
      for (let index = workbook.definedNameModels.length - 1; index >= 0; index -= 1) {
        const entry = workbook.definedNameModels[index];
        if (entry?.scope === 'sheet' && entry.sheetId === source.id) workbook.definedNameModels.splice(index, 1);
      }
    },
  };
}

export function assertNoDanglingSheetReferences(workbook: WorkbookModel): void {
  const known = new Set(workbook.sheetOrder);
  for (const sheet of workbook.getSheets()) {
    const ranges: RangeRef[] = [
      ...sheet.merges.map((entry) => entry.range),
      ...sheet.conditionalFormats.flatMap((entry) => entry.ranges),
      ...sheet.dataValidations.flatMap((entry) => entry.ranges),
      ...sheet.spillRanges.map((entry) => entry.range),
      ...sheet.protectionRules.flatMap((entry) => entry.range ? [entry.range] : []),
    ];
    for (const range of ranges) if (!known.has(range.sheetId)) throw new SheetIdentityTransformError(`Dangling range reference: ${range.sheetId}`);
    for (const hyperlink of sheet.hyperlinks.values()) if (hyperlink.target.kind === 'sheet' && !known.has(hyperlink.target.sheetId)) throw new SheetIdentityTransformError(`Dangling hyperlink reference: ${hyperlink.target.sheetId}`);
  }
}
