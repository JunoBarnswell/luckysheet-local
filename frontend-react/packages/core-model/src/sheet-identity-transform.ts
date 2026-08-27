import {
  formatFormula,
  parseFormula,
  renameAstSheetReferences,
} from '@react-sheets/formula-engine';
import type {
  RangeRef,
  SheetId,
  WorksheetModel,
  WorkbookModel,
  AutoFilterModel,
  ConditionalFormatRule,
  DataValidationRule,
} from './index';
import type {
  DrawingPayload,
  DefinedNameModel,
  HyperlinkTarget,
} from './domain';
import type { PivotModel, PivotSource } from './pivot';

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

export interface SheetIdentityTransformPlan {
  readonly spec: Readonly<SheetIdentityTransformSpec>;
  readonly invalidations: readonly SheetReferenceInvalidation[];
  apply(): void;
}

type FormulaChange = { sheetId: SheetId; row: number; column: number; formula: string };

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

function collectFormulaChanges(workbook: WorkbookModel, oldName: string, newName: string): FormulaChange[] {
  const changes: FormulaChange[] = [];
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      if (!cell.formula) return;
      const formula = mapFormula(cell.formula, oldName, newName, `cell:${sheet.id}!${row},${column}`);
      if (formula !== cell.formula) changes.push({ sheetId: sheet.id, row, column, formula });
    });
  }
  return changes;
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
    case 'data-chart':
      next.source = next.source.kind === 'report-sheet'
        ? { ...next.source, range: mapRange(next.source.range, sourceSheetId, targetSheetId) }
        : { ...next.source, tableId: tableIds.get(next.source.tableId) ?? next.source.tableId };
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
      break;
    case 'chart':
      next.sourceRanges = next.sourceRanges.map((range) => mapRange(range, sourceSheetId, targetSheetId));
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
  const payloads = [...copy.drawingPayloads.entries()].map(([id, payload]) => [payloadIds.get(id) ?? id, remapDrawingPayload(payload, source.id, targetSheetId, drawingIds, pivotIds, tableIds)] as const);
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
  copy.cells.forEach((cell, row, column) => {
    if (!cell.formula) return;
    const formula = mapFormula(cell.formula, source.name, targetName, `duplicate-cell:${source.id}!${row},${column}`);
    if (formula !== cell.formula) copy.cells.set(row, column, { ...cell, formula });
  });
  copy.conditionalFormats.splice(0, copy.conditionalFormats.length, ...copy.conditionalFormats.map((rule) => rewriteRuleFormulas(rule, source.name, targetName)));
  copy.dataValidations.splice(0, copy.dataValidations.length, ...copy.dataValidations.map((rule) => rewriteRuleFormulas(rule, source.name, targetName)));
  return copy;
}

function collectDeletedSheetReferences(workbook: WorkbookModel, sourceSheetId: SheetId, sourceName: string): SheetReferenceInvalidation[] {
  const invalidations: SheetReferenceInvalidation[] = [];
  for (const sheet of workbook.getSheets()) {
    if (sheet.id === sourceSheetId) continue;
    sheet.cells.forEach((cell) => {
      if (cell.formula && formulaReferencesSheet(cell.formula, sourceName, 'cell-formula', sheet.id)) invalidations.push({ participant: 'cell-formula', ownerSheetId: sheet.id, reference: cell.formula, reason: 'deleted-sheet-reference' });
    });
    for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
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
      ...sheet.protectionRules.flatMap((entry) => entry.range ? [{ participant: 'protection', range: entry.range }] : []),
    ];
    for (const entry of ranges) if (entry.range.sheetId === sourceSheetId) invalidations.push({ participant: entry.participant, ownerSheetId: sheet.id, reference: JSON.stringify(entry.range), reason: 'deleted-sheet-reference' });
    for (const pivot of sheet.pivots) {
      if (pivot.target.sheetId === sourceSheetId) invalidations.push({ participant: 'pivot-target', ownerSheetId: sheet.id, reference: pivot.id, reason: 'deleted-sheet-reference' });
      if (pivot.source.kind === 'worksheet-range' && pivot.source.range.sheetId === sourceSheetId) invalidations.push({ participant: 'pivot-source', ownerSheetId: sheet.id, reference: pivot.id, reason: 'deleted-sheet-reference' });
      if (pivot.source.kind === 'worksheet-ranges' && pivot.source.ranges.some((entry) => entry.range.sheetId === sourceSheetId)) invalidations.push({ participant: 'pivot-source', ownerSheetId: sheet.id, reference: pivot.id, reason: 'deleted-sheet-reference' });
      if (pivot.source.kind === 'named-range' && pivot.source.sheetId === sourceSheetId) invalidations.push({ participant: 'pivot-named-range', ownerSheetId: sheet.id, reference: pivot.id, reason: 'deleted-sheet-reference' });
    }
    for (const hyperlink of sheet.hyperlinks.values()) if (hyperlink.target.kind === 'sheet' && hyperlink.target.sheetId === sourceSheetId) invalidations.push({ participant: 'hyperlink', ownerSheetId: sheet.id, reference: hyperlink.id, reason: 'deleted-sheet-reference' });
  }
  for (const name of workbook.definedNameModels) {
    if (name.scope === 'sheet' && name.sheetId === sourceSheetId) continue;
    if (formulaReferencesSheet(name.formula, sourceName, `defined-name:${name.name}`, name.sheetId)) invalidations.push({ participant: 'defined-name', ownerSheetId: name.sheetId, reference: name.name, reason: 'deleted-sheet-reference' });
  }
  for (const document of workbook.printDocuments.values()) {
    if (document.sheetId === sourceSheetId) continue;
    if (document.printAreas.some((area) => area.range.sheetId === sourceSheetId)) invalidations.push({ participant: 'print-document', ownerSheetId: document.sheetId, reference: document.sheetId, reason: 'deleted-sheet-reference' });
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
    const formulaChanges = targetName === sourceName ? [] : collectFormulaChanges(workbook, sourceName, targetName);
    const definedNames = targetName === sourceName ? workbook.definedNameModels.map((entry) => structuredClone(entry)) : transformDefinedNames(workbook, sourceName, targetName);
    const conditionalFormatChanges = new Map(workbook.getSheets().map((sheet) => [sheet.id, targetName === sourceName ? structuredClone(sheet.conditionalFormats) : sheet.conditionalFormats.map((rule) => rewriteRuleFormulas(rule, sourceName, targetName))] as const));
    const dataValidationChanges = new Map(workbook.getSheets().map((sheet) => [sheet.id, targetName === sourceName ? structuredClone(sheet.dataValidations) : sheet.dataValidations.map((rule) => rewriteRuleFormulas(rule, sourceName, targetName))] as const));
    return {
      spec: { ...spec, targetName },
      invalidations: [],
      apply: () => {
        source.name = targetName;
        for (const change of formulaChanges) {
          const cell = workbook.getSheet(change.sheetId).cells.get(change.row, change.column);
          if (!cell) throw new SheetIdentityTransformError(`Formula owner disappeared: ${change.sheetId}!${change.row},${change.column}`);
          workbook.getSheet(change.sheetId).cells.set(change.row, change.column, { ...cell, formula: change.formula });
        }
        workbook.definedNameModels.splice(0, workbook.definedNameModels.length, ...definedNames);
        for (const sheet of workbook.getSheets()) {
          sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...(conditionalFormatChanges.get(sheet.id) ?? []));
          sheet.dataValidations.splice(0, sheet.dataValidations.length, ...(dataValidationChanges.get(sheet.id) ?? []));
        }
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
        const scopedNames = workbook.definedNameModels.filter((entry) => entry.scope === 'sheet' && entry.sheetId === source.id).map((entry) => ({ ...structuredClone(entry), sheetId: targetSheetId, formula: mapFormula(entry.formula, source.name, targetName, `defined-name:${entry.name}`) }));
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
