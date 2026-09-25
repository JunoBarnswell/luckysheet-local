import type { CellAddress, CellData, CellStyleTemplate, ConditionalFormatRule, DataValidationRule, RangeRef, Row, WorkbookModel, WorksheetModel } from './index';
import { cellKey, hasFormulaGroupMetadata } from './index';
import type { DefinedNameModel, DrawingObject, SpillRange } from './domain';
import { sheetRuleRegistry, type RuleTransform } from './rule-lifecycle';
import { mapReportSheetCoordinates } from './report-sheet-transform';
import { formatFormula, MAX_COLUMN_INDEX, MAX_ROW_INDEX, offsetAst, parseFormula } from '@react-sheets/formula-engine';

/** Canonical, prevalidated permutation shared by local execution and replay. */
export interface RowPermutationPlan {
  readonly range: RangeRef;
  readonly metadataScope: RangeRef;
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

export function createRowPermutationPlan(range: RangeRef, sourceRows: readonly Row[], affectedColumnEnd: number): RowPermutationPlan {
  const normalized = normalizeRange(range);
  if (!Number.isSafeInteger(affectedColumnEnd) || affectedColumnEnd < normalized.endColumn || affectedColumnEnd > MAX_COLUMN_INDEX) {
    throw new Error('Row permutation metadata extent is outside worksheet bounds');
  }
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
  const metadataScope = {
    sheetId: normalized.sheetId,
    startRow: normalized.startRow,
    endRow: normalized.endRow,
    startColumn: 0,
    endColumn: affectedColumnEnd,
  };
  return Object.freeze({ range: Object.freeze(normalized), metadataScope: Object.freeze(metadataScope), sourceRows: Object.freeze([...sourceRows]), sourceToTarget });
}

function inRange(range: RangeRef, row: number, column: number): boolean {
  return range.startRow <= row && row <= range.endRow && range.startColumn <= column && column <= range.endColumn;
}

/** A table's header/total rows are not part of a data-body sort permutation. */
function isTableBodyPermutation(table: WorksheetModel['sheetTables'][number], range: RangeRef): boolean {
  const bodyEnd = table.range.endRow - (table.hasTotalRow ? 1 : 0);
  return table.hasHeaderRow
    && range.startRow === table.range.startRow + 1
    && range.endRow === bodyEnd
    && range.startColumn === table.range.startColumn
    && range.endColumn === table.range.endColumn;
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
function remapRangeExact(range: RangeRef, plan: RowPermutationPlan, scope = plan.range): RangeRef[] {
  if (range.sheetId !== scope.sheetId || !rangesIntersect(range, scope)) return [structuredClone(range)];
  const selected = scope;
  const firstRow = Math.max(range.startRow, selected.startRow);
  const lastRow = Math.min(range.endRow, selected.endRow);
  const firstColumn = Math.max(range.startColumn, selected.startColumn);
  const lastColumn = Math.min(range.endColumn, selected.endColumn);
  const result: RangeRef[] = [];
  if (range.startRow < firstRow) result.push(cloneRange(range, range.startRow, firstRow - 1));
  if (lastRow < range.endRow) result.push(cloneRange(range, lastRow + 1, range.endRow));
  if (range.startColumn < firstColumn) result.push(cloneRange(range, firstRow, lastRow, range.startColumn, firstColumn - 1));
  if (lastColumn < range.endColumn) result.push(cloneRange(range, firstRow, lastRow, lastColumn + 1, range.endColumn));
  const affectedCells = (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
  if (!Number.isSafeInteger(affectedCells) || affectedCells > MAX_SEGMENT_CELLS) throw new Error('Row permutation metadata range cannot be represented exactly within the bounded plan');
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

function ruleTransformForPlan(plan: RowPermutationPlan): RuleTransform {
  return {
    mapRange: (range) => remapRangeExact(range, plan, plan.metadataScope),
    mapAddress: (address) => ({
      ...address,
      row: address.sheetId === plan.metadataScope.sheetId && inRange(plan.metadataScope, address.row, address.column)
        ? remapRow(address.row, plan)
        : address.row,
    }),
  };
}

function remapSingleRange(owner: string, range: RangeRef, plan: RowPermutationPlan, scope = plan.range): RangeRef {
  const segments = remapRangeExact(range, plan, scope);
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
    const nextKey = cellKey(inRange(plan.range, row, column) ? remapRow(row, plan) : row, column);
    if (next.has(nextKey)) throw new Error(`Sort produced duplicate cell metadata at ${nextKey}`);
    next.set(nextKey, value);
  }
  return next;
}

export function rowPermutationAffectedColumnEnd(workbook: WorkbookModel, range: RangeRef): number {
  const sheet = workbook.getSheet(range.sheetId);
  let end = sheetRuleRegistry.affectedColumnEnd(sheet, range.endColumn);
  const includeAnchor = (anchor: CellAddress | undefined): void => {
    if (!anchor || anchor.sheetId !== sheet.id) return;
    if (!Number.isSafeInteger(anchor.row) || anchor.row < 0 || anchor.row > MAX_ROW_INDEX
      || !Number.isSafeInteger(anchor.column) || anchor.column < 0 || anchor.column > MAX_COLUMN_INDEX) {
      throw new Error('Row permutation formula anchor is outside worksheet bounds');
    }
    if (anchor.row >= range.startRow && anchor.row <= range.endRow) {
      end = Math.max(end, anchor.column);
    }
  };
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) includeAnchor(rule.formulaAnchor);
  for (const name of workbook.definedNameModels) includeAnchor(name.anchor);
  for (const template of workbook.cellStyleTemplates.values()) includeAnchor(template.dataValidation?.formulaAnchor);
  return end;
}

type PermutationFormulaFields = {
  operator?: string;
  value1?: string | number;
  value2?: string | number;
  type?: string;
  formula1?: string;
  formula2?: string;
  listSource?: DataValidationRule['listSource'];
};

function offsetPermutationFormula(formula: string, rowDelta: number, owner: string): string {
  try {
    const hasPrefix = formula.startsWith('=');
    const ast = parseFormula(hasPrefix ? formula : `=${formula}`);
    if (containsUnsupportedRowOffsetReference(ast)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: row sort cannot safely offset an external-workbook or whole-row reference in ${owner}`);
    }
    const shifted = offsetAst(ast, rowDelta, 0);
    if (countInvalidReferenceNodes(shifted) > countInvalidReferenceNodes(ast)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: row sort would move a formula reference outside worksheet bounds in ${owner}`);
    }
    const formatted = formatFormula(shifted);
    return hasPrefix ? formatted : formatted.slice(1);
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('UNSUPPORTED_STRUCTURAL_REFERENCE:')) throw error;
    throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: row sort cannot parse and safely offset formula in ${owner}`);
  }
}

function offsetPermutationFormulaFields(owner: PermutationFormulaFields, rowDelta: number, identity: string): void {
  const remap = (formula: string): string => offsetPermutationFormula(formula, rowDelta, identity);
  const formulaOperator = owner.operator === 'formula';
  if (typeof owner.value1 === 'string' && (formulaOperator || owner.value1.trim().startsWith('='))) owner.value1 = remap(owner.value1);
  else {
    if (typeof owner.value1 === 'string' && owner.value1.trim().startsWith('=')) owner.value1 = remap(owner.value1);
    if (typeof owner.value2 === 'string' && owner.value2.trim().startsWith('=')) owner.value2 = remap(owner.value2);
  }
  if (typeof owner.formula1 === 'string' && owner.formula1
    && (owner.formula1.trim().startsWith('=') || owner.operator === 'formula' || owner.type === 'custom')) {
    owner.formula1 = remap(owner.formula1);
  }
  if (typeof owner.formula2 === 'string' && owner.formula2
    && (owner.formula2.trim().startsWith('=') || owner.type === 'custom')) {
    owner.formula2 = remap(owner.formula2);
  }
  if (owner.listSource?.kind === 'formula') owner.listSource.formula = remap(owner.listSource.formula);
}

function remapRuleForPermutation<T extends ConditionalFormatRule | DataValidationRule>(rule: T, transform: RuleTransform, plan: RowPermutationPlan, changesRows: boolean): T {
  const next = sheetRuleRegistry.transform(rule, transform);
  const firstRange = rule.ranges[0];
  if (!firstRange && !rule.formulaAnchor) throw new Error(`Row permutation rule ${rule.id} has no formula anchor`);
  const oldAnchor = rule.formulaAnchor ?? { sheetId: firstRange!.sheetId, row: firstRange!.startRow, column: firstRange!.startColumn };
  const mappedAnchor = transform.mapAddress(oldAnchor);
  const rowDelta = mappedAnchor.row - oldAnchor.row;
  if (rule.formulaAnchor === undefined && changesRows && oldAnchor.sheetId === plan.metadataScope.sheetId
    && inRange(plan.metadataScope, oldAnchor.row, oldAnchor.column)) {
    next.formulaAnchor = mappedAnchor;
  }
  if (rowDelta !== 0) {
    next.formulaAnchor = mappedAnchor;
    offsetPermutationFormulaFields(next, rowDelta, `rule ${rule.id}`);
  }
  return next;
}

interface RowPermutationOwnerChanges {
  readonly conditionalFormats: ConditionalFormatRule[];
  readonly dataValidations: DataValidationRule[];
  readonly definedNames: Array<{ entry: DefinedNameModel; formula: string; anchor: DefinedNameModel['anchor'] }>;
  readonly templates: CellStyleTemplate[];
  readonly reportSheet?: WorksheetModel['reportSheet'];
}

/** Validate all owners before the first cell changes. */
export function validatePermutationMetadata(workbook: WorkbookModel, plan: RowPermutationPlan): RowPermutationOwnerChanges {
  const sheet = workbook.getSheet(plan.range.sheetId);
  const range = plan.range;
  if (range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw new Error('Row permutation range is outside worksheet bounds');
  if (plan.metadataScope.sheetId !== range.sheetId || plan.metadataScope.startColumn !== 0 || plan.metadataScope.startRow !== range.startRow
    || plan.metadataScope.endRow !== range.endRow || plan.metadataScope.endColumn > MAX_COLUMN_INDEX
    || plan.metadataScope.endColumn < rowPermutationAffectedColumnEnd(workbook, range)) {
    throw new Error('Row permutation metadata scope does not cover its canonical owners');
  }
  const changesRows = plan.sourceRows.some((sourceRow, targetOffset) => sourceRow !== range.startRow + targetOffset);
  if (changesRows) {
    sheet.cells.forEachInRows(new Set(plan.sourceRows), (cell, row, column) => {
      if (column < range.startColumn || column > range.endColumn) return;
      if (hasFormulaGroupMetadata(cell)) {
        throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: row sort cannot remap formula-group metadata at ${sheet.id}!${row}:${column}`);
      }
    });
  }
  // Detect cell-owner collisions before any cell record is cleared. Notes and
  // hyperlinks are single-owner maps, so a collision is an atomic rejection.
  sheet.review.validateRemapCoordinates((row, column) => ({ row: inRange(plan.range, row, column) ? remapRow(row, plan) : row, column }));
  remapCellMap(sheet.hyperlinks, plan);
  for (const merge of sheet.merges) {
    if (!rangesIntersect(merge.range, range)) continue;
    if (!(merge.range.startRow >= range.startRow && merge.range.endRow <= range.endRow)) throw new Error('Sort cannot partially intersect a merged range');
    remapSingleRange('merge', merge.range, plan);
  }
  for (const table of sheet.sheetTables) {
    if (rangesIntersect(table.range, range)) {
      const completeTable = table.range.startRow === range.startRow && table.range.endRow === range.endRow;
      if (!completeTable && !isTableBodyPermutation(table, range)) throw new Error('Sort requires the complete table row range or its data body');
      if (completeTable) remapSingleRange(`table ${table.id}`, table.range, plan);
    }
    if (table.autoFilter && !isTableBodyPermutation(table, range)) remapSingleRange(`table ${table.id} filter`, table.autoFilter.range, plan);
  }
  for (const group of sheet.outline?.groups ?? []) {
    if (group.axis !== 'row' || group.start > range.endRow || group.end < range.startRow) continue;
    if (group.start < range.startRow || group.end > range.endRow) throw new Error('Sort cannot partially intersect an outline group');
    const groupRange = { sheetId: sheet.id, startRow: group.start, endRow: group.end, startColumn: range.startColumn, endColumn: range.endColumn };
    if (remapRangeExact(groupRange, plan).length !== 1) throw new Error('Sort cannot exactly remap an outline group');
  }
  for (const drawing of sheet.drawings) remapDrawingAnchor(drawing, plan);
  for (const sparkline of sheet.sparklines) remapSingleRange(`sparkline ${sparkline.id}`, sparkline.sourceRange, plan);
  for (const spill of sheet.spillRanges) remapSpill(spill, plan);
  const ruleTransform = ruleTransformForPlan(plan);
  const conditionalFormats = sheet.conditionalFormats.map((rule) => remapRuleForPermutation(rule, ruleTransform, plan, changesRows));
  const dataValidations = sheet.dataValidations.map((rule) => remapRuleForPermutation(rule, ruleTransform, plan, changesRows));
  if (sheet.autoFilter) remapSingleRange('auto filter', sheet.autoFilter.range, plan);
  for (const pivot of sheet.pivots) {
    if (pivot.source.kind === 'worksheet-range') remapSingleRange(`pivot ${pivot.id} source`, pivot.source.range, plan);
    if (pivot.source.kind === 'worksheet-ranges') for (const source of pivot.source.ranges) remapSingleRange(`pivot ${pivot.id} source`, source.range, plan);
  }
  for (const rule of sheet.protectionRules) if (rule.range) remapSingleRange(`protection ${rule.id}`, rule.range, plan, plan.metadataScope);
  if (sheet.bandedRule) remapSingleRange('banded rule', sheet.bandedRule.range, plan);
  const definedNames = workbook.definedNameModels.flatMap((entry) => {
    const anchor = entry.anchor;
    if (!anchor || anchor.sheetId !== sheet.id || !inRange(plan.metadataScope, anchor.row, anchor.column)) return [];
    const mappedAnchor = { ...anchor, row: remapRow(anchor.row, plan) };
    const rowDelta = mappedAnchor.row - anchor.row;
    return rowDelta === 0 ? [] : [{ entry, formula: offsetPermutationFormula(entry.formula, rowDelta, `defined name ${entry.name}`), anchor: mappedAnchor }];
  });
  const templates = [...workbook.cellStyleTemplates.values()].flatMap((template) => {
    const validation = template.dataValidation;
    const anchor = validation?.formulaAnchor;
    if (!validation || !anchor || anchor.sheetId !== sheet.id || !inRange(plan.metadataScope, anchor.row, anchor.column)) return [];
    const mappedAnchor = { ...anchor, row: remapRow(anchor.row, plan) };
    const rowDelta = mappedAnchor.row - anchor.row;
    if (rowDelta === 0) return [];
    const next = structuredClone(template);
    next.dataValidation!.formulaAnchor = mappedAnchor;
    offsetPermutationFormulaFields(next.dataValidation!, rowDelta, `cell-style template ${template.id}`);
    return [next];
  });
  const reportSheet = sheet.reportSheet
    ? mapReportSheetCoordinates(
      sheet.reportSheet,
      (cell) => inRange(plan.metadataScope, cell.row, cell.column)
        ? { ...cell, row: remapRow(cell.row, plan) }
        : { ...cell },
      (row) => row >= plan.range.startRow && row <= plan.range.endRow ? remapRow(row, plan) : row,
      'row-permutation',
    )
    : undefined;
  return { conditionalFormats, dataValidations, definedNames, templates, reportSheet };
}

export function applyRowPermutation(workbook: WorkbookModel, plan: RowPermutationPlan): void {
  const sheet = workbook.getSheet(plan.range.sheetId);
  const ownerChanges = validatePermutationMetadata(workbook, plan);
  const { range, sourceRows } = plan;
  const cellsByRow = new Map<number, Array<{ column: number; cell: CellData }>>();
  sheet.cells.forEachInRows(new Set(sourceRows), (cell, row, column) => {
    if (column < range.startColumn || column > range.endColumn) return;
    const targetRow = plan.sourceToTarget.get(row);
    if (targetRow === undefined) throw new Error(`ROW_PERMUTATION_INVARIANT: cell owner row ${row} is outside its source map`);
    const rowDelta = targetRow - row;
    const nextCell = rowDelta === 0 ? structuredClone(cell) : remapPermutedFormulaOwner(cell, rowDelta, sheet.id, row, column);
    const entries = cellsByRow.get(row) ?? [];
    entries.push({ column, cell: nextCell });
    cellsByRow.set(row, entries);
  });
  for (const [row, entries] of cellsByRow) for (const entry of entries) sheet.cells.delete(row, entry.column);
  sourceRows.forEach((sourceRow, targetOffset) => { for (const entry of cellsByRow.get(sourceRow) ?? []) sheet.cells.set(range.startRow + targetOffset, entry.column, entry.cell); });

  sheet.review.remapCoordinates((row, column) => ({ row: inRange(plan.range, row, column) ? remapRow(row, plan) : row, column }));
  const hyperlinks = remapCellMap(sheet.hyperlinks, plan); sheet.hyperlinks.clear(); for (const [key, value] of hyperlinks) sheet.hyperlinks.set(key, value);
  for (const drawing of sheet.drawings) Object.assign(drawing, remapDrawingAnchor(drawing, plan));
  for (const sparkline of sheet.sparklines) { sparkline.sourceRange = remapSingleRange(`sparkline ${sparkline.id}`, sparkline.sourceRange, plan); if (inRange(range, sparkline.anchor.row, sparkline.anchor.column)) sparkline.anchor.row = remapRow(sparkline.anchor.row, plan); }
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...sheet.spillRanges.map((spill) => remapSpill(spill, plan)));
  sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...ownerChanges.conditionalFormats);
  sheet.dataValidations.splice(0, sheet.dataValidations.length, ...ownerChanges.dataValidations);
  if (sheet.autoFilter) sheet.autoFilter.range = remapSingleRange('auto filter', sheet.autoFilter.range, plan);
  for (const table of sheet.sheetTables) {
    const bodyPermutation = isTableBodyPermutation(table, range);
    if (!bodyPermutation) {
      table.range = remapSingleRange(`table ${table.id}`, table.range, plan);
      if (table.autoFilter) table.autoFilter.range = remapSingleRange(`table ${table.id} filter`, table.autoFilter.range, plan);
    }
  }
  for (const pivot of sheet.pivots) { if (pivot.source.kind === 'worksheet-range') pivot.source.range = remapSingleRange(`pivot ${pivot.id} source`, pivot.source.range, plan); if (pivot.source.kind === 'worksheet-ranges') for (const source of pivot.source.ranges) source.range = remapSingleRange(`pivot ${pivot.id} source`, source.range, plan); if (pivot.target.sheetId === sheet.id && inRange(range, pivot.target.anchor.row, pivot.target.anchor.column)) pivot.target.anchor.row = remapRow(pivot.target.anchor.row, plan); }
  for (const merge of sheet.merges) { merge.range = remapSingleRange('merge', merge.range, plan); if (inRange(range, merge.anchor.row, merge.anchor.column)) merge.anchor.row = remapRow(merge.anchor.row, plan); }
  for (const group of sheet.outline?.groups ?? []) if (group.axis === 'row' && group.start >= range.startRow && group.end <= range.endRow) { const mapped = remapRangeExact({ sheetId: sheet.id, startRow: group.start, endRow: group.end, startColumn: range.startColumn, endColumn: range.endColumn }, plan); if (mapped.length !== 1) throw new Error('Sort cannot exactly remap outline group'); group.start = mapped[0]!.startRow; group.end = mapped[0]!.endRow; }
  for (const rule of sheet.protectionRules) if (rule.range) rule.range = remapSingleRange(`protection ${rule.id}`, rule.range, plan, plan.metadataScope);
  if (sheet.bandedRule) sheet.bandedRule.range = remapSingleRange('banded rule', sheet.bandedRule.range, plan);
  for (const change of ownerChanges.definedNames) {
    change.entry.formula = change.formula;
    change.entry.anchor = change.anchor;
  }
  for (const template of ownerChanges.templates) workbook.cellStyleTemplates.set(template.id, template);
  if (ownerChanges.reportSheet) sheet.reportSheet = ownerChanges.reportSheet;
}

function remapPermutedFormulaOwner(cell: CellData, rowDelta: number, sheetId: string, row: number, column: number): CellData {
  const next = structuredClone(cell);
  const remap = (formula: string): string => offsetPermutationFormula(formula, rowDelta, `${sheetId}!${row}:${column}`);
  if (next.formula !== undefined) {
    next.formula = remap(next.formula);
    delete next.formulaValue;
  }
  if (next.formulaMetadata?.sourceFormula !== undefined) {
    next.formulaMetadata = { ...next.formulaMetadata, sourceFormula: remap(next.formulaMetadata.sourceFormula) };
  }
  if (next.presentation?.kind === 'barcode' && next.presentation.source.kind === 'formula') {
    next.presentation = { ...next.presentation, source: { ...next.presentation.source, formula: remap(next.presentation.source.formula) } };
  }
  return next;
}

function containsUnsupportedRowOffsetReference(value: unknown): boolean {
  if (Array.isArray(value)) return value.some(containsUnsupportedRowOffsetReference);
  if (!value || typeof value !== 'object') return false;
  const candidate = value as { readonly type?: unknown };
  if (candidate.type === 'external-reference' || candidate.type === 'whole-row-reference') return true;
  return Object.values(value).some(containsUnsupportedRowOffsetReference);
}

function countInvalidReferenceNodes(value: unknown): number {
  if (Array.isArray(value)) return value.reduce((count, child) => count + countInvalidReferenceNodes(child), 0);
  if (!value || typeof value !== 'object') return 0;
  const candidate = value as { readonly type?: unknown };
  const ownInvalidReference = candidate.type === 'invalid-reference' ? 1 : 0;
  return ownInvalidReference + Object.values(value).reduce<number>((count, child) => count + countInvalidReferenceNodes(child), 0);
}
