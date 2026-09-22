import type {
  PivotAggregateFunction,
  PivotFieldCatalog,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotShowAs,
  PivotSourceRowPath,
  PivotValueField,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { allowsMultiplePivotFilters, pivotFilterIdentity } from '@react-sheets/core-model';
import { PIVOT_RESULT_TREE_SCHEMA } from '@react-sheets/core-model';
import { getPivotSourceRanges, normalizePivotDefinitionFromCatalog } from './engine';

export interface PivotPanelState {
  pivotId: string;
  sheetId: string;
  fieldCatalog: PivotFieldCatalog;
  layout: PivotLayout;
  resultTreeSchema: string;
}

export type PivotLayoutArea = 'filters' | 'columns' | 'rows' | 'values';

function boundedInsertIndex(index: number, length: number): number {
  if (!Number.isSafeInteger(index)) throw new Error('Pivot field insertion index is invalid');
  return Math.max(0, Math.min(index, length));
}

function rescopePivotFieldFilters(layout: PivotLayout, fieldId: string, from: 'report' | 'field', to: 'report' | 'field'): PivotFilter[] {
  const filters = layout.filters.map((filter) => filter.fieldId === fieldId && (filter.scope ?? 'report') === from
    ? { ...filter, scope: to }
    : structuredClone(filter)) as PivotFilter[];
  const identities = new Set<string>();
  const fieldScopes = new Set<string>();
  for (const filter of filters) {
    const identity = pivotFilterIdentity(filter);
    if (identities.has(identity)) throw new Error(`Pivot field move would merge duplicate filter family: ${identity}`);
    identities.add(identity);
    const fieldScope = `${filter.fieldId}|${filter.scope ?? 'report'}`;
    if (!allowsMultiplePivotFilters(layout) && fieldScopes.has(fieldScope)) {
      throw new Error(`Pivot field move would create multiple filters while multiple filters are disabled: ${fieldScope}`);
    }
    fieldScopes.add(fieldScope);
  }
  return filters;
}

function reorderPivotReportFilterField(filters: PivotFilter[], fieldId: string, index: number): PivotFilter[] {
  const grouped = new Map<string, PivotFilter[]>();
  const order: string[] = [];
  for (const filter of filters) {
    if ((filter.scope ?? 'report') !== 'report') continue;
    const existing = grouped.get(filter.fieldId);
    if (existing) existing.push(filter);
    else {
      grouped.set(filter.fieldId, [filter]);
      order.push(filter.fieldId);
    }
  }
  const previousIndex = order.indexOf(fieldId);
  if (previousIndex < 0) throw new Error(`Pivot Filters placement is missing: ${fieldId}`);
  order.splice(previousIndex, 1);
  const adjustedIndex = previousIndex < index ? index - 1 : index;
  order.splice(boundedInsertIndex(adjustedIndex, order.length), 0, fieldId);
  const orderedReports = order.flatMap((id) => grouped.get(id)!);
  let reportIndex = 0;
  return filters.map((filter) => (filter.scope ?? 'report') === 'report' ? orderedReports[reportIndex++]! : filter);
}

export function movePivotLayoutField(layout: PivotLayout, field: PivotFieldDefinition, area: PivotLayoutArea, index: number): PivotLayout {
  const next = structuredClone(layout);
  if (area === 'values') {
    const base = `value:${field.fieldId}`;
    let valueId = base;
    for (let suffix = 2; next.values.some((value) => value.valueId === valueId); suffix += 1) valueId = `${base}:${suffix}`;
    next.values.splice(boundedInsertIndex(index, next.values.length), 0, { valueId, fieldId: field.fieldId, summarizeBy: field.dataType === 'number' ? 'sum' : 'count' });
    return next;
  }
  const existingRow = next.rows.find((placement) => placement.fieldId === field.fieldId);
  const existingColumn = next.columns.find((placement) => placement.fieldId === field.fieldId);
  const placement = structuredClone(existingRow ?? existingColumn ?? { fieldId: field.fieldId });
  const previousIndex = area === 'rows'
    ? next.rows.findIndex((candidate) => candidate.fieldId === field.fieldId)
    : area === 'columns' ? next.columns.findIndex((candidate) => candidate.fieldId === field.fieldId) : -1;
  next.rows = next.rows.filter((candidate) => candidate.fieldId !== field.fieldId);
  next.columns = next.columns.filter((candidate) => candidate.fieldId !== field.fieldId);
  if (area === 'filters') {
    next.filters = rescopePivotFieldFilters(next, field.fieldId, 'field', 'report');
    if (!next.filters.some((filter) => filter.fieldId === field.fieldId && (filter.scope ?? 'report') === 'report')) {
      next.filters.push({ kind: 'manual', family: 'manual', fieldId: field.fieldId, scope: 'report', mode: 'all', memberKeys: [] });
    }
    next.filters = reorderPivotReportFilterField(next.filters, field.fieldId, index);
    return next;
  }
  if (!existingRow && !existingColumn) next.filters = rescopePivotFieldFilters(next, field.fieldId, 'report', 'field');
  const target = area === 'rows' ? next.rows : next.columns;
  const adjustedIndex = previousIndex >= 0 && previousIndex < index ? index - 1 : index;
  const targetIndex = boundedInsertIndex(adjustedIndex, target.length);
  target.splice(targetIndex, 0, placement);
  return next;
}

export function movePivotValuePlacement(layout: PivotLayout, valueId: string, index: number): PivotLayout {
  const next = structuredClone(layout);
  const currentIndex = next.values.findIndex((value) => value.valueId === valueId);
  if (currentIndex < 0) throw new Error(`Unknown Pivot Values placement: ${valueId}`);
  const [value] = next.values.splice(currentIndex, 1);
  const adjustedIndex = currentIndex < index ? index - 1 : index;
  const targetIndex = boundedInsertIndex(adjustedIndex, next.values.length);
  next.values.splice(targetIndex, 0, value!);
  return next;
}

export function removePivotLayoutPlacement(layout: PivotLayout, area: PivotLayoutArea, placementId: string): PivotLayout {
  const next = structuredClone(layout);
  if (area === 'values') {
    if (!next.values.some((value) => value.valueId === placementId)) throw new Error(`Unknown Pivot Values placement: ${placementId}`);
    next.values = next.values.filter((value) => value.valueId !== placementId);
    next.filters = next.filters.filter((filter) => !(filter.kind === 'top-items' || (filter.kind === 'condition' && filter.family === 'value')) || filter.valueId !== placementId);
  } else if (area === 'filters') {
    next.filters = next.filters.filter((filter) => filter.fieldId !== placementId || (filter.scope ?? 'report') === 'field');
  } else {
    next[area] = next[area].filter((placement) => placement.fieldId !== placementId);
    const remainsOnAxis = next.rows.some((placement) => placement.fieldId === placementId)
      || next.columns.some((placement) => placement.fieldId === placementId);
    if (!remainsOnAxis) next.filters = next.filters.filter((filter) => filter.fieldId !== placementId || (filter.scope ?? 'report') !== 'field');
  }
  return next;
}

export function replacePivotValuePlacement(layout: PivotLayout, value: PivotValueField): PivotLayout {
  if (!layout.values.some((candidate) => candidate.valueId === value.valueId)) throw new Error(`Unknown Pivot Values placement: ${value.valueId}`);
  return { ...structuredClone(layout), values: layout.values.map((candidate) => candidate.valueId === value.valueId ? structuredClone(value) : candidate) };
}

function hasPivotHeaderData(workbook: WorkbookModel, pivot: PivotModel): boolean {
  return getPivotSourceRanges(workbook, pivot).some((range) => {
    const sheet = workbook.getSheet(range.sheetId);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (sheet.cells.get(range.startRow, column)?.value != null) return true;
    }
    return false;
  });
}

export function buildPivotPanelState(workbook: WorkbookModel, pivot: PivotModel): PivotPanelState {
  assertPivotDefinition(workbook, pivot);
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  return {
    pivotId: definition.id,
    sheetId: definition.target.sheetId,
    fieldCatalog: definition.fieldCatalog,
    layout: structuredClone(definition.layout),
    resultTreeSchema: PIVOT_RESULT_TREE_SCHEMA,
  };
}

export function listAvailablePivotFields(workbook: WorkbookModel, pivot: PivotModel): string[] {
  assertPivotDefinition(workbook, pivot);
  return pivot.fieldCatalog.fields.map((field) => field.fieldId);
}

/** Validate a field reference against the live source before a command mutates the model. */
export function assertPivotField(workbook: WorkbookModel, pivot: PivotModel, fieldId: string): void {
  if (!hasPivotHeaderData(workbook, pivot)) return;
  const catalog = pivot.fieldCatalog;
  const names = new Set(catalog.fields.flatMap((entry) => [entry.fieldId, entry.name]));
  for (const calculated of pivot.layout.calculatedFields ?? []) names.add(calculated.fieldId);
  if (!names.has(fieldId)) throw new Error(`Unknown pivot field: ${fieldId}`);
}

/** Fail closed for malformed definitions instead of producing an empty pivot silently. */
export function assertPivotDefinition(workbook: WorkbookModel, pivot: PivotModel): void {
  if (!pivot.id.trim()) throw new Error('Pivot id is required');
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  workbook.getSheet(definition.target.sheetId);
  if (!Number.isSafeInteger(definition.target.anchor.row) || definition.target.anchor.row < 0 || !Number.isSafeInteger(definition.target.anchor.column) || definition.target.anchor.column < 0) throw new Error('Pivot target anchor is invalid');
  const ranges = getPivotSourceRanges(workbook, pivot);
  if (!ranges.length) throw new Error('Pivot source range is required');
  for (const range of ranges) {
    const sheet = workbook.getSheet(range.sheetId);
    if (range.startRow < 0 || range.endRow < range.startRow || range.endRow >= sheet.rowCount || range.startColumn < 0 || range.endColumn < range.startColumn || range.endColumn >= sheet.columnCount) throw new Error('Pivot source range is invalid');
  }
  if (!hasPivotHeaderData(workbook, pivot)) return;
  const fields = new Set(definition.fieldCatalog.fields.flatMap((entry) => [entry.fieldId, entry.name]));
  for (const calculated of definition.layout.calculatedFields ?? []) fields.add(calculated.fieldId);
  for (const calculated of definition.layout.calculatedItems ?? []) fields.add(calculated.fieldId);
  const valueSourceReferences = definition.layout.filters.flatMap((filter) => {
    const valueId = filter.kind === 'top-items' || (filter.kind === 'condition' && filter.valueId) ? filter.valueId : undefined;
    if (!valueId) return [];
    const value = definition.layout.values.find((entry) => entry.valueId === valueId);
    if (!value) throw new Error(`Unknown Pivot Values placement: ${valueId}`);
    return [value.fieldId];
  });
  const references = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.map((filter) => filter.fieldId),
    ...valueSourceReferences,
    ...definition.layout.values.map((entry) => entry.fieldId),
  ];
  const unknown = references.find((field) => field && !fields.has(field));
  if (unknown) throw new Error(`Unknown pivot field: ${unknown}`);
}

export function patchPivotValueField(layout: PivotLayout, valueId: string, patch: Partial<PivotValueField>): PivotLayout {
  if (!layout.values.some((entry) => entry.valueId === valueId)) throw new Error(`Unknown Pivot Values placement: ${valueId}`);
  return { ...layout, values: layout.values.map((entry) => (entry.valueId === valueId ? { ...entry, ...patch, valueId } : entry)) };
}

export function patchPivotRowField(layout: PivotLayout, fieldId: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.rows.some((entry) => entry.fieldId === fieldId)) throw new Error(`Unknown pivot row field: ${fieldId}`);
  return { ...layout, rows: layout.rows.map((entry) => (entry.fieldId === fieldId ? { ...entry, ...patch, fieldId } : entry)) };
}

export function patchPivotColumnField(layout: PivotLayout, fieldId: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.columns.some((entry) => entry.fieldId === fieldId)) throw new Error(`Unknown pivot column field: ${fieldId}`);
  return { ...layout, columns: layout.columns.map((entry) => (entry.fieldId === fieldId ? { ...entry, ...patch, fieldId } : entry)) };
}

export function setPivotAggregate(layout: PivotLayout, valueId: string, summarizeBy: PivotAggregateFunction): PivotLayout {
  return patchPivotValueField(layout, valueId, { summarizeBy });
}

export function setPivotShowAs(layout: PivotLayout, valueId: string, showAs: PivotShowAs): PivotLayout {
  return patchPivotValueField(layout, valueId, { showAs });
}

/** Update the Excel row-grand-total column without changing column totals. */
export function setPivotRowGrandTotals(layout: PivotLayout, enabled: boolean): PivotLayout {
  if (typeof enabled !== 'boolean') throw new Error('Pivot row grand-total state is invalid');
  return { ...layout, showRowGrandTotals: enabled };
}

/** Update the Excel column-grand-total row without changing row totals. */
export function setPivotColumnGrandTotals(layout: PivotLayout, enabled: boolean): PivotLayout {
  if (typeof enabled !== 'boolean') throw new Error('Pivot column grand-total state is invalid');
  return { ...layout, showColumnGrandTotals: enabled };
}

export function setPivotGroup(layout: PivotLayout, axis: 'rows' | 'columns', fieldId: string, group: PivotGroup): PivotLayout {
  return axis === 'rows' ? patchPivotRowField(layout, fieldId, { group }) : patchPivotColumnField(layout, fieldId, { group });
}

export function upsertPivotFilter(layout: PivotLayout, filter: PivotFilter): PivotLayout {
  const fieldId = filter.fieldId;
  const identity = pivotFilterIdentity(filter);
  const filters = layout.filters.filter((entry) => {
    if (pivotFilterIdentity(entry) === identity) return false;
    if (!allowsMultiplePivotFilters(layout)
      && entry.fieldId === fieldId
      && (entry.scope ?? 'report') === (filter.scope ?? 'report')) return false;
    return true;
  });
  filters.push(structuredClone({ ...filter, fieldId }));
  return { ...layout, filters };
}

export function clearPivotFilterFamily(layout: PivotLayout, fieldId: string, family: PivotFilter['family'], scope: 'report' | 'field' = 'report'): PivotLayout {
  return { ...layout, filters: layout.filters.filter((entry) => !(entry.fieldId === fieldId && entry.family === family && (entry.scope ?? 'report') === scope)) };
}

export function clearPivotFiltersForField(layout: PivotLayout, fieldId: string, scope?: 'report' | 'field'): PivotLayout {
  return { ...layout, filters: layout.filters.filter((entry) => entry.fieldId !== fieldId || (scope !== undefined && (entry.scope ?? 'report') !== scope)) };
}

export interface PivotDrillDownTarget {
  sheetId: string;
  pivotId: string;
  targetSheetId: string;
  target: { row: number; column: number };
  sourceRowPaths: PivotSourceRowPath[];
}

export function createPivotDrillDownSheetName(pivot: PivotModel, label: string): string {
  return `Drill ${pivot.id} ${label}`.slice(0, 31);
}
