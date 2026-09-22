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
