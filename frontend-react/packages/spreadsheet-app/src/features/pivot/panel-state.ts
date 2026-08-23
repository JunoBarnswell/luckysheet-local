import type {
  PivotAggregateFunction,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotShowAs,
  PivotValueField,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { computePivotResult, getPivotFieldCatalog, getPivotSourceRanges, normalizePivotDefinition } from './engine';

export interface PivotPanelState {
  pivotId: string;
  sheetId: string;
  fieldCatalog: ReturnType<typeof getPivotFieldCatalog>;
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
  const definition = normalizePivotDefinition(workbook, pivot);
  const result = computePivotResult(workbook, pivot);
  return {
    pivotId: definition.id,
    sheetId: definition.target.sheetId,
    fieldCatalog: definition.fieldCatalog,
    layout: structuredClone(definition.layout),
    resultTreeSchema: result.schema,
  };
}

export function listAvailablePivotFields(workbook: WorkbookModel, pivot: PivotModel): string[] {
  return getPivotFieldCatalog(workbook, pivot).fields.map((field) => field.fieldId);
}

/** Validate a field reference against the live source before a command mutates the model. */
export function assertPivotField(workbook: WorkbookModel, pivot: PivotModel, field: string): void {
  if (!hasPivotHeaderData(workbook, pivot)) return;
  const catalog = getPivotFieldCatalog(workbook, pivot);
  const names = new Set(catalog.fields.flatMap((entry) => [entry.fieldId, entry.name]));
  for (const calculated of pivot.layout.calculatedFields ?? []) names.add(calculated.fieldId);
  if (!names.has(field)) throw new Error(`Unknown pivot field: ${field}`);
}

/** Fail closed for malformed definitions instead of producing an empty pivot silently. */
export function assertPivotDefinition(workbook: WorkbookModel, pivot: PivotModel): void {
  if (!pivot.id.trim()) throw new Error('Pivot id is required');
  const definition = normalizePivotDefinition(workbook, pivot);
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
  const references = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.flatMap((filter) => filter.kind === 'top-items' ? [filter.fieldId, filter.valueFieldId] : [filter.fieldId]),
    ...definition.layout.values.map((entry) => entry.fieldId),
  ];
  const unknown = references.find((field) => field && !fields.has(field));
  if (unknown) throw new Error(`Unknown pivot field: ${unknown}`);
}

export function patchPivotValueField(layout: PivotLayout, field: string, patch: Partial<PivotValueField>): PivotLayout {
  if (!layout.values.some((entry) => entry.fieldId === field)) throw new Error(`Unknown pivot value field: ${field}`);
  return { ...layout, values: layout.values.map((entry) => (entry.fieldId === field ? { ...entry, ...patch, fieldId: field } : entry)) };
}

export function patchPivotRowField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.rows.some((entry) => entry.fieldId === field)) throw new Error(`Unknown pivot row field: ${field}`);
  return { ...layout, rows: layout.rows.map((entry) => (entry.fieldId === field ? { ...entry, ...patch, fieldId: field } : entry)) };
}

export function patchPivotColumnField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.columns.some((entry) => entry.fieldId === field)) throw new Error(`Unknown pivot column field: ${field}`);
  return { ...layout, columns: layout.columns.map((entry) => (entry.fieldId === field ? { ...entry, ...patch, fieldId: field } : entry)) };
}

export function setPivotAggregate(layout: PivotLayout, field: string, summarizeBy: PivotAggregateFunction): PivotLayout {
  return patchPivotValueField(layout, field, { summarizeBy });
}

export function setPivotShowAs(layout: PivotLayout, field: string, showAs: PivotShowAs): PivotLayout {
  return patchPivotValueField(layout, field, { showAs });
}

export function setPivotGroup(layout: PivotLayout, axis: 'rows' | 'columns', field: string, group: PivotGroup): PivotLayout {
  return axis === 'rows' ? patchPivotRowField(layout, field, { group }) : patchPivotColumnField(layout, field, { group });
}

export function upsertPivotFilter(layout: PivotLayout, filter: PivotFilter): PivotLayout {
  const fieldId = filter.fieldId;
  const filters = layout.filters.filter((entry) => entry.fieldId !== fieldId);
  filters.push(structuredClone({ ...filter, fieldId }));
  return { ...layout, filters };
}

export interface PivotDrillDownTarget {
  sheetId: string;
  pivotId: string;
  targetSheetId: string;
  target: { row: number; column: number };
  sourceRowPaths: Array<{ sheetId: string; row: number }>;
}

export function createPivotDrillDownSheetName(pivot: PivotModel, label: string): string {
  return `Drill ${pivot.id} ${label}`.slice(0, 31);
}
