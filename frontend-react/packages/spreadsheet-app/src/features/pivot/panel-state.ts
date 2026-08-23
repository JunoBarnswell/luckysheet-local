import type {
  PivotAggregateFunction,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotShowAs,
  PivotSlicer,
  PivotTimeline,
  PivotValueField,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { computePivotResult, getPivotFieldCatalog } from './engine';

export interface PivotPanelState {
  pivotId: string;
  sheetId: string;
  fieldCatalog: ReturnType<typeof getPivotFieldCatalog>;
  layout: PivotLayout;
  slicers: PivotSlicer[];
  timelines: PivotTimeline[];
  resultTreeSchema: string;
  refreshRevision: number;
}

function pivotSourceRanges(pivot: PivotModel): RangeRef[] {
  return pivot.dataSource?.kind === 'worksheet-ranges'
    ? pivot.dataSource.ranges
    : [pivot.dataSource?.range ?? pivot.sourceRange];
}

function hasPivotHeaderData(workbook: WorkbookModel, pivot: PivotModel): boolean {
  return pivotSourceRanges(pivot).some((range) => {
    const sheet = workbook.getSheet(range.sheetId);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (sheet.cells.get(range.startRow, column)?.value != null) return true;
    }
    return false;
  });
}

export function buildPivotPanelState(workbook: WorkbookModel, pivot: PivotModel): PivotPanelState {
  assertPivotDefinition(workbook, pivot);
  const result = computePivotResult(workbook, pivot);
  return {
    pivotId: pivot.id,
    sheetId: pivot.sheetId,
    fieldCatalog: getPivotFieldCatalog(workbook, pivot),
    layout: structuredClone(pivot.layout),
    slicers: structuredClone(pivot.slicers ?? []),
    timelines: structuredClone(pivot.timelines ?? []),
    resultTreeSchema: result.schema,
    refreshRevision: pivot.refreshRevision ?? 0,
  };
}

export function listAvailablePivotFields(workbook: WorkbookModel, pivot: PivotModel): string[] {
  return getPivotFieldCatalog(workbook, pivot).fields.map((field) => field.name);
}

/** Validate a field reference against the live source before a command mutates the model. */
export function assertPivotField(workbook: WorkbookModel, pivot: PivotModel, field: string): void {
  if (!hasPivotHeaderData(workbook, pivot)) return;
  const names = new Set(getPivotFieldCatalog(workbook, pivot).fields.map((entry) => entry.name));
  for (const calculated of pivot.layout.calculatedFields ?? []) names.add(calculated.name);
  if (!names.has(field)) throw new Error(`Unknown pivot field: ${field}`);
}

/** Fail closed for malformed definitions instead of producing an empty pivot silently. */
export function assertPivotDefinition(workbook: WorkbookModel, pivot: PivotModel): void {
  if (!pivot.id.trim()) throw new Error('Pivot id is required');
  if (!pivot.sheetId.trim()) throw new Error('Pivot display sheet is required');
  workbook.getSheet(pivot.sheetId);
  const ranges = pivot.dataSource?.kind === 'worksheet-ranges'
    ? pivot.dataSource.ranges
    : [pivot.dataSource?.range ?? pivot.sourceRange];
  if (ranges.length === 0) throw new Error('Pivot source range is required');
  for (const range of ranges) {
    workbook.getSheet(range.sheetId);
    if (range.startRow < 0 || range.endRow < range.startRow || range.startColumn < 0 || range.endColumn < range.startColumn) {
      throw new Error('Pivot source range is invalid');
    }
  }
  const hasHeaderData = hasPivotHeaderData(workbook, pivot);
  // An empty source is a valid, not-yet-populated pivot definition.  There is
  // no field contract to validate until a header exists; non-empty sources
  // remain fail-closed below.
  if (!hasHeaderData) return;
  const fields = new Set(getPivotFieldCatalog(workbook, pivot).fields.map((entry) => entry.name));
  for (const calculated of pivot.layout.calculatedFields ?? []) fields.add(calculated.name);
  for (const calculated of pivot.layout.calculatedItems ?? []) fields.add(calculated.name);
  const references = [
    ...pivot.layout.rows.map((entry) => entry.field),
    ...pivot.layout.columns.map((entry) => entry.field),
    ...pivot.layout.filters.map((entry) => entry.field),
    ...pivot.layout.values.map((entry) => entry.field),
    ...(pivot.slicers ?? []).map((entry) => entry.field),
    ...(pivot.timelines ?? []).map((entry) => entry.field),
  ];
  const unknown = references.find((field) => !fields.has(field));
  if (unknown) throw new Error(`Unknown pivot field: ${unknown}`);
}

export function patchPivotValueField(layout: PivotLayout, field: string, patch: Partial<PivotValueField>): PivotLayout {
  if (!layout.values.some((entry) => entry.field === field)) throw new Error(`Unknown pivot value field: ${field}`);
  return {
    ...layout,
    values: layout.values.map((entry) => (entry.field === field ? { ...entry, ...patch } : entry)),
  };
}

export function patchPivotRowField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.rows.some((entry) => entry.field === field)) throw new Error(`Unknown pivot row field: ${field}`);
  return {
    ...layout,
    rows: layout.rows.map((entry) => (entry.field === field ? { ...entry, ...patch } : entry)),
  };
}

export function patchPivotColumnField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  if (!layout.columns.some((entry) => entry.field === field)) throw new Error(`Unknown pivot column field: ${field}`);
  return {
    ...layout,
    columns: layout.columns.map((entry) => (entry.field === field ? { ...entry, ...patch } : entry)),
  };
}

export function setPivotAggregate(layout: PivotLayout, field: string, summarizeBy: PivotAggregateFunction): PivotLayout {
  return patchPivotValueField(layout, field, { summarizeBy });
}

export function setPivotShowAs(layout: PivotLayout, field: string, showAs: PivotShowAs): PivotLayout {
  return patchPivotValueField(layout, field, { showAs });
}

export function setPivotGroup(layout: PivotLayout, axis: 'rows' | 'columns', field: string, group: PivotGroup): PivotLayout {
  if (axis === 'rows') return patchPivotRowField(layout, field, { group });
  return patchPivotColumnField(layout, field, { group });
}

export function upsertPivotFilter(layout: PivotLayout, filter: PivotFilter): PivotLayout {
  const filters = layout.filters.filter((entry) => entry.field !== filter.field);
  filters.push(structuredClone(filter));
  return { ...layout, filters };
}

export function upsertPivotSlicer(pivot: PivotModel, slicer: PivotSlicer): PivotSlicer[] {
  const slicers = [...(pivot.slicers ?? [])];
  const index = slicers.findIndex((entry) => entry.id === slicer.id);
  if (index >= 0) slicers[index] = structuredClone(slicer);
  else slicers.push(structuredClone(slicer));
  return slicers;
}

export function upsertPivotTimeline(pivot: PivotModel, timeline: PivotTimeline): PivotTimeline[] {
  const timelines = [...(pivot.timelines ?? [])];
  const index = timelines.findIndex((entry) => entry.id === timeline.id);
  if (index >= 0) timelines[index] = structuredClone(timeline);
  else timelines.push(structuredClone(timeline));
  return timelines;
}

export interface PivotDrillDownTarget {
  sheetId: string;
  pivotId: string;
  targetSheetId: string;
  targetAnchor: { row: number; column: number };
  sourceRowPaths: Array<{ sheetId: string; row: number }>;
}

export function createPivotDrillDownSheetName(pivot: PivotModel, label: string): string {
  return `Drill ${pivot.id} ${label}`.slice(0, 31);
}
