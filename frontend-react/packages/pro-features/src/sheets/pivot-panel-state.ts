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
  WorkbookModel,
} from '@react-sheets/core-model';
import { computePivotResult, getPivotFieldCatalog } from './pivot-engine';

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

export function buildPivotPanelState(workbook: WorkbookModel, pivot: PivotModel): PivotPanelState {
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

export function patchPivotValueField(layout: PivotLayout, field: string, patch: Partial<PivotValueField>): PivotLayout {
  return {
    ...layout,
    values: layout.values.map((entry) => (entry.field === field ? { ...entry, ...patch } : entry)),
  };
}

export function patchPivotRowField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
  return {
    ...layout,
    rows: layout.rows.map((entry) => (entry.field === field ? { ...entry, ...patch } : entry)),
  };
}

export function patchPivotColumnField(layout: PivotLayout, field: string, patch: Partial<PivotFieldPlacement>): PivotLayout {
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
