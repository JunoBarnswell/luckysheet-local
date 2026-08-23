import type {
  PivotAggregateFunction,
  PivotLayout,
  PivotModel,
  PivotSlicer,
  PivotTimeline,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { getPivotFieldCatalog } from '@react-sheets/pro-features';

export function buildDefaultPivotLayout(
  workbook: WorkbookModel,
  sheetId: string,
  sourceRange: RangeRef,
): PivotLayout | undefined {
  const catalog = getPivotFieldCatalog(workbook, {
    id: 'pivot-layout-draft',
    sheetId,
    sourceRange,
    layout: {
      rows: [],
      columns: [],
      filters: [],
      values: [],
      showSubtotals: true,
      showGrandTotals: true,
      compact: true,
      repeatLabels: false,
      calculatedFields: [],
      calculatedItems: [],
    },
  }).fields;
  const rowField = catalog.find((field) => field.dataType !== 'number')?.name ?? catalog[0]?.name;
  const valueField = catalog.find((field) => field.dataType === 'number')?.name ?? catalog[0]?.name;
  if (!rowField || !valueField) return undefined;
  const summarizeBy: PivotAggregateFunction = catalog.find((field) => field.name === valueField)?.dataType === 'number' ? 'sum' : 'count';
  return {
    rows: [{ field: rowField }],
    columns: [],
    filters: [],
    values: [{ field: valueField, summarizeBy }],
    showSubtotals: true,
    showGrandTotals: true,
    compact: true,
    repeatLabels: false,
    calculatedFields: [],
    calculatedItems: [],
  };
}

export function buildPivotModel(
  workbook: WorkbookModel,
  sheetId: string,
  pivotId: string,
  sourceRange: RangeRef,
): PivotModel | undefined {
  const layout = buildDefaultPivotLayout(workbook, sheetId, sourceRange);
  if (!layout) return undefined;
  return {
    id: pivotId,
    sheetId,
    sourceRange: { ...sourceRange, sheetId },
    refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    layout,
  };
}

export function connectedPivotIdsForSource(workbook: WorkbookModel, sheetId: string, sourceRange: RangeRef): string[] {
  const sheet = workbook.getSheet(sheetId);
  return sheet.pivots
    .filter((pivot) =>
      pivot.sourceRange.startRow === sourceRange.startRow
      && pivot.sourceRange.endRow === sourceRange.endRow
      && pivot.sourceRange.startColumn === sourceRange.startColumn
      && pivot.sourceRange.endColumn === sourceRange.endColumn)
    .map((pivot) => pivot.id);
}

export function buildPivotSlicer(
  pivot: PivotModel,
  workbook: WorkbookModel,
  field: string,
  connectedPivotIds: string[],
): PivotSlicer {
  return {
    id: `slicer-${field}`,
    field,
    selected: pivot.slicers?.find((entry) => entry.field === field)?.selected ?? [],
    connectedPivotIds,
  };
}

export function buildPivotTimeline(
  pivot: PivotModel,
  field: string,
  connectedPivotIds: string[],
  start?: string,
  end?: string,
): PivotTimeline {
  const existing = pivot.timelines?.find((entry) => entry.field === field);
  return {
    id: existing?.id ?? `timeline-${field}`,
    field,
    start: start ?? existing?.start,
    end: end ?? existing?.end,
    connectedPivotIds,
  };
}
