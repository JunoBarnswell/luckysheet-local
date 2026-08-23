import type {
  PivotAggregateFunction,
  PivotLayout,
  PivotModel,
  PivotSlicer,
  PivotTimeline,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { getPivotFieldCatalog } from './engine';

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
    // The pivot is displayed on `sheetId`, but its source may intentionally
    // be another worksheet.  Keep the source reference authoritative.
    sourceRange: structuredClone(sourceRange),
    refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    layout,
  };
}

export function connectedPivotIdsForSource(workbook: WorkbookModel, sheetId: string, sourceRange: RangeRef): string[] {
  // A source mutation can originate on a worksheet different from the pivot
  // display sheet.  Search the workbook and compare sheet identity as well as
  // coordinates; checking only `sheetId` made linked pivots stale silently.
  workbook.getSheet(sheetId);
  const sameRange = (left: RangeRef, right: RangeRef): boolean =>
    left.sheetId === right.sheetId
    && left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startColumn === right.startColumn
    && left.endColumn === right.endColumn;
  return workbook.getSheets()
    .flatMap((sheet) => sheet.pivots)
    .filter((pivot) => {
      if (sameRange(pivot.sourceRange, sourceRange)) return true;
      if (pivot.dataSource?.kind === 'worksheet-range') return sameRange(pivot.dataSource.range, sourceRange);
      return pivot.dataSource?.ranges.some((range) => sameRange(range, sourceRange)) ?? false;
    })
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
