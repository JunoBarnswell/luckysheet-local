import type {
  PivotAggregateFunction,
  PivotFieldCatalog,
  PivotFieldPlacement,
  PivotLayout,
  PivotModel,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { getPivotFieldCatalog, getPivotSourceRanges } from './engine';

export function buildDefaultPivotLayout(workbook: WorkbookModel, sheetId: string, sourceRange: RangeRef): PivotLayout | undefined {
  const draft = {
    id: 'pivot-layout-draft',
    schema: 'PivotDefinition' as const,
    source: { kind: 'worksheet-range' as const, range: structuredClone(sourceRange) },
    target: { sheetId, anchor: { row: sourceRange.endRow + 2, column: sourceRange.startColumn } },
    fieldCatalog: { fields: [] },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
    layout: {
      rows: [], columns: [], filters: [], values: [], showSubtotals: true, showGrandTotals: true,
      compact: true, repeatLabels: false, calculatedFields: [], calculatedItems: [],
    },
  } satisfies PivotModel;
  const catalog = getPivotFieldCatalog(workbook, draft);
  const rowField = catalog.fields.find((field) => field.dataType === 'text' || field.dataType === 'boolean' || field.dataType === 'mixed');
  const dateField = catalog.fields.find((field) => field.dataType === 'date');
  const valueField = catalog.fields.find((field) => field.dataType === 'number') ?? catalog.fields[0];
  if (!valueField) return undefined;
  const valueId = valueField.fieldId;
  const summarizeBy: PivotAggregateFunction = valueField.dataType === 'number' ? 'sum' : 'count';
  const rows: PivotFieldPlacement[] = rowField ? [{ fieldId: rowField.fieldId }] : [];
  const columns: PivotFieldPlacement[] = dateField ? [{ fieldId: dateField.fieldId }] : [];
  // A date field is intentionally placed on COLUMNS by default, matching the
  // Excel create-Pivot heuristic; text/boolean fields are placed on ROWS.
  return {
    rows,
    columns,
    filters: [],
    values: [{ fieldId: valueId, summarizeBy }],
    showSubtotals: true,
    showGrandTotals: true,
    compact: true,
    repeatLabels: false,
    calculatedFields: [],
    calculatedItems: [],
    expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
  };
}

export function buildPivotModel(workbook: WorkbookModel, sheetId: string, pivotId: string, sourceRange: RangeRef): PivotModel | undefined {
  const layout = buildDefaultPivotLayout(workbook, sheetId, sourceRange);
  if (!layout) return undefined;
  const source = { kind: 'worksheet-range' as const, range: structuredClone(sourceRange) };
  const draft: PivotModel = {
    schema: 'PivotDefinition',
    id: pivotId,
    source,
    target: { sheetId, anchor: { row: sourceRange.endRow + 2, column: sourceRange.startColumn } },
    fieldCatalog: getPivotFieldCatalog(workbook, { id: pivotId, source, target: { sheetId, anchor: { row: 0, column: 0 } }, layout, refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true } }),
    refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    layout,
  };
  return draft;
}

export function connectedPivotIdsForSource(workbook: WorkbookModel, _sheetId: string, sourceRange: RangeRef): string[] {
  const sameRange = (left: RangeRef, right: RangeRef): boolean => left.sheetId === right.sheetId
    && left.startRow === right.startRow && left.endRow === right.endRow
    && left.startColumn === right.startColumn && left.endColumn === right.endColumn;
  return workbook.getSheets()
    .flatMap((sheet) => sheet.pivots)
    .filter((pivot) => getPivotSourceRanges(workbook, pivot).some((range) => sameRange(range, sourceRange)))
    .map((pivot) => pivot.id);
}
