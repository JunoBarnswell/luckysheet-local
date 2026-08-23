import type { Column, RangeRef, Row, SheetId } from './index';
export type PivotScalar = string | number | boolean | null;
export type PivotAggregateFunction = 'sum' | 'count' | 'count-numbers' | 'average' | 'min' | 'max' | 'product' | 'stdev' | 'stdevp' | 'var' | 'varp' | 'distinct-count';
export interface PivotSourceRelationship { id: string; left: { sheetId: SheetId; field: string }; right: { sheetId: SheetId; field: string }; join: 'inner' | 'left'; }
export type PivotDataSource = { kind: 'worksheet-range'; range: RangeRef } | { kind: 'worksheet-ranges'; ranges: RangeRef[]; relationships: PivotSourceRelationship[] };
export type PivotFieldDataType = 'text' | 'number' | 'date' | 'boolean' | 'mixed';
export interface PivotFieldDefinition { id: string; name: string; dataType: PivotFieldDataType; ordinal: number; values?: PivotScalar[]; }
export interface PivotFieldCatalog { fields: PivotFieldDefinition[]; }
export interface PivotManualGroup { name: string; items: PivotScalar[]; }
export type PivotGroup = { kind: 'date'; unit: 'year' | 'quarter' | 'month' | 'week' | 'day'; startOfWeek?: 0 | 1 | 2 | 3 | 4 | 5 | 6 } | { kind: 'number'; interval: number; start?: number; end?: number } | { kind: 'manual'; groups: PivotManualGroup[] };
export type PivotSort = { direction: 'ascending' | 'descending'; by?: 'label' | 'value'; valueField?: string };
export interface PivotFieldPlacement { field: string; sort?: PivotSort; group?: PivotGroup; }
export type PivotFilter = { kind: 'manual'; field: string; selected: PivotScalar[]; exclude?: boolean } | { kind: 'condition'; field: string; operator: 'equals' | 'not-equals' | 'contains' | 'greater-than' | 'greater-or-equal' | 'less-than' | 'less-or-equal'; value: PivotScalar } | { kind: 'top-items'; field: string; count: number; valueField: string; direction: 'top' | 'bottom' };
export type PivotShowAs = { kind: 'normal' } | { kind: 'grand-percentage' } | { kind: 'row-percentage' } | { kind: 'column-percentage' } | { kind: 'parent-percentage' } | { kind: 'difference'; base: 'grand' | 'row' | 'column' | 'parent' } | { kind: 'percentage-difference'; base: 'grand' | 'row' | 'column' | 'parent' } | { kind: 'running-total'; axis: 'row' | 'column' } | { kind: 'rank'; axis: 'row' | 'column'; direction: 'ascending' | 'descending' } | { kind: 'index' };
export interface PivotValueField { field: string; summarizeBy: PivotAggregateFunction; displayName?: string; showAs?: PivotShowAs; }
export interface PivotLayout { rows: PivotFieldPlacement[]; columns: PivotFieldPlacement[]; filters: PivotFilter[]; values: PivotValueField[]; showSubtotals: boolean; showGrandTotals: boolean; compact: boolean; repeatLabels: boolean; expandedFieldIds?: string[]; }
export interface PivotRefreshPolicy { mode: 'manual' | 'on-open' | 'on-change'; preserveFormatting: boolean; refreshOnLoad: boolean; }
export interface PivotSlicer { id: string; field: string; selected: PivotScalar[]; }
export interface PivotTimeline { id: string; field: string; start?: string; end?: string; }
export interface PivotChartReference { chartId: string; role: 'source' | 'linked'; }
export interface PivotSourceRowPath { sheetId: SheetId; row: Row; }
export interface PivotResultCell { kind?: 'detail' | 'grand-total'; columnPath: PivotScalar[]; values: PivotScalar[]; sourceRowPaths: PivotSourceRowPath[]; }
export interface PivotResultNode { kind: 'leaf' | 'subtotal'; field?: string; key: PivotScalar; label: string; depth: number; children: PivotResultNode[]; values: PivotResultCell[]; subtotal: boolean; sourceRowPaths: PivotSourceRowPath[]; }
export interface PivotResultTree { schema: 'PivotResultTreeV1'; pivotId: string; fields: PivotFieldCatalog; columnPaths: PivotScalar[][]; rows: PivotResultNode[]; grandTotal: PivotResultCell | null; sourceRowPaths: PivotSourceRowPath[]; }
export interface PivotModel { id: string; sheetId: SheetId; sourceRange: RangeRef; dataSource?: PivotDataSource; fieldCatalog?: PivotFieldCatalog; layout?: PivotLayout; refreshPolicy?: PivotRefreshPolicy; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[]; targetAnchor?: { row: Row; column: Column }; rowFields: string[]; columnFields: string[]; valueFields: PivotValueField[]; filterFields: string[]; }
export type NormalizedPivotModel = Omit<PivotModel, 'rowFields' | 'columnFields' | 'valueFields' | 'filterFields' | 'layout'> & { layout: PivotLayout };
export function normalizePivotDefinition(pivot: PivotModel): NormalizedPivotModel {
  const layout = pivot.layout ?? {
    rows: pivot.rowFields.map((field) => ({ field })), columns: pivot.columnFields.map((field) => ({ field })), filters: [], values: pivot.valueFields,
    showSubtotals: true, showGrandTotals: true, compact: false, repeatLabels: false,
  };
  const { rowFields: _rowFields, columnFields: _columnFields, valueFields: _valueFields, filterFields: _filterFields, ...canonical } = pivot;
  return { ...canonical, layout: structuredClone(layout) };
}
