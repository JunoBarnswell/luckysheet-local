import type { Column, RangeRef, Row, SheetId } from './index';

/**
 * Pivot is a derived view over workbook data. The definition below is the
 * persistence contract; result trees/projections are intentionally separate
 * runtime values and must never be copied into WorksheetModel.cells.
 */
export const PIVOT_DEFINITION_SCHEMA = 'PivotDefinition' as const;
export const PIVOT_RESULT_TREE_SCHEMA = 'PivotResultTree' as const;
export const PIVOT_GRID_PROJECTION_SCHEMA = 'PivotGridProjection' as const;

export type PivotScalar = string | number | boolean | null;
export type PivotScalarType = 'text' | 'number' | 'boolean' | 'blank';

/** A member key keeps `1`, `"1"`, `true`, and blank members distinct. */
export interface PivotMemberKey {
  type: PivotScalarType;
  value: string | number | boolean | null;
}

export function createPivotMemberKey(value: PivotScalar): PivotMemberKey {
  if (value === null || value === '') return { type: 'blank', value: null };
  if (typeof value === 'number') return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'text', value };
}

export function pivotMemberKey(value: PivotMemberKey): string {
  return `${value.type}:${JSON.stringify(value.value)}`;
}

export function pivotMemberKeyEquals(left: PivotMemberKey, right: PivotMemberKey): boolean {
  return left.type === right.type && left.value === right.value;
}

export function pivotScalarFromMemberKey(value: PivotMemberKey): PivotScalar {
  return value.type === 'blank' ? null : value.value as Exclude<PivotScalar, null>;
}

export type PivotAggregateFunction =
  | 'sum'
  | 'count'
  | 'count-numbers'
  | 'average'
  | 'min'
  | 'max'
  | 'product'
  | 'stdev'
  | 'stdevp'
  | 'var'
  | 'varp'
  | 'distinct-count';

export interface PivotSourceRelationship {
  id: string;
  left: { sheetId: SheetId; fieldId: string };
  right: { sheetId: SheetId; fieldId: string };
  join: 'inner' | 'left';
}

/** Canonical Pivot source. */
export type PivotWorksheetDataSource =
  | { kind: 'worksheet-range'; range: RangeRef }
  | { kind: 'worksheet-ranges'; ranges: RangeRef[]; relationships: PivotSourceRelationship[] };

export type PivotSource = PivotWorksheetDataSource
  | { kind: 'table'; tableId: string }
  | { kind: 'named-range'; name: string }
  | { kind: 'data-source'; dataSourceId: string };

export type PivotFieldDataType = 'text' | 'number' | 'date' | 'boolean' | 'mixed';

export interface PivotFieldDefinition {
  /** Stable identity. It is derived from the source column, never from a row value. */
  fieldId: string;
  name: string;
  dataType: PivotFieldDataType;
  ordinal: number;
  values?: PivotScalar[];
}

export interface PivotFieldCatalog {
  schema?: 'PivotFieldCatalog';
  fields: PivotFieldDefinition[];
}

export interface PivotManualGroup {
  groupId: string;
  name: string;
  items: PivotMemberKey[];
}

export type PivotGroup =
  | { kind: 'date'; unit: 'year' | 'quarter' | 'month' | 'week' | 'day'; startOfWeek?: 0 | 1 | 2 | 3 | 4 | 5 | 6 }
  | { kind: 'number'; interval: number; start?: number; end?: number }
  | { kind: 'manual'; groups: PivotManualGroup[] };

export type PivotSort = {
  direction: 'ascending' | 'descending';
  by?: 'label' | 'value';
  valueFieldId?: string;
};

export interface PivotFieldPlacement {
  fieldId: string;
  sort?: PivotSort;
  group?: PivotGroup;
}

export type PivotManualFilter = {
  kind: 'manual';
  fieldId: string;
  mode: 'all' | 'include' | 'exclude';
  memberKeys: PivotMemberKey[];
};

export type PivotFilter =
  | PivotManualFilter
  | {
    kind: 'condition';
    fieldId: string;
    operator: 'equals' | 'not-equals' | 'contains' | 'greater-than' | 'greater-or-equal' | 'less-than' | 'less-or-equal';
    value: PivotScalar;
  }
  | {
    kind: 'top-items';
    fieldId: string;
    count: number;
    valueFieldId: string;
    direction: 'top' | 'bottom';
  };

export type PivotShowAs =
  | { kind: 'normal' }
  | { kind: 'grand-percentage' }
  | { kind: 'row-percentage' }
  | { kind: 'column-percentage' }
  | { kind: 'parent-percentage' }
  | { kind: 'difference'; base: 'grand' | 'row' | 'column' | 'parent' }
  | { kind: 'percentage-difference'; base: 'grand' | 'row' | 'column' | 'parent' }
  | { kind: 'running-total'; axis: 'row' | 'column' }
  | { kind: 'rank'; axis: 'row' | 'column'; direction: 'ascending' | 'descending' }
  | { kind: 'index' };

export interface PivotValueField {
  fieldId: string;
  summarizeBy: PivotAggregateFunction;
  displayName?: string;
  numberFormat?: string;
  baseFieldId?: string;
  baseItem?: PivotMemberKey | string;
  showAs?: PivotShowAs;
}

export interface PivotCalculatedField {
  fieldId: string;
  name: string;
  formula: string;
}

export interface PivotCalculatedItem {
  fieldId: string;
  targetFieldId: string;
  name: string;
  formula: string;
}

export interface PivotExpansionState {
  /** Stable result-node paths. This is intentionally not a field-level flag. */
  expandedNodeIds: string[];
  collapsedNodeIds: string[];
  showButtons: boolean;
}

export interface PivotLayout {
  rows: PivotFieldPlacement[];
  columns: PivotFieldPlacement[];
  filters: PivotFilter[];
  values: PivotValueField[];
  calculatedFields?: PivotCalculatedField[];
  calculatedItems?: PivotCalculatedItem[];
  showSubtotals: boolean;
  showGrandTotals: boolean;
  compact: boolean;
  repeatLabels: boolean;
  expansion?: PivotExpansionState;
}

export interface PivotRefreshPolicy {
  mode: 'manual' | 'on-open' | 'on-change';
  preserveFormatting: boolean;
  refreshOnLoad: boolean;
}

export interface PivotNativeMetadata {
  cacheId?: number;
  cacheDefinitionPart?: string;
  cacheRecordsPart?: string;
  pivotTablePart?: string;
  fieldBindings?: Record<string, { cacheFieldIndex: number; sourceName?: string }>;
  /** Only identifiers and style/display attributes may cross the model boundary. */
  preservedFeatures?: Array<'external-connection' | 'olap' | 'consolidation' | 'macro' | 'custom-xml' | 'slicer' | 'timeline'>;
}

export interface PivotTarget {
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
}

/** Canonical persisted Pivot definition. */
export interface PivotDefinition {
  schema: typeof PIVOT_DEFINITION_SCHEMA;
  id: string;
  source: PivotSource;
  target: PivotTarget;
  fieldCatalog: PivotFieldCatalog;
  layout: PivotLayout;
  refreshPolicy: PivotRefreshPolicy;
  nativeMetadata?: PivotNativeMetadata;
}

/** Public alias used by workbook collections; there is exactly one Pivot shape. */
export type PivotModel = PivotDefinition;

export interface PivotSourceRowPath {
  sheetId: SheetId;
  row: Row;
}

export interface PivotResultCell {
  id?: string;
  nodePath?: string[];
  kind?: 'detail' | 'subtotal' | 'grand-total';
  columnPath: PivotScalar[];
  values: PivotScalar[];
  sourceRowPaths: PivotSourceRowPath[];
}

export interface PivotResultNode {
  nodeId?: string;
  path?: string[];
  kind: 'leaf' | 'subtotal';
  fieldId?: string;
  memberKey?: PivotMemberKey;
  key: PivotScalar;
  label: string;
  depth: number;
  children: PivotResultNode[];
  values: PivotResultCell[];
  subtotal: boolean;
  sourceRowPaths: PivotSourceRowPath[];
}

export interface PivotResultTree {
  schema: typeof PIVOT_RESULT_TREE_SCHEMA;
  pivotId: string;
  fields: PivotFieldCatalog;
  columnPaths: PivotScalar[][];
  rows: PivotResultNode[];
  grandTotal: PivotResultCell | null;
  sourceRowPaths: PivotSourceRowPath[];
  sourceRevision?: string;
  layoutRevision?: string;
  filterRevision?: string;
}

export type PivotProjectionCellKind =
  | 'title'
  | 'filter'
  | 'row-header'
  | 'column-header'
  | 'value'
  | 'subtotal'
  | 'grand-total'
  | 'expand-toggle'
  | 'loading'
  | 'error';

export interface PivotProjectionCell {
  id: string;
  pivotId: string;
  row: number;
  column: number;
  kind: PivotProjectionCellKind;
  value: PivotScalar;
  text: string;
  nodeId?: string;
  resultCellId?: string;
  columnPath?: PivotScalar[];
  sourceRowPaths?: PivotSourceRowPath[];
  expandable?: boolean;
  expanded?: boolean;
}

export interface PivotCollision {
  status: 'clear' | 'collision';
  reasons: Array<'cell-data' | 'merge' | 'pivot' | 'worksheet-bounds'>;
  conflictingRanges: RangeRef[];
}

export type PivotRefreshStatus = 'idle' | 'refreshing' | 'ready' | 'stale' | 'error' | 'collision';

export interface PivotRefreshState {
  status: PivotRefreshStatus;
  revision: number;
  sourceRevision: string;
  requestedAt?: string;
  completedAt?: string;
  error?: string;
}

export interface PivotGridProjection {
  schema: typeof PIVOT_GRID_PROJECTION_SCHEMA;
  pivotId: string;
  sheetId: SheetId;
  target: PivotTarget;
  occupiedRange: RangeRef;
  cells: PivotProjectionCell[];
  collision: PivotCollision;
  refresh: PivotRefreshState;
}

export type PivotHitKind = 'cell' | 'expand-toggle' | 'filter' | 'header' | 'none';

export interface PivotHitTest {
  kind: PivotHitKind;
  pivotId: string;
  cellId?: string;
  row?: number;
  column?: number;
  nodeId?: string;
  sourceRowPaths?: PivotSourceRowPath[];
}

/** Generic context-hit shape used by the UI resolver for Pivot cells. */
export interface ContextHit extends PivotHitTest {
  context: 'pivot';
  priority: 30;
}

/**
 * Pure snapshot migration. It only translates shape/identity fields; live
 * source inspection and field catalog inference belong to the spreadsheet-app
 * engine where a WorkbookModel is available.
 */
export function canonicalizePivotDefinition(input: PivotDefinition): PivotDefinition {
  if (input.schema !== PIVOT_DEFINITION_SCHEMA) throw new Error(`Pivot ${input.id} is not a canonical definition`);
  if (!input.source || !input.target || !input.fieldCatalog || !input.refreshPolicy) throw new Error(`Pivot ${input.id} is missing canonical fields`);
  if (input.layout.rows.some((entry) => !entry.fieldId)
    || input.layout.columns.some((entry) => !entry.fieldId)
    || input.layout.values.some((entry) => !entry.fieldId)
    || input.layout.filters.some((entry) => !entry.fieldId)) {
    throw new Error(`Pivot ${input.id} has non-canonical field references`);
  }
  return structuredClone(input);
}
