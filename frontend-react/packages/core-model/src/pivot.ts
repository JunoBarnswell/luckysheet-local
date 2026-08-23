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
  left: { sheetId: SheetId; fieldId?: string; field?: string };
  right: { sheetId: SheetId; fieldId?: string; field?: string };
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

/** @deprecated Use PivotSource. Kept as a type name for package consumers during migration. */
export type PivotDataSource = PivotWorksheetDataSource;

export type PivotFieldDataType = 'text' | 'number' | 'date' | 'boolean' | 'mixed';

export interface PivotFieldDefinition {
  /** Stable identity. It is derived from the source column, never from a row value. */
  fieldId?: string;
  name: string;
  dataType: PivotFieldDataType;
  ordinal: number;
  values?: PivotScalar[];
  /** @deprecated read-only migration hint; new definitions use fieldId. */
  id?: string;
}

export interface PivotFieldCatalog {
  schema?: 'PivotFieldCatalog';
  fields: PivotFieldDefinition[];
}

export interface PivotManualGroup {
  groupId?: string;
  name: string;
  items: PivotMemberKey[];
  /** @deprecated accepts source values only while migrating old snapshots. */
  legacyItems?: PivotScalar[];
}

export type PivotGroup =
  | { kind: 'date'; unit: 'year' | 'quarter' | 'month' | 'week' | 'day'; startOfWeek?: 0 | 1 | 2 | 3 | 4 | 5 | 6 }
  | { kind: 'number'; interval: number; start?: number; end?: number }
  | { kind: 'manual'; groups: PivotManualGroup[] };

export type PivotSort = {
  direction: 'ascending' | 'descending';
  by?: 'label' | 'value';
  valueFieldId?: string;
  /** @deprecated use valueFieldId. */
  valueField?: string;
};

export interface PivotFieldPlacement {
  fieldId?: string;
  /** @deprecated use fieldId. New layouts emitted by the app never set this. */
  field?: string;
  sort?: PivotSort;
  group?: PivotGroup;
}

export type PivotManualFilter = {
  kind: 'manual';
  fieldId?: string;
  /** @deprecated use fieldId. */
  field?: string;
  mode?: 'all' | 'include' | 'exclude';
  memberKeys?: PivotMemberKey[];
  /** @deprecated old snapshots are normalized at the calculation boundary. */
  selected?: PivotScalar[];
  /** @deprecated old snapshots are normalized at the calculation boundary. */
  exclude?: boolean;
};

export type PivotFilter =
  | PivotManualFilter
  | {
    kind: 'condition';
    fieldId?: string;
    field?: string;
    operator: 'equals' | 'not-equals' | 'contains' | 'greater-than' | 'greater-or-equal' | 'less-than' | 'less-or-equal';
    value: PivotScalar;
  }
  | {
    kind: 'top-items';
    fieldId?: string;
    field?: string;
    count: number;
    valueFieldId?: string;
    valueField?: string;
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
  fieldId?: string;
  /** @deprecated use fieldId. */
  field?: string;
  summarizeBy: PivotAggregateFunction;
  displayName?: string;
  numberFormat?: string;
  baseFieldId?: string;
  baseItem?: PivotMemberKey | string;
  /** @deprecated use baseFieldId. */
  baseField?: string;
  showAs?: PivotShowAs;
}

export interface PivotCalculatedField {
  fieldId?: string;
  name: string;
  formula: string;
}

export interface PivotCalculatedItem {
  fieldId?: string;
  field?: string;
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
  /** @deprecated field-level expansion is migrated to expansion node paths. */
  expandedFieldIds?: string[];
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

/** Floating Pivot controls are owned by their drawing/object records. These
 * narrow records remain useful as an input projection while those records are
 * being migrated out of PivotModel. */
export interface PivotSlicer {
  id: string;
  pivotId?: string;
  fieldId?: string;
  field?: string;
  mode?: 'all' | 'include' | 'exclude';
  memberKeys?: PivotMemberKey[];
  /** @deprecated use mode/memberKeys. */
  selected?: PivotScalar[];
  connectedPivotIds?: string[];
}

export interface PivotTimeline {
  id: string;
  pivotId?: string;
  fieldId?: string;
  field?: string;
  start?: string;
  end?: string;
  connectedPivotIds?: string[];
}

export interface PivotChartReference {
  chartId: string;
  role: 'source' | 'linked';
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

/**
 * WorkbookModel currently exposes this name. New code should construct a
 * PivotDefinition; optional legacy members let a loader inspect old records
 * and hand them to the explicit migration boundary without creating a second
 * calculation path.
 */
export interface PivotModel extends Partial<Omit<PivotDefinition, 'id' | 'layout'>> {
  id: string;
  layout: PivotLayout;
  /** @deprecated old snapshot members; never emitted by canonical builders. */
  sourceRange?: RangeRef;
  /** @deprecated old snapshot members; never emitted by canonical builders. */
  dataSource?: PivotDataSource;
  /** @deprecated use target.sheetId. */
  sheetId?: SheetId;
  /** @deprecated use target.anchor. */
  targetAnchor?: { row: Row; column: Column };
  /** @deprecated old runtime status; refresh status is derived. */
  refreshRevision?: number;
  lastRefreshedAt?: string;
  /** @deprecated Pivot controls are moving to floating-object records. */
  slicers?: PivotSlicer[];
  /** @deprecated Pivot controls are moving to floating-object records. */
  timelines?: PivotTimeline[];
  /** @deprecated chart relation is derived from drawing payloads. */
  chartReferences?: PivotChartReference[];
}

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
  /** @deprecated use fieldId. */
  field?: string;
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
export function migratePivotDefinition(input: PivotModel): PivotDefinition {
  const source = input.source ?? input.dataSource ?? (input.sourceRange ? { kind: 'worksheet-range' as const, range: structuredClone(input.sourceRange) } : undefined);
  if (!source) throw new Error(`Pivot ${input.id} requires a source`);
  const sheetId = input.target?.sheetId ?? input.sheetId;
  if (!sheetId) throw new Error(`Pivot ${input.id} requires a target sheet`);
  const catalog: PivotFieldCatalog = {
    schema: 'PivotFieldCatalog',
    fields: structuredClone(input.fieldCatalog?.fields ?? []).map((field, ordinal) => ({
      ...field,
      fieldId: field.fieldId ?? field.id ?? `field:${ordinal}:${field.name}`,
      id: field.fieldId ?? field.id ?? `field:${ordinal}:${field.name}`,
      ordinal,
    })),
  };
  const fieldId = (value: string | undefined): string | undefined => {
    if (!value) return undefined;
    return catalog.fields.find((field) => field.fieldId === value || field.id === value || field.name === value)?.fieldId ?? value;
  };
  const layout: PivotLayout = structuredClone(input.layout);
  layout.rows = layout.rows.map((entry) => ({ ...entry, fieldId: fieldId(entry.fieldId ?? entry.field) }));
  layout.columns = layout.columns.map((entry) => ({ ...entry, fieldId: fieldId(entry.fieldId ?? entry.field) }));
  layout.values = layout.values.map((entry) => ({ ...entry, fieldId: fieldId(entry.fieldId ?? entry.field), baseFieldId: fieldId(entry.baseFieldId ?? entry.baseField) }));
  layout.filters = layout.filters.map((entry) => {
    const normalizedFieldId = fieldId(entry.fieldId ?? entry.field);
    if (entry.kind !== 'manual') return { ...entry, fieldId: normalizedFieldId };
    const mode = entry.mode ?? (entry.exclude ? 'exclude' : (entry.selected?.length ? 'include' : 'all'));
    const memberKeys = entry.memberKeys?.length ? entry.memberKeys : (entry.selected ?? []).map(createPivotMemberKey);
    return { kind: 'manual' as const, fieldId: normalizedFieldId, mode, memberKeys };
  });
  return {
    schema: PIVOT_DEFINITION_SCHEMA,
    id: input.id,
    source: structuredClone(source),
    target: { sheetId, anchor: structuredClone(input.target?.anchor ?? input.targetAnchor ?? { row: 0, column: 0 }) },
    fieldCatalog: catalog,
    layout,
    refreshPolicy: structuredClone(input.refreshPolicy ?? { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true }),
    nativeMetadata: input.nativeMetadata ? structuredClone(input.nativeMetadata) : undefined,
  };
}
