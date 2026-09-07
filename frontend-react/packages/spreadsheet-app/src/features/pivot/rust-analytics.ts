import type {
  PivotDefinition,
  PivotFieldCatalog,
  PivotResultCell,
  PivotResultNode,
  PivotResultTree,
  PivotResultValueField,
  PivotScalar,
  PivotFilter,
  RangeRef,
} from '@react-sheets/core-model';
import { PIVOT_RESULT_TREE_SCHEMA, createPivotMemberKey, formatPivotMember, pivotScalarFromMemberKey } from '@react-sheets/core-model';
import { kernelInvoke } from '@react-sheets/kernel-client';

/** Canonical wire shape owned by kernel/analytics. Keep this type local so the
 * presentation tree never becomes an analytics input or a second read model. */
export interface PivotAnalyticsAxisField {
  column: number;
  group?: unknown;
  sort?: unknown;
  subtotal?: unknown;
}

export interface PivotAnalyticsValueField {
  valueId: string;
  column: number;
  aggregate: string;
  name?: string;
  showAs?: unknown;
}

export interface PivotAnalyticsFilterColumn {
  column: number;
  predicate: unknown;
}

export interface PivotAnalyticsValueFilter {
  axis: 'rows' | 'columns';
  depth: number;
  valueId: string;
  predicate?: unknown;
  top?: { direction: 'top' | 'bottom'; mode: 'items' | 'percent' | 'sum'; threshold: number };
}

export interface PivotAnalyticsRequest {
  kind: 'pivot';
  revision: number;
  source: RangeRef;
  rowFields: PivotAnalyticsAxisField[];
  columnFields: PivotAnalyticsAxisField[];
  valueFields: PivotAnalyticsValueField[];
  filters: PivotAnalyticsFilterColumn[];
  valueFilters: PivotAnalyticsValueFilter[];
  includeRowTotals: boolean;
  includeColumnTotals: boolean;
  includeSubtotals: boolean;
  viewport?: { rowOffset: number; columnOffset: number; rowLimit: number; columnLimit: number };
  drilldown?: { rowKeys: PivotScalar[]; columnKeys: PivotScalar[]; offset: number; limit: number };
  budget?: unknown;
}

export interface PivotAnalyticsSparseResult {
  kind?: 'pivot';
  revision: number;
  rows: Array<{ rowId: number; keys: PivotScalar[]; subtotal: boolean; grandTotal: boolean }>;
  columns: Array<{ columnId: number; keys: PivotScalar[] } | PivotScalar[]>;
  cells: Array<{ rowId: number; columnId: number; values: PivotScalar[] }>;
  totalGroups: number;
  totalColumns?: number;
  fields?: PivotAnalyticsValueField[];
  drilldown?: { rows?: Array<{ row: number; sourceId?: string; recordId?: string }>; total?: number };
}

export interface ExecutePivotAnalyticsOptions {
  unitId: string;
  revision: number;
  source: RangeRef;
  definition: PivotDefinition;
  filters?: PivotAnalyticsFilterColumn[];
  valueFilters?: PivotAnalyticsValueFilter[];
  viewport?: PivotAnalyticsRequest['viewport'];
  drilldown?: PivotAnalyticsRequest['drilldown'];
  budget?: unknown;
}

function fieldOrdinal(definition: PivotDefinition, fieldId: string): number {
  const field = definition.fieldCatalog.fields.find((candidate) => candidate.fieldId === fieldId);
  if (!field) throw new Error(`Pivot analytics field is missing from the catalog: ${fieldId}`);
  return field.ordinal;
}

function aggregateName(value: string): string {
  if (['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'].includes(value)) return value;
  throw new Error(`Unsupported Pivot aggregate: ${value}`);
}

function sourceBody(source: RangeRef): RangeRef {
  // The canonical source contract includes one header row. Rust receives the
  // body range so field ordinals and row provenance cannot drift by one row.
  if (source.endRow <= source.startRow) throw new Error('Pivot source must contain a header and at least one body row');
  return { ...source, startRow: source.startRow + 1 };
}

function conditionPredicate(filter: Extract<PivotFilter, { kind: 'condition' }>): unknown {
  if (filter.dynamic) return { kind: 'dynamic', type: filter.dynamic.replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase()) };
  const operator = filter.operator;
  if (operator === 'between' || operator === 'not-between') {
    if (filter.value2 === undefined) throw new Error('PIVOT_FILTER_INVALID: range filter needs both bounds');
    const predicate = { kind: 'custom', join: 'and', conditions: [{ operator: 'greaterThanOrEqual', value: filter.value }, { operator: 'lessThanOrEqual', value: filter.value2 }] };
    return operator === 'not-between' ? { kind: 'not', predicate } : predicate;
  }
  const operators: Record<string, string> = { equals: 'equals', 'not-equals': 'notEquals', 'begins-with': 'beginsWith', 'ends-with': 'endsWith', contains: 'contains', 'not-contains': 'notContains', 'greater-than': 'greaterThan', 'greater-or-equal': 'greaterThanOrEqual', 'less-than': 'lessThan', 'less-or-equal': 'lessThanOrEqual', before: 'lessThan', after: 'greaterThan' };
  if (operator === 'not-begins-with' || operator === 'not-ends-with') return { kind: 'not', predicate: { kind: 'custom', join: 'and', conditions: [{ operator: operator === 'not-begins-with' ? 'beginsWith' : 'endsWith', value: filter.value }] } };
  const resolved = operators[operator];
  if (!resolved) throw new Error(`PIVOT_FILTER_INVALID: unknown operator ${operator}`);
  return { kind: 'custom', join: 'and', conditions: [{ operator: resolved, value: filter.value }] };
}

export function createPivotAnalyticsRequest(options: ExecutePivotAnalyticsOptions): PivotAnalyticsRequest {
  const { definition } = options;
  const layout = definition.layout;
  const column = (fieldId: string) => options.source.startColumn + fieldOrdinal(definition, fieldId);
  const filters: PivotAnalyticsFilterColumn[] = [];
  const valueFilters: PivotAnalyticsValueFilter[] = [];
  for (const filter of layout.filters) {
    if (filter.kind === 'manual') {
      if (filter.mode === 'all') continue;
      const values = filter.memberKeys.map(pivotScalarFromMemberKey);
      const predicate = { kind: 'values', values: values.filter((value) => value !== null), includeBlank: values.some((value) => value === null) };
      filters.push({ column: column(filter.fieldId), predicate: filter.mode === 'exclude' ? { kind: 'not', predicate } : predicate });
      continue;
    }
    if (filter.kind === 'top-items' || filter.family === 'value') {
      const rowDepth = layout.rows.findIndex((placement) => placement.fieldId === filter.fieldId);
      const columnDepth = layout.columns.findIndex((placement) => placement.fieldId === filter.fieldId);
      if (rowDepth < 0 && columnDepth < 0) throw new Error(`PIVOT_VALUE_FILTER_INVALID: ${filter.fieldId} is not on an axis`);
      valueFilters.push({ axis: rowDepth >= 0 ? 'rows' : 'columns', depth: rowDepth >= 0 ? rowDepth : columnDepth, valueId: filter.valueId,
        ...(filter.kind === 'top-items' ? { top: { direction: filter.direction, mode: filter.mode, threshold: filter.threshold } } : { predicate: conditionPredicate(filter) }),
      });
      continue;
    }
    filters.push({ column: column(filter.fieldId), predicate: conditionPredicate(filter) });
  }
  const request: PivotAnalyticsRequest = {
    kind: 'pivot',
    revision: options.revision,
    source: sourceBody(options.source),
    rowFields: layout.rows.map((placement) => ({
      column: column(placement.fieldId),
      ...(placement.group ? { group: placement.group } : {}),
      ...(placement.sort ? { sort: placement.sort } : {}),
      ...(placement.subtotal ? { subtotal: placement.subtotal } : {}),
    })),
    columnFields: layout.columns.map((placement) => ({
      column: column(placement.fieldId),
      ...(placement.group ? { group: placement.group } : {}),
      ...(placement.sort ? { sort: placement.sort } : {}),
      ...(placement.subtotal ? { subtotal: placement.subtotal } : {}),
    })),
    valueFields: layout.values.map((value) => ({
      valueId: value.valueId,
      column: column(value.fieldId),
      aggregate: aggregateName(value.summarizeBy),
      ...(value.displayName ? { name: value.displayName } : {}),
      ...(value.showAs ? { showAs: 'baseFieldId' in value.showAs ? { ...Object.fromEntries(Object.entries(value.showAs).filter(([key]) => key !== 'baseFieldId')), baseColumn: column(value.showAs.baseFieldId) } : value.showAs } : {}),
    })),
    filters: [...filters, ...(options.filters ?? [])],
    valueFilters: [...valueFilters, ...(options.valueFilters ?? [])],
    includeRowTotals: layout.showRowGrandTotals,
    includeColumnTotals: layout.showColumnGrandTotals,
    includeSubtotals: layout.subtotalLocation !== 'off',
    ...(options.viewport ? { viewport: options.viewport } : {}),
    ...(options.drilldown ? { drilldown: options.drilldown } : {}),
    ...(options.budget ? { budget: options.budget } : {}),
  };
  if (!request.valueFields.length) throw new Error('PIVOT_VALUES_REQUIRED: Pivot requires at least one value field');
  return request;
}

export function executePivotAnalytics(options: ExecutePivotAnalyticsOptions): PivotResultTree {
  const request = createPivotAnalyticsRequest(options);
  const response = kernelInvoke<PivotAnalyticsSparseResult & { kind: 'pivot' }>('analytics.execute', {
    unitId: options.unitId,
    revision: options.revision,
    request,
  });
  if (response.kind !== 'pivot' || response.revision !== options.revision) {
    throw new Error('KERNEL_ANALYTICS_RESPONSE_INVALID: pivot response revision or kind is invalid');
  }
  return pivotTreeFromSparseResult(options.definition, response);
}

function columnKeys(column: PivotAnalyticsSparseResult['columns'][number]): PivotScalar[] {
  return Array.isArray(column) ? column : column.keys;
}

function cellFor(rowId: number, columnId: number, result: PivotAnalyticsSparseResult, definition: PivotDefinition): PivotResultCell {
  const found = result.cells.find((cell) => cell.rowId === rowId && cell.columnId === columnId);
  return {
    id: `${definition.id}|row:${rowId}|column:${columnId}`,
    kind: 'detail',
    columnPath: columnKeys(result.columns[columnId] ?? []),
    values: found?.values ?? definition.layout.values.map(() => null),
    sourceRowPaths: [],
  };
}

function buildNodes(definition: PivotDefinition, result: PivotAnalyticsSparseResult, rowIds: number[]): PivotResultNode[] {
  const rows = result.rows.filter((row) => rowIds.includes(row.rowId) && !row.grandTotal);
  const roots: PivotResultNode[] = [];
  const byPath = new Map<string, PivotResultNode>();
  for (const row of rows) {
    let parent: PivotResultNode | undefined;
    for (let depth = 0; depth < row.keys.length; depth += 1) {
      const path = row.keys.slice(0, depth + 1).map((key) => JSON.stringify(key)).join('\u001f');
      let node = byPath.get(path);
      if (!node) {
        const placement = definition.layout.rows[depth];
        const key = row.keys[depth] ?? null;
        node = { nodeId: `${definition.id}|${path}`, path: path ? path.split('\u001f') : [], kind: row.subtotal ? 'subtotal' : 'leaf', fieldId: placement?.fieldId, memberKey: createPivotMemberKey(key), key, label: formatPivotMember(key), depth, children: [], values: [], subtotal: row.subtotal, sourceRowPaths: [] };
        byPath.set(path, node);
        if (parent) parent.children.push(node); else roots.push(node);
      }
      parent = node;
    }
    if (!parent) {
      const key = 'Values';
      parent = byPath.get(key) ?? { nodeId: `${definition.id}|root`, path: ['__root__'], kind: 'leaf', key: null, label: key, depth: 0, children: [], values: [], subtotal: false, sourceRowPaths: [] };
      if (!byPath.has(key)) { byPath.set(key, parent); roots.push(parent); }
    }
    parent.values = result.columns.map((_, columnId) => cellFor(row.rowId, columnId, result, definition));
  }
  return roots;
}

export function pivotTreeFromSparseResult(definition: PivotDefinition, result: PivotAnalyticsSparseResult): PivotResultTree {
  const grand = result.rows.find((row) => row.grandTotal);
  const resultFields: PivotResultValueField[] = definition.layout.values.map((value) => ({ ...value, sourceFieldId: value.fieldId }));
  const grandTotal = grand ? { ...cellFor(grand.rowId, 0, result, definition), id: `${definition.id}|grand-total`, kind: 'grand-total' as const, columnPath: [] } : null;
  return {
    schema: PIVOT_RESULT_TREE_SCHEMA,
    pivotId: definition.id,
    fields: definition.fieldCatalog,
    columnPaths: result.columns.map(columnKeys),
    valueFields: resultFields,
    rows: buildNodes(definition, result, result.rows.map((row) => row.rowId)),
    grandTotal,
    columnGrandTotals: [],
    sourceRowPaths: [],
    sourceRevision: String(result.revision),
  };
}
