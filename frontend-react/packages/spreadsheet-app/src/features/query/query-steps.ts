export type QueryStepKind =
  | 'source'
  | 'filter'
  | 'select-columns'
  | 'rename-column'
  | 'sort'
  | 'group-by'
  | 'join'
  | 'pivot'
  | 'custom';

export interface QueryStep {
  id: string;
  kind: QueryStepKind;
  name: string;
  config: Record<string, unknown>;
  enabled: boolean;
}

export type QueryRefreshMode = 'manual' | 'on-open' | 'interval';

export interface QueryRefreshPolicy {
  mode: QueryRefreshMode;
  intervalMs?: number;
}

export interface QueryDefinition {
  id: string;
  name: string;
  connectorId: string;
  connectorConfig: Record<string, unknown>;
  steps: QueryStep[];
  refreshOnOpen?: boolean;
  sourceRevision?: number;
  refreshPolicy?: QueryRefreshPolicy;
  /** Persisted so a rehydrated session can refresh into the same destination. */
  lastTarget?: LoadTarget;
}

export type LoadTargetKind = 'range' | 'sheet-table' | 'workbook-table' | 'pivot-source';

export interface LoadTarget {
  kind: LoadTargetKind;
  sheetId?: string;
  range?: { startRow: number; startColumn: number; endRow?: number; endColumn?: number };
  tableId?: string;
  pivotId?: string;
}

type Scalar = string | number | boolean | null;
type Table = { columns: string[]; rows: Scalar[][] };

const IMPLEMENTED_QUERY_STEP_KINDS = new Set<QueryStepKind>([
  'source', 'filter', 'select-columns', 'rename-column', 'sort', 'group-by', 'join', 'pivot',
]);
const FILTER_OPERATORS = new Set(['eq', 'neq', 'contains', 'startsWith', 'endsWith', 'gt', 'gte', 'lt', 'lte', 'isNull', 'notNull']);
const AGGREGATIONS = new Set(['sum', 'count', 'average', 'min', 'max']);

export class QueryStepPipeline {
  private steps: QueryStep[];

  constructor(steps: QueryStep[]) {
    this.steps = structuredClone(steps);
  }

  reorder(fromIndex: number, toIndex: number): void {
    const [step] = this.steps.splice(fromIndex, 1);
    if (!step) return;
    this.steps.splice(Math.max(0, Math.min(toIndex, this.steps.length)), 0, step);
  }

  add(step: QueryStep): void {
    validateQuerySteps([step]);
    this.steps.push(structuredClone(step));
  }

  remove(stepId: string): void {
    this.steps = this.steps.filter((step) => step.id !== stepId);
  }

  getSteps(): QueryStep[] {
    return structuredClone(this.steps);
  }

  applySteps(input: { columns: string[]; rows: unknown[][] }): { columns: string[]; rows: unknown[][] } {
    validateQuerySteps(this.steps);
    let current: Table = { columns: [...input.columns], rows: input.rows.map((row) => row.map(toScalar)) };
    for (const step of this.steps) {
      if (!step.enabled || step.kind === 'source') continue;
      current = this.applyStep(step, current);
    }
    return current;
  }

  private applyStep(step: QueryStep, input: Table): Table {
    switch (step.kind) {
      case 'filter': return filterTable(input, step.config, step.id);
      case 'select-columns': return selectColumns(input, step.config, step.id);
      case 'rename-column': return renameColumn(input, step.config, step.id);
      case 'sort': return sortTable(input, step.config, step.id);
      case 'group-by': return groupTable(input, step.config, step.id);
      case 'join': return joinTable(input, step.config, step.id);
      case 'pivot': return pivotTable(input, step.config, step.id);
      case 'custom': throw new Error(`Query step "${step.id}" of kind "custom" is not implemented`);
      case 'source': return input;
      default: throw new Error(`Query step "${step.id}" has an unsupported kind`);
    }
  }
}

function toScalar(value: unknown): Scalar {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  throw new Error('Query transforms only support scalar cell values');
}

function requireString(config: Record<string, unknown>, key: string, stepId: string): string {
  const value = config[key];
  if (typeof value !== 'string' || !value.trim()) throw new Error(`Query step "${stepId}" requires a non-empty config.${key}`);
  return value;
}

function requireStringArray(config: Record<string, unknown>, keys: string[], stepId: string): string[] {
  for (const key of keys) {
    const value = config[key];
    if (Array.isArray(value) && value.length > 0 && value.every((entry) => typeof entry === 'string' && entry.trim())) return value.map((entry) => String(entry));
  }
  throw new Error(`Query step "${stepId}" requires a non-empty column list`);
}

function columnIndex(columns: string[], name: string, stepId: string): number {
  const index = columns.indexOf(name);
  if (index < 0) throw new Error(`Query step "${stepId}" references missing column "${name}"`);
  return index;
}

function filterTable(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const name = requireString(config, 'column', stepId);
  const index = columnIndex(input.columns, name, stepId);
  const operator = typeof config.operator === 'string' ? config.operator : 'eq';
  if (!FILTER_OPERATORS.has(operator)) throw new Error(`Query step "${stepId}" has unsupported filter operator "${operator}"`);
  const expected = config.value as Scalar;
  const caseSensitive = config.caseSensitive !== false;
  const normalize = (value: Scalar) => caseSensitive || typeof value !== 'string' ? value : value.toLocaleLowerCase();
  const compare = (left: Scalar, right: Scalar): number => {
    if (left === right) return 0;
    if (left === null) return -1;
    if (right === null) return 1;
    return left < right ? -1 : 1;
  };
  return {
    columns: [...input.columns],
    rows: input.rows.filter((row) => {
      const actual = normalize(row[index] ?? null);
      const wanted = normalize(expected ?? null);
      switch (operator) {
        case 'eq': return actual === wanted;
        case 'neq': return actual !== wanted;
        case 'contains': return typeof actual === 'string' && typeof wanted === 'string' && actual.includes(wanted);
        case 'startsWith': return typeof actual === 'string' && typeof wanted === 'string' && actual.startsWith(wanted);
        case 'endsWith': return typeof actual === 'string' && typeof wanted === 'string' && actual.endsWith(wanted);
        case 'gt': return compare(actual, wanted) > 0;
        case 'gte': return compare(actual, wanted) >= 0;
        case 'lt': return compare(actual, wanted) < 0;
        case 'lte': return compare(actual, wanted) <= 0;
        case 'isNull': return actual === null;
        case 'notNull': return actual !== null;
        default: return false;
      }
    }).map((row) => [...row]),
  };
}

function selectColumns(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const selected = requireStringArray(config, ['columns'], stepId);
  const indices = selected.map((name) => columnIndex(input.columns, name, stepId));
  return { columns: [...selected], rows: input.rows.map((row) => indices.map((index) => row[index] ?? null)) };
}

function renameColumn(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const from = requireString(config, 'from', stepId);
  const to = requireString(config, 'to', stepId);
  const index = columnIndex(input.columns, from, stepId);
  if (input.columns.includes(to) && from !== to) throw new Error(`Query step "${stepId}" cannot rename to existing column "${to}"`);
  const columns = [...input.columns];
  columns[index] = to;
  return { columns, rows: input.rows.map((row) => [...row]) };
}

function sortTable(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const criteria = Array.isArray(config.by)
    ? config.by.map((entry) => {
      if (!entry || typeof entry !== 'object') throw new Error(`Query step "${stepId}" has invalid sort criteria`);
      const item = entry as Record<string, unknown>;
      return { column: requireString(item, 'column', stepId), ascending: item.ascending !== false };
    })
    : [{ column: requireString(config, 'column', stepId), ascending: config.ascending !== false }];
  const indexed = input.rows.map((row, index) => ({ row, index }));
  indexed.sort((left, right) => {
    for (const criterion of criteria) {
      const index = columnIndex(input.columns, criterion.column, stepId);
      const a = left.row[index] ?? null;
      const b = right.row[index] ?? null;
      if (a === b) continue;
      const comparison = a === null ? -1 : b === null ? 1 : a < b ? -1 : 1;
      return criterion.ascending ? comparison : -comparison;
    }
    return left.index - right.index;
  });
  return { columns: [...input.columns], rows: indexed.map((entry) => [...entry.row]) };
}

interface AggregateSpec { column: string; fn: 'sum' | 'count' | 'average' | 'min' | 'max'; as: string }

function aggregationSpecs(config: Record<string, unknown>, input: Table, stepId: string): AggregateSpec[] {
  const raw = Array.isArray(config.aggregations) ? config.aggregations : [];
  if (raw.length === 0) return [{ column: '*', fn: 'count', as: 'count' }];
  return raw.map((entry, index) => {
    if (!entry || typeof entry !== 'object') throw new Error(`Query step "${stepId}" has invalid aggregation ${index}`);
    const item = entry as Record<string, unknown>;
    const fn = String(item.function ?? item.fn ?? 'sum') as AggregateSpec['fn'];
    if (!AGGREGATIONS.has(fn)) throw new Error(`Query step "${stepId}" has unsupported aggregation "${fn}"`);
    const column = item.column === '*' ? '*' : requireString(item, 'column', stepId);
    if (column !== '*') columnIndex(input.columns, column, stepId);
    const as = typeof item.as === 'string' && item.as.trim() ? item.as : `${fn}_${column}`;
    return { column, fn, as };
  });
}

function aggregate(values: Scalar[], fn: AggregateSpec['fn']): Scalar {
  if (fn === 'count') return values.filter((value) => value !== null).length;
  const numeric = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
  if (fn === 'sum') return numeric.reduce((total, value) => total + value, 0);
  if (fn === 'average') return numeric.length === 0 ? null : numeric.reduce((total, value) => total + value, 0) / numeric.length;
  if (numeric.length === 0) return null;
  return fn === 'min' ? Math.min(...numeric) : Math.max(...numeric);
}

function groupTable(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const groups = requireStringArray(config, ['by', 'columns', 'groupBy'], stepId);
  const groupIndices = groups.map((name) => columnIndex(input.columns, name, stepId));
  const specs = aggregationSpecs(config, input, stepId);
  const buckets = new Map<string, { key: Scalar[]; rows: Scalar[][] }>();
  for (const row of input.rows) {
    const key = groupIndices.map((index) => row[index] ?? null);
    const serialized = JSON.stringify(key);
    const bucket = buckets.get(serialized) ?? { key, rows: [] };
    bucket.rows.push(row);
    buckets.set(serialized, bucket);
  }
  return {
    columns: [...groups, ...specs.map((spec) => spec.as)],
    rows: [...buckets.values()].map((bucket) => [
      ...bucket.key,
      ...specs.map((spec) => aggregate(spec.column === '*' ? bucket.rows.map(() => 1) : bucket.rows.map((row) => row[columnIndex(input.columns, spec.column, stepId)] ?? null), spec.fn)),
    ]),
  };
}

function tableFromConfig(value: unknown, stepId: string): Table {
  if (Array.isArray(value)) {
    if (value.some((row) => !row || typeof row !== 'object' || Array.isArray(row))) throw new Error(`Query step "${stepId}" right records are invalid`);
    const result = recordsToTable(value as Array<Record<string, unknown>>);
    return result;
  }
  if (!value || typeof value !== 'object') throw new Error(`Query step "${stepId}" requires a right table`);
  const record = value as Record<string, unknown>;
  if (Array.isArray(record.data)) return recordsToTable(record.data as Array<Record<string, unknown>>);
  if (!Array.isArray(record.columns) || !record.columns.every((column) => typeof column === 'string')) throw new Error(`Query step "${stepId}" right table columns are invalid`);
  if (!Array.isArray(record.rows) || !record.rows.every((row) => Array.isArray(row))) throw new Error(`Query step "${stepId}" right table rows are invalid`);
  const columns = record.columns as string[];
  const rows = (record.rows as unknown[][]).map((row) => {
    if (row.length !== columns.length) throw new Error(`Query step "${stepId}" right table row width is invalid`);
    return row.map(toScalar);
  });
  return { columns: [...columns], rows };
}

function recordsToTable(records: Array<Record<string, unknown>>): Table {
  const columns = [...new Set(records.flatMap((record) => Object.keys(record)))];
  return { columns, rows: records.map((record) => columns.map((column) => toScalar(record[column] ?? null))) };
}

function joinTable(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const right = tableFromConfig(config.right ?? config.rightTable, stepId);
  const leftOn = requireStringArray(config, ['leftOn', 'on'], stepId);
  const rightOn = Array.isArray(config.rightOn) ? config.rightOn.map((value) => String(value)) : leftOn;
  if (leftOn.length !== rightOn.length) throw new Error(`Query step "${stepId}" join keys must have equal length`);
  const leftIndices = leftOn.map((name) => columnIndex(input.columns, name, stepId));
  const rightIndices = rightOn.map((name) => columnIndex(right.columns, name, stepId));
  const rightKey = (row: Scalar[]) => JSON.stringify(rightIndices.map((index) => row[index] ?? null));
  const index = new Map<string, Scalar[][]>();
  for (const row of right.rows) index.set(rightKey(row), [...(index.get(rightKey(row)) ?? []), row]);
  const type = config.type === 'left' || config.type === 'full' ? config.type : 'inner';
  const rightColumns = right.columns.map((column) => input.columns.includes(column) ? `${column}_right` : column);
  const rows: Scalar[][] = [];
  const matchedRight = new Set<Scalar[]>();
  for (const left of input.rows) {
    const key = JSON.stringify(leftIndices.map((index) => left[index] ?? null));
    const matches = index.get(key) ?? [];
    if (matches.length === 0) {
      if (type !== 'inner') rows.push([...left, ...right.columns.map(() => null)]);
      continue;
    }
    for (const match of matches) {
      matchedRight.add(match);
      rows.push([...left, ...match]);
    }
  }
  if (type === 'full') for (const row of right.rows) if (!matchedRight.has(row)) rows.push([...input.columns.map(() => null), ...row]);
  return { columns: [...input.columns, ...rightColumns], rows };
}

function pivotTable(input: Table, config: Record<string, unknown>, stepId: string): Table {
  const rowFields = requireStringArray(config, ['rows', 'rowFields'], stepId);
  const columnFields = requireStringArray(config, ['columns', 'columnFields'], stepId);
  const valueFields = requireStringArray(config, ['values', 'valueFields'], stepId);
  const rowIndices = rowFields.map((name) => columnIndex(input.columns, name, stepId));
  const colIndices = columnFields.map((name) => columnIndex(input.columns, name, stepId));
  const valueIndices = valueFields.map((name) => columnIndex(input.columns, name, stepId));
  const fn = String(config.aggregation ?? 'sum') as AggregateSpec['fn'];
  if (!AGGREGATIONS.has(fn)) throw new Error(`Query step "${stepId}" has unsupported pivot aggregation "${fn}"`);
  const columnKeys = [...new Set(input.rows.map((row) => JSON.stringify(colIndices.map((index) => row[index] ?? null))))];
  const groups = new Map<string, Scalar[][]>();
  for (const row of input.rows) {
    const key = JSON.stringify(rowIndices.map((index) => row[index] ?? null));
    groups.set(key, [...(groups.get(key) ?? []), row]);
  }
  const columns = [...rowFields];
  for (const columnKey of columnKeys) {
    const values = JSON.parse(columnKey!) as Scalar[];
    for (const valueField of valueFields) columns.push(`${values.join(' / ')} · ${valueField}`);
  }
  const rows = [...groups.values()].map((group) => {
    const first = group[0]!;
    const output: Scalar[] = rowIndices.map((index) => first[index] ?? null);
    for (const columnKey of columnKeys) {
      const matches = group.filter((row) => JSON.stringify(colIndices.map((index) => row[index] ?? null)) === columnKey);
      for (const valueIndex of valueIndices) output.push(aggregate(matches.map((row) => row[valueIndex] ?? null), fn));
    }
    return output;
  });
  return { columns, rows };
}

export function validateQuerySteps(steps: readonly QueryStep[]): void {
  if (!Array.isArray(steps)) throw new Error('Query steps must be an array');
  for (const step of steps) {
    if (!step || typeof step.id !== 'string' || !step.id.trim()) throw new Error('Every query step requires a non-empty id');
    if (!IMPLEMENTED_QUERY_STEP_KINDS.has(step.kind)) throw new Error(`Query step "${step.id}" of kind "${step.kind}" is not implemented`);
    if (typeof step.name !== 'string' || !step.name.trim()) throw new Error(`Query step "${step.id}" requires a name`);
    if (!step.config || typeof step.config !== 'object' || Array.isArray(step.config)) throw new Error(`Query step "${step.id}" has invalid configuration`);
    if (typeof step.enabled !== 'boolean') throw new Error(`Query step "${step.id}" must declare enabled explicitly`);
    switch (step.kind) {
      case 'filter': {
        requireString(step.config, 'column', step.id);
        const operator = typeof step.config.operator === 'string' ? step.config.operator : 'eq';
        if (!FILTER_OPERATORS.has(operator)) throw new Error(`Query step "${step.id}" has unsupported filter operator "${operator}"`);
        break;
      }
      case 'select-columns': requireStringArray(step.config, ['columns'], step.id); break;
      case 'rename-column': requireString(step.config, 'from', step.id); requireString(step.config, 'to', step.id); break;
      case 'sort':
        if (Array.isArray(step.config.by)) {
          if (step.config.by.length === 0) throw new Error(`Query step "${step.id}" requires sort criteria`);
          step.config.by.forEach((entry: unknown) => { if (!entry || typeof entry !== 'object') throw new Error(`Query step "${step.id}" has invalid sort criteria`); requireString(entry as Record<string, unknown>, 'column', step.id); });
        } else requireString(step.config, 'column', step.id);
        break;
      case 'group-by': requireStringArray(step.config, ['by', 'columns', 'groupBy'], step.id); break;
      case 'join': requireStringArray(step.config, ['leftOn', 'on'], step.id); tableFromConfig(step.config.right ?? step.config.rightTable, step.id); break;
      case 'pivot': requireStringArray(step.config, ['rows', 'rowFields'], step.id); requireStringArray(step.config, ['columns', 'columnFields'], step.id); requireStringArray(step.config, ['values', 'valueFields'], step.id); break;
      case 'source': break;
      default: throw new Error(`Query step "${step.id}" is not implemented`);
    }
  }
}
