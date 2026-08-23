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
  /** Monotonic revision supplied by the source system, when available. */
  sourceRevision?: number;
  refreshPolicy?: QueryRefreshPolicy;
}

export type LoadTargetKind = 'range' | 'sheet-table' | 'workbook-table' | 'pivot-source';

export interface LoadTarget {
  kind: LoadTargetKind;
  sheetId?: string;
  range?: { startRow: number; startColumn: number; endRow?: number; endColumn?: number };
  tableId?: string;
  pivotId?: string;
}

export class QueryStepPipeline {
  constructor(private steps: QueryStep[]) {}

  reorder(fromIndex: number, toIndex: number): void {
    const [step] = this.steps.splice(fromIndex, 1);
    if (!step) return;
    this.steps.splice(toIndex, 0, step);
  }

  add(step: QueryStep): void {
    this.steps.push(step);
  }

  remove(stepId: string): void {
    this.steps = this.steps.filter((s) => s.id !== stepId);
  }

  getSteps(): QueryStep[] {
    return [...this.steps];
  }

  /** 按顺序应用 transform steps（filter/select/sort 等） */
  applySteps(input: { columns: string[]; rows: unknown[][] }): { columns: string[]; rows: unknown[][] } {
    validateQuerySteps(this.steps);
    let current = input;
    for (const step of this.steps) {
      if (!step.enabled) continue;
      current = this.applyStep(step, current);
    }
    return current;
  }

  private applyStep(step: QueryStep, input: { columns: string[]; rows: unknown[][] }): { columns: string[]; rows: unknown[][] } {
    switch (step.kind) {
      case 'filter': {
        const column = step.config.column as string;
        const value = step.config.value;
        const colIndex = input.columns.indexOf(column);
        if (colIndex < 0) return input;
        return {
          columns: input.columns,
          rows: input.rows.filter((row) => row[colIndex] === value),
        };
      }
      case 'select-columns': {
        const selected = (step.config.columns as string[]) ?? input.columns;
        const indices = selected.map((c) => input.columns.indexOf(c)).filter((i) => i >= 0);
        return {
          columns: selected.filter((c) => input.columns.includes(c)),
          rows: input.rows.map((row) => indices.map((i) => row[i])),
        };
      }
      case 'rename-column': {
        const from = readStringConfig(step, 'from');
        const to = readStringConfig(step, 'to');
        if (!input.columns.includes(from)) {
          throw new Error(`Query step "${step.id}" cannot rename missing column "${from}"`);
        }
        if (input.columns.includes(to) && to !== from) {
          throw new Error(`Query step "${step.id}" cannot rename to existing column "${to}"`);
        }
        const index = input.columns.indexOf(from);
        return {
          columns: input.columns.map((column, columnIndex) => columnIndex === index ? to : column),
          rows: input.rows.map((row) => [...row]),
        };
      }
      case 'sort': {
        const column = step.config.column as string;
        const ascending = step.config.ascending !== false;
        const colIndex = input.columns.indexOf(column);
        if (colIndex < 0) return input;
        const sorted = [...input.rows].sort((a, b) => {
          const av = a[colIndex];
          const bv = b[colIndex];
          if (av === bv) return 0;
          const cmp = av == null ? -1 : bv == null ? 1 : av < bv ? -1 : 1;
          return ascending ? cmp : -cmp;
        });
        return { columns: input.columns, rows: sorted };
      }
      case 'source':
        return input;
      case 'group-by':
      case 'join':
      case 'pivot':
      case 'custom':
        throw new Error(`Query step "${step.id}" of kind "${step.kind}" is not implemented`);
      default:
        // Keep this exhaustive even when a new kind is added. Unknown steps
        // must fail closed instead of silently passing data through.
        return input;
    }
  }
}

const IMPLEMENTED_QUERY_STEP_KINDS = new Set<QueryStepKind>([
  'source',
  'filter',
  'select-columns',
  'rename-column',
  'sort',
]);

export function validateQuerySteps(steps: readonly QueryStep[]): void {
  for (const step of steps) {
    if (!step || typeof step.id !== 'string' || !step.id.trim()) {
      throw new Error('Every query step requires a non-empty id');
    }
    if (!IMPLEMENTED_QUERY_STEP_KINDS.has(step.kind)) {
      throw new Error(`Query step "${step.id}" of kind "${step.kind}" is not implemented`);
    }
    if (!step.config || typeof step.config !== 'object' || Array.isArray(step.config)) {
      throw new Error(`Query step "${step.id}" has invalid configuration`);
    }
    if (typeof step.enabled !== 'boolean') {
      throw new Error(`Query step "${step.id}" must declare enabled explicitly`);
    }
    if (step.kind === 'filter') {
      readStringConfig(step, 'column');
    } else if (step.kind === 'select-columns') {
      if (!Array.isArray(step.config.columns) || step.config.columns.some((column) => typeof column !== 'string' || !column.trim())) {
        throw new Error(`Query step "${step.id}" requires a columns array`);
      }
    } else if (step.kind === 'rename-column') {
      readStringConfig(step, 'from');
      readStringConfig(step, 'to');
    } else if (step.kind === 'sort') {
      readStringConfig(step, 'column');
      if (step.config.ascending !== undefined && typeof step.config.ascending !== 'boolean') {
        throw new Error(`Query step "${step.id}" ascending must be boolean`);
      }
    }
  }
}

function readStringConfig(step: QueryStep, key: string): string {
  const value = step.config[key];
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(`Query step "${step.id}" requires a non-empty string config.${key}`);
  }
  return value;
}
