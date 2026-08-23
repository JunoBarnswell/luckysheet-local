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

export interface QueryDefinition {
  id: string;
  name: string;
  connectorId: string;
  connectorConfig: Record<string, unknown>;
  steps: QueryStep[];
  refreshOnOpen?: boolean;
}

export type LoadTargetKind = 'range' | 'sheet-table' | 'workbook-table' | 'pivot-source';

export interface LoadTarget {
  kind: LoadTargetKind;
  sheetId?: string;
  range?: { startRow: number; startColumn: number };
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
      default:
        return input;
    }
  }
}
