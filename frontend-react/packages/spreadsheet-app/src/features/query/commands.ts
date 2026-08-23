import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { TableScalar } from '@react-sheets/core-model';
import type { ConnectorRegistry } from './index';
import { type LoadTarget, type QueryDefinition } from './query-steps';

export interface QueryLoadParams {
  query: QueryDefinition;
  target: LoadTarget;
}

export interface QueryRefreshParams {
  queryId: string;
  query: QueryDefinition;
  target: LoadTarget;
}

export function registerQueryCommands(registry: CommandRegistry, connectors: ConnectorRegistry): void {
  registry.registerMutation('query.load-data', (item, context) => {
    const params = item.params as { target: LoadTarget; columns: string[]; rows: TableScalar[][] };
    const { target, columns, rows } = params;
    if (target.kind === 'range' && target.sheetId && target.range) {
      const sheet = context.workbook.getSheet(target.sheetId);
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < columns.length; c++) {
          const value = rows[r]?.[c] ?? null;
          sheet.cells.set(target.range.startRow + r, target.range.startColumn + c, { value });
        }
      }
    }
    if (target.kind === 'workbook-table' && target.tableId) {
      const table = context.workbook.getTable(target.tableId);
      table.rowCount = rows.length;
      table.fields = columns.map((name, i) => ({ id: `f-${i}`, name, type: 'text' as const, ordinal: i }));
    }
  });

  registry.registerCommand({
    id: 'query.load',
    execute(params: QueryLoadParams, context): CommandResult {
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
    },
  });

  registry.registerCommand({
    id: 'query.refresh',
    execute(params: QueryRefreshParams, context): CommandResult {
      registry.getCommand<QueryLoadParams>('query.load').execute(
        { query: params.query, target: params.target },
        context,
      );
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
    },
  });
}
