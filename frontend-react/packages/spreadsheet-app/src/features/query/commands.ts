import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { SetRangeValuesParams } from '@react-sheets/sheet-features';
import type { QueryResult } from './index';
import { queryResultToRangeValues } from './runtime';
import { type LoadTarget, type QueryDefinition } from './query-steps';

export interface QueryLoadParams {
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

export interface QueryRefreshParams {
  queryId: string;
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

function writeQueryResult(
  registry: CommandRegistry,
  params: QueryLoadParams,
  context: CommandContext,
): CommandResult {
  const sheetId = params.target.sheetId ?? context.workbook.activeSheetId;
  const startRow = params.target.range?.startRow ?? 0;
  const startColumn = params.target.range?.startColumn ?? 0;
  const values = queryResultToRangeValues(params.result);
  const rangeCommand = registry.getCommand<SetRangeValuesParams>('sheet.range.set');
  return rangeCommand.execute({ sheetId, startRow, startColumn, values }, context);
}

export function registerQueryCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'query.load',
    execute(params: QueryLoadParams, context): CommandResult {
      return writeQueryResult(registry, params, context);
    },
  });

  registry.registerCommand({
    id: 'query.refresh',
    execute(params: QueryRefreshParams, context): CommandResult {
      return writeQueryResult(registry, { query: params.query, target: params.target, result: params.result }, context);
    },
  });
}
