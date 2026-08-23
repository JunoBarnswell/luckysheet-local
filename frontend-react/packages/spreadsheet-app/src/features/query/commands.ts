import type { CellData } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  buildQueryLoadPlan,
  InMemoryWorkbookTableQueryStore,
  type QueryCellLoadPayload,
  type QueryLoadCommandPayload,
  type QueryWorkbookTableLoadPayload,
  type WorkbookTableQueryStore,
} from './runtime';
import type { QueryResult } from './index';
import { type LoadTarget, type QueryDefinition } from './query-steps';

export interface QueryLoadParams extends QueryLoadCommandPayload {}

export interface QueryRefreshParams {
  queryId: string;
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

interface QueryCellRestorePayload extends QueryCellLoadPayload {
  previousCells: Array<{ row: number; column: number; value?: CellData }>;
}

interface QueryWorkbookTableRestorePayload extends QueryWorkbookTableLoadPayload {
  previousTable?: QueryWorkbookTableLoadPayload['table'];
  previousRecord?: ReturnType<WorkbookTableQueryStore['get']>;
}

type QueryMutationPayload = QueryCellLoadPayload | QueryCellRestorePayload | QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload;

export interface QueryCommandOptions {
  tableStore?: WorkbookTableQueryStore;
}

function applyCells(context: CommandContext, payload: QueryCellLoadPayload): void {
  const sheet = context.workbook.getSheet(payload.clearRange.sheetId);
  for (let row = payload.clearRange.startRow; row <= payload.clearRange.endRow; row += 1) {
    for (let column = payload.clearRange.startColumn; column <= payload.clearRange.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (let rowOffset = 0; rowOffset < payload.values.length; rowOffset += 1) {
    const row = payload.values[rowOffset] ?? [];
    for (let columnOffset = 0; columnOffset < row.length; columnOffset += 1) {
      const value = row[columnOffset];
      if (value) sheet.cells.set(payload.clearRange.startRow + rowOffset, payload.clearRange.startColumn + columnOffset, structuredClone(value));
    }
  }
  if (payload.pivot) {
    const pivot = context.workbook.getSheet(payload.pivot.sheetId).pivots.find((entry) => entry.id === payload.pivot!.pivotId);
    if (!pivot) throw new Error(`Unknown pivot: ${payload.pivot.pivotId}`);
    pivot.refreshRevision = payload.pivot.nextRefreshRevision;
    pivot.lastRefreshedAt = payload.pivot.nextRefreshedAt;
  }
}

function restoreCells(context: CommandContext, payload: QueryCellRestorePayload): void {
  const sheet = context.workbook.getSheet(payload.clearRange.sheetId);
  for (let row = payload.clearRange.startRow; row <= payload.clearRange.endRow; row += 1) {
    for (let column = payload.clearRange.startColumn; column <= payload.clearRange.endColumn; column += 1) sheet.cells.delete(row, column);
  }
  for (const previous of payload.previousCells) {
    if (previous.value) sheet.cells.set(previous.row, previous.column, structuredClone(previous.value));
  }
  if (payload.pivot) {
    const pivot = context.workbook.getSheet(payload.pivot.sheetId).pivots.find((entry) => entry.id === payload.pivot!.pivotId);
    if (pivot) {
      pivot.refreshRevision = payload.pivot.nextRefreshRevision;
      pivot.lastRefreshedAt = payload.pivot.nextRefreshedAt;
    }
  }
}

function applyWorkbookTable(context: CommandContext, payload: QueryWorkbookTableLoadPayload, store: WorkbookTableQueryStore): void {
  context.workbook.tables.set(payload.table.id, structuredClone(payload.table));
  store.set(payload.tableId, { result: structuredClone(payload.result), sourceRevision: payload.sourceRevision });
}

function restoreWorkbookTable(context: CommandContext, payload: QueryWorkbookTableRestorePayload, store: WorkbookTableQueryStore): void {
  if (payload.previousTable) context.workbook.tables.set(payload.tableId, structuredClone(payload.previousTable));
  else context.workbook.tables.delete(payload.tableId);
  if (payload.previousRecord) store.set(payload.tableId, structuredClone(payload.previousRecord));
  else store.delete(payload.tableId);
}

function registerQueryMutations(registry: CommandRegistry, store: WorkbookTableQueryStore): void {
  for (const id of ['query.load.range', 'query.load.sheet-table', 'query.load.pivot-source'] as const) {
    registry.registerMutation<QueryCellLoadPayload | QueryCellRestorePayload>(id, (item, context) => {
      const payload = item.params as QueryCellLoadPayload | QueryCellRestorePayload;
      if ('previousCells' in payload) restoreCells(context, payload);
      else applyCells(context, payload);
    });
  }
  registry.registerMutation<QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload>('query.load.workbook-table', (item, context) => {
    const payload = item.params as QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload;
    if ('previousTable' in payload || 'previousRecord' in payload) restoreWorkbookTable(context, payload, store);
    else applyWorkbookTable(context, payload, store);
  });
}

function executeLoad(
  registry: CommandRegistry,
  params: QueryLoadCommandPayload,
  context: CommandContext,
  store: WorkbookTableQueryStore,
): CommandResult {
  const plan = buildQueryLoadPlan(context.workbook, params, store);
  const sheetId = params.target.sheetId ?? context.workbook.activeSheetId;
  context.applyMutation({
    id: plan.mutationId,
    unitId: context.workbook.unitId,
    sheetId,
    params: plan.payload,
    affectedRanges: plan.affectedRanges,
    inverse: [{ id: plan.mutationId, unitId: context.workbook.unitId, sheetId, params: plan.inverse, affectedRanges: plan.affectedRanges }],
    apply: () => registry.getMutation<QueryMutationPayload>(plan.mutationId)({
      id: plan.mutationId,
      unitId: context.workbook.unitId,
      sheetId,
      params: plan.payload,
      affectedRanges: plan.affectedRanges,
    }, context),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges: plan.affectedRanges };
}

export function registerQueryCommands(registry: CommandRegistry, options: QueryCommandOptions = {}): void {
  const store = options.tableStore ?? new InMemoryWorkbookTableQueryStore();
  registerQueryMutations(registry, store);
  registry.registerCommand<QueryLoadParams>({
    id: 'query.load',
    execute: (params, context) => executeLoad(registry, params, context, store),
  });

  registry.registerCommand<QueryRefreshParams>({
    id: 'query.refresh',
    execute: (params, context) => executeLoad(registry, { query: params.query, target: params.target, result: params.result }, context, store),
  });
}
