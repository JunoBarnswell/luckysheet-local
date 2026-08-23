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
import { serializeQueryDefinition, type QueryDefinitionPersistence, type QueryResult } from './index';
import { type LoadTarget, type QueryDefinition } from './query-steps';

export interface QueryLoadParams extends QueryLoadCommandPayload {}

export interface QueryRefreshParams {
  queryId: string;
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

export interface QueryDefinitionReplaceParams {
  definition: QueryDefinition;
}

interface QueryDefinitionReplaceMutationParams {
  queryId: string;
  definition: QueryDefinitionPersistence | null;
}

interface QueryCellRestorePayload extends QueryCellLoadPayload {
  previousCells: Array<{ row: number; column: number; value?: CellData }>;
}

interface QueryWorkbookTableRestorePayload extends QueryWorkbookTableLoadPayload {
  previousTable?: QueryWorkbookTableLoadPayload['table'];
  previousRecord?: ReturnType<WorkbookTableQueryStore['get']>;
}

type QueryMutationPayload = QueryCellLoadPayload | QueryCellRestorePayload | QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload;

function isQueryDefinitionReplacePayload(value: unknown): value is QueryDefinitionReplaceMutationParams {
  if (!value || typeof value !== 'object') return false;
  const payload = value as Record<string, unknown>;
  return typeof payload.queryId === 'string' && payload.queryId.trim().length > 0
    && (payload.definition === null || (typeof payload.definition === 'object' && (payload.definition as { schema?: unknown }).schema === 'QueryDefinitionV1'));
}

function isQueryLoadPayload(value: unknown): value is QueryMutationPayload {
  if (!value || typeof value !== 'object') return false;
  const payload = value as Record<string, unknown>;
  if (typeof payload.queryId !== 'string' || !payload.queryId.trim()) return false;
  if (!(payload.queryDefinition === null || (typeof payload.queryDefinition === 'object' && (payload.queryDefinition as { schema?: unknown }).schema === 'QueryDefinitionV1'))) return false;
  if (payload.kind === 'workbook-table') return typeof payload.tableId === 'string' && Boolean(payload.table) && Boolean(payload.result);
  if (payload.kind !== 'cells') return false;
  return Boolean(payload.clearRange) && Array.isArray(payload.values);
}

function queryMutationRanges(params: QueryMutationPayload): readonly import('@react-sheets/core-model').RangeRef[] {
  return params.kind === 'cells' ? [params.clearRange] : [];
}

export interface QueryCommandOptions {
  tableStore?: WorkbookTableQueryStore;
}

function applyCells(context: CommandContext, payload: QueryCellLoadPayload): void {
  if (payload.queryDefinition === null) context.workbook.removeQueryDefinition(payload.queryId);
  else context.workbook.setQueryDefinition(payload.queryDefinition);
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
  if (payload.queryDefinition === null) context.workbook.removeQueryDefinition(payload.queryId);
  else context.workbook.setQueryDefinition(payload.queryDefinition);
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
  if (payload.queryDefinition === null) context.workbook.removeQueryDefinition(payload.queryId);
  else context.workbook.setQueryDefinition(payload.queryDefinition);
  context.workbook.tables.set(payload.table.id, structuredClone(payload.table));
  store.set(payload.tableId, { result: structuredClone(payload.result), sourceRevision: payload.sourceRevision });
}

function restoreWorkbookTable(context: CommandContext, payload: QueryWorkbookTableRestorePayload, store: WorkbookTableQueryStore): void {
  if (payload.queryDefinition === null) context.workbook.removeQueryDefinition(payload.queryId);
  else context.workbook.setQueryDefinition(payload.queryDefinition);
  if (payload.previousTable) context.workbook.tables.set(payload.tableId, structuredClone(payload.previousTable));
  else context.workbook.tables.delete(payload.tableId);
  if (payload.previousRecord) store.set(payload.tableId, structuredClone(payload.previousRecord));
  else store.delete(payload.tableId);
}

function registerQueryMutations(registry: CommandRegistry, store: WorkbookTableQueryStore): void {
  registry.registerMutation<QueryDefinitionReplaceMutationParams>({
    id: 'query.definition.replace',
    handler: (item, context) => {
      if (!isQueryDefinitionReplacePayload(item.params)) throw new Error('Invalid query.definition.replace mutation payload');
      if (item.params.definition === null) context.workbook.removeQueryDefinition(item.params.queryId);
      else context.workbook.setQueryDefinition(item.params.definition);
    },
    metadata: {
      schema: { name: 'QueryDefinitionReplaceMutationParams', validate: isQueryDefinitionReplacePayload },
      permission: { capability: 'query.definition.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['query.definition.replace'],
    },
  });
  for (const id of ['query.load.range', 'query.load.sheet-table', 'query.load.pivot-source'] as const) {
    registry.registerMutation<QueryCellLoadPayload | QueryCellRestorePayload>({
      id,
      handler: (item, context) => {
        const payload = item.params as QueryCellLoadPayload | QueryCellRestorePayload;
        if (!isQueryLoadPayload(payload)) throw new Error(`Invalid ${id} mutation payload`);
        if ('previousCells' in payload) restoreCells(context, payload);
        else applyCells(context, payload);
      },
      metadata: {
        schema: { name: `${id}Params`, validate: isQueryLoadPayload },
        permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
        affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
        inverseIds: [id],
      },
    });
  }
  registry.registerMutation<QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload>({
    id: 'query.load.workbook-table',
    handler: (item, context) => {
      const payload = item.params as QueryWorkbookTableLoadPayload | QueryWorkbookTableRestorePayload;
      if (!isQueryLoadPayload(payload)) throw new Error('Invalid query.load.workbook-table mutation payload');
      if ('previousTable' in payload || 'previousRecord' in payload) restoreWorkbookTable(context, payload, store);
      else applyWorkbookTable(context, payload, store);
    },
    metadata: {
      schema: { name: 'query.load.workbook-tableParams', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
      inverseIds: ['query.load.workbook-table'],
    },
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

  registry.registerCommand<QueryDefinitionReplaceParams>({
    id: 'query.definition.replace',
    execute: (params, context) => {
      const next = serializeQueryDefinition(params.definition);
      const previous = context.workbook.getQueryDefinition(next.id) ?? null;
      context.applyMutation({
        id: 'query.definition.replace',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.activeSheetId,
        params: { queryId: next.id, definition: next },
        affectedRanges: [],
        inverse: [{
          id: 'query.definition.replace',
          unitId: context.workbook.unitId,
          sheetId: context.workbook.activeSheetId,
          params: { queryId: next.id, definition: previous },
          affectedRanges: [],
        }],
        apply: () => context.workbook.setQueryDefinition(next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<QueryRefreshParams>({
    id: 'query.refresh',
    execute: (params, context) => executeLoad(registry, { query: params.query, target: params.target, result: params.result }, context, store),
  });
}
