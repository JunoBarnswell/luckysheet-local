import type { RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  buildQueryLoadPlan,
  type QueryLoadCommandPayload,
  type QueryLoadMutationPayload,
} from './runtime';
import { serializeQueryDefinition, type QueryDefinitionPersistence } from './index';
import { type LoadTarget, type QueryDefinition } from './query-steps';

export interface QueryLoadParams extends QueryLoadCommandPayload {}

export interface QueryRefreshParams extends QueryLoadCommandPayload {
  queryId: string;
}

export interface QueryDefinitionReplaceParams {
  definition: QueryDefinition;
}

interface QueryDefinitionReplaceMutationParams {
  queryId: string;
  definition: QueryDefinitionPersistence | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isQueryDefinition(value: unknown): value is QueryDefinitionPersistence {
  return value === null || (isRecord(value) && value.schema === 'QueryDefinition');
}

function isQueryLoadPayload(value: unknown): value is QueryLoadMutationPayload {
  if (!isRecord(value) || value.kind !== 'data-source-load') return false;
  if (typeof value.queryId !== 'string' || !value.queryId.trim() || typeof value.sourceId !== 'string' || !value.sourceId.trim()) return false;
  if (!isQueryDefinition(value.queryDefinition) || !isRecord(value.target) || !['range', 'sheet-table', 'pivot-source', 'workbook-table'].includes(String(value.target.kind))) return false;
  if (!isRecord(value.source) || !isRecord(value.binding)) return false;
  if (value.sourceId !== value.source.id) return false;
  const proofFields = [value.executionToken, value.resultHash, value.sourceRevision];
  if (proofFields.some((field) => field !== undefined) && (typeof value.executionToken !== 'string' || !value.executionToken.trim() || typeof value.resultHash !== 'string' || !value.resultHash.trim() || !Number.isSafeInteger(value.sourceRevision) || Number(value.sourceRevision) < 0)) return false;
  return true;
}

function queryMutationRanges(params: QueryLoadMutationPayload): readonly RangeRef[] {
  return params.binding.kind === 'sheet-region' ? [params.binding.region.range] : [];
}

function isQueryDefinitionReplacePayload(value: unknown): value is QueryDefinitionReplaceMutationParams {
  if (!isRecord(value)) return false;
  return typeof value.queryId === 'string' && value.queryId.trim().length > 0 && isQueryDefinition(value.definition);
}

function registerQueryMutations(registry: CommandRegistry): void {
  registry.registerMutation<QueryDefinitionReplaceMutationParams>({
    id: 'query.definition.replace',
    metadata: {
      schema: { name: 'QueryDefinitionReplaceMutationParams', validate: isQueryDefinitionReplacePayload },
      permission: { capability: 'query.definition.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.range',
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.sheet-table',
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.pivot-source',
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.workbook-table',
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
    },
  });
}

function executeLoad(params: QueryLoadCommandPayload, context: CommandContext): CommandResult {
  const plan = buildQueryLoadPlan(context.workbook, params);
  const sheetId = plan.payload.binding.kind === 'sheet-region'
    ? plan.payload.binding.region.range.sheetId
    : plan.payload.target.sheetId ?? context.workbook.primarySheetId;
  const mutation = {
    unitId: context.workbook.unitId,
    sheetId,
    params: plan.payload,
    affectedRanges: plan.affectedRanges,
  };
  switch (plan.mutationId) {
    case 'query.load.range': context.applyMutation({ id: 'query.load.range', ...mutation }); break;
    case 'query.load.sheet-table': context.applyMutation({ id: 'query.load.sheet-table', ...mutation }); break;
    case 'query.load.pivot-source': context.applyMutation({ id: 'query.load.pivot-source', ...mutation }); break;
    case 'query.load.workbook-table': context.applyMutation({ id: 'query.load.workbook-table', ...mutation }); break;
  }
  return { operationId: context.operationId, mutationCount: 1, affectedRanges: plan.affectedRanges };
}

export function registerQueryCommands(registry: CommandRegistry): void {
  registerQueryMutations(registry);
  registry.registerCommand<QueryLoadParams>({ id: 'query.load', execute: (params, context) => executeLoad(params, context) });
  registry.registerCommand<QueryDefinitionReplaceParams>({
    id: 'query.definition.replace',
    execute: (params, context) => {
      const next = serializeQueryDefinition(params.definition);
      context.applyMutation({
        id: 'query.definition.replace', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId,
        params: { queryId: next.id, definition: next }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
  registry.registerCommand<QueryRefreshParams>({ id: 'query.refresh', execute: (params, context) => executeLoad(params, context) });
}
