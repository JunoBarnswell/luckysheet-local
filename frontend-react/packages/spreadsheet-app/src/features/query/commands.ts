import type { DataSourceManifest, RangeRef, TableScalar } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  buildQueryLoadPlan,
  type QueryLoadBinding,
  type QueryLoadCommandPayload,
  type QueryLoadMutationPayload,
} from './runtime';
import { normalizeDataSourceManifest } from '@react-sheets/core-model';
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
  if (value.source !== null && !isRecord(value.source)) return false;
  if (value.binding !== null && !isRecord(value.binding)) return false;
  if (value.source !== null && value.binding !== null && value.sourceId !== value.source.id) return false;
  return true;
}

function sameRange(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startColumn === right.startColumn
    && left.endColumn === right.endColumn;
}

function queryMutationRanges(params: QueryLoadMutationPayload): readonly RangeRef[] {
  return params.binding?.kind === 'sheet-region' ? [params.binding.region.range] : [];
}

function clearRegionCells(context: CommandContext, range: RangeRef): void {
  const sheet = context.workbook.getSheet(range.sheetId);
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = sheet.cells.get(row, column) as (Record<string, unknown> & { __cellPatch?: unknown }) | undefined;
      // CellPatch is the sole local-edit overlay; it must survive a source refresh.
      if (cell?.__cellPatch) continue;
      sheet.cells.delete(row, column);
    }
  }
}

function removeCurrentBinding(context: CommandContext, sourceId: string): void {
  for (const sheet of context.workbook.getSheets()) {
    for (let index = sheet.dataRegions.length - 1; index >= 0; index -= 1) {
      const region = sheet.dataRegions[index]!;
      if (region.sourceId !== sourceId) continue;
      clearRegionCells(context, region.range);
      sheet.dataRegions.splice(index, 1);
    }
  }
  for (const table of context.workbook.dataModel.tables.values()) {
    if (table.sourceId === sourceId) delete table.sourceId;
  }
  context.workbook.dataModel.sources.delete(sourceId);
}

function applySource(context: CommandContext, source: DataSourceManifest | null): void {
  if (!source) return;
  if (context.workbook.dataModel.sources.has(source.id)) context.workbook.updateDataSource(source);
  else context.workbook.addDataSource(source);
}

function writeHeader(context: CommandContext, header: readonly TableScalar[], region: QueryLoadBinding & { kind: 'sheet-region' }): void {
  const sheet = context.workbook.getSheet(region.region.range.sheetId);
  for (let offset = 0; offset < header.length; offset += 1) {
    const value = header[offset];
    if (value !== undefined && value !== null) sheet.cells.set(region.region.headerRow, region.region.range.startColumn + offset, { value });
  }
}

function applyBinding(context: CommandContext, binding: QueryLoadBinding | null, extent?: { sheetId: string; rowCount: number; columnCount: number }): void {
  if (extent) {
    const sheet = context.workbook.getSheet(extent.sheetId);
    if (extent.rowCount < 1 || extent.columnCount < 1) throw new Error('Query load extent must be positive');
    sheet.rowCount = extent.rowCount;
    sheet.columnCount = extent.columnCount;
  }
  if (!binding) return;
  if (binding.kind === 'workbook-table') {
    if (!context.workbook.dataModel.tables.has(binding.tableId)) throw new Error(`Unknown workbook table: ${binding.tableId}`);
    context.workbook.dataModel.tables.set(binding.tableId, structuredClone(binding.table));
    return;
  }
  const sheet = context.workbook.getSheet(binding.region.range.sheetId);
  if (extent) {
    if (extent.sheetId !== sheet.id || extent.rowCount < binding.region.range.endRow + 1 || extent.columnCount < binding.region.range.endColumn + 1) throw new Error(`Query load extent does not contain region ${binding.region.id}`);
  } else {
    sheet.ensureRangeExtent(binding.region.range.startRow, binding.region.range.endRow, binding.region.range.startColumn, binding.region.range.endColumn);
  }
  clearRegionCells(context, binding.region.range);
  sheet.dataRegions.push(structuredClone(binding.region));
  writeHeader(context, binding.header, binding);
}

function applyQueryLoad(context: CommandContext, payload: QueryLoadMutationPayload): void {
  if (!isQueryLoadPayload(payload)) throw new Error('Invalid block-backed query load mutation payload');
  const normalizedSource = payload.source ? normalizeDataSourceManifest(structuredClone(payload.source)) : null;
  if (payload.binding && !normalizedSource) throw new Error(`Query load binding has no source manifest: ${payload.sourceId}`);
  if (normalizedSource && payload.binding?.kind === 'sheet-region') {
    const regionWidth = payload.binding.region.range.endColumn - payload.binding.region.range.startColumn + 1;
    if (payload.binding.region.sourceId !== payload.sourceId
      || payload.binding.header.length !== normalizedSource.fields.length
      || regionWidth !== normalizedSource.fields.length
      || payload.binding.region.revision !== normalizedSource.revision) {
      throw new Error(`Query load source and sheet-region binding revision/shape do not match: ${payload.sourceId}`);
    }
    if (normalizedSource.sourceSheetId !== undefined && normalizedSource.sourceSheetId !== payload.binding.region.range.sheetId) {
      throw new Error(`Query load source sheet does not match sheet-region binding: ${payload.sourceId}`);
    }
    if (normalizedSource.sourceRange !== undefined && !sameRange(normalizedSource.sourceRange, payload.binding.region.range)) {
      throw new Error(`Query load source range does not match sheet-region binding: ${payload.sourceId}`);
    }
  }
  if (payload.binding?.kind === 'workbook-table') {
    const currentTable = context.workbook.dataModel.tables.get(payload.binding.tableId);
    if (!currentTable || payload.binding.table.id !== payload.binding.tableId) throw new Error(`Unknown workbook table: ${payload.binding.tableId}`);
    if (normalizedSource && payload.binding.table.sourceId !== payload.sourceId) throw new Error(`Query load table source does not match sourceId: ${payload.sourceId}`);
  }
  if (payload.binding?.kind === 'sheet-region') {
    const range = payload.binding.region.range;
    if (![range.startRow, range.endRow, range.startColumn, range.endColumn, payload.binding.region.headerRow].every((value) => Number.isSafeInteger(value) && value >= 0)
      || range.endRow < range.startRow || range.endColumn < range.startColumn
      || payload.binding.region.headerRow !== range.startRow) {
      throw new Error(`Query load sheet-region range is invalid: ${payload.binding.region.id}`);
    }
    context.workbook.getSheet(range.sheetId);
  }
  if (payload.extent) {
    const extentSheet = context.workbook.getSheet(payload.extent.sheetId);
    if (!Number.isSafeInteger(payload.extent.rowCount) || payload.extent.rowCount < 1
      || !Number.isSafeInteger(payload.extent.columnCount) || payload.extent.columnCount < 1) {
      throw new Error(`Query load extent is invalid: ${payload.extent.sheetId}`);
    }
    if (payload.binding?.kind === 'sheet-region') {
      const range = payload.binding.region.range;
      if (extentSheet.id !== range.sheetId || payload.extent.rowCount < range.endRow + 1 || payload.extent.columnCount < range.endColumn + 1) {
        throw new Error(`Query load extent does not contain region ${payload.binding.region.id}`);
      }
    }
  }
  if (payload.queryDefinition === null) context.workbook.removeQueryDefinition(payload.queryId);
  else context.workbook.setQueryDefinition(payload.queryDefinition);
  removeCurrentBinding(context, payload.sourceId);
  applySource(context, normalizedSource);
  applyBinding(context, payload.binding, payload.extent);
  if (payload.pivotSource && payload.target.pivotId) {
    const owner = context.workbook.getSheets().find((sheet) => sheet.pivots.some((pivot) => pivot.id === payload.target.pivotId));
    const pivot = owner?.pivots.find((entry) => entry.id === payload.target.pivotId);
    if (!pivot) throw new Error(`Unknown pivot: ${payload.target.pivotId}`);
    pivot.source = structuredClone(payload.pivotSource);
  }
}

function isQueryDefinitionReplacePayload(value: unknown): value is QueryDefinitionReplaceMutationParams {
  if (!isRecord(value)) return false;
  return typeof value.queryId === 'string' && value.queryId.trim().length > 0 && isQueryDefinition(value.definition);
}

function registerQueryMutations(registry: CommandRegistry): void {
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
  const queryLoadHandler = (item: { params: QueryLoadMutationPayload }, context: CommandContext): void => applyQueryLoad(context, item.params);
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.range',
    handler: queryLoadHandler,
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['query.load.range'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.sheet-table',
    handler: queryLoadHandler,
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['query.load.sheet-table'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.pivot-source',
    handler: queryLoadHandler,
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['query.load.pivot-source'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<QueryLoadMutationPayload>({
    id: 'query.load.workbook-table',
    handler: queryLoadHandler,
    metadata: {
      schema: { name: 'QueryLoadDataSource', validate: isQueryLoadPayload },
      permission: { capability: 'query.load.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: queryMutationRanges, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['query.load.workbook-table'], minCount: 1, maxCount: 1 },
    },
  });
}

function applyQueryLoadMutation(registry: CommandRegistry, plan: ReturnType<typeof buildQueryLoadPlan>, context: CommandContext): void {
  const sheetId = plan.payload.binding.kind === 'sheet-region'
    ? plan.payload.binding.region.range.sheetId
    : plan.payload.target.sheetId ?? context.workbook.primarySheetId;
  switch (plan.mutationId) {
    case 'query.load.range':
      context.applyMutation({
        id: 'query.load.range', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges,
        inverse: [{ id: 'query.load.range', unitId: context.workbook.unitId, sheetId, params: plan.inverse, affectedRanges: plan.affectedRanges }],
        apply: () => registry.getMutation<QueryLoadMutationPayload>('query.load.range')({ id: 'query.load.range', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges }, context),
      });
      return;
    case 'query.load.sheet-table':
      context.applyMutation({
        id: 'query.load.sheet-table', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges,
        inverse: [{ id: 'query.load.sheet-table', unitId: context.workbook.unitId, sheetId, params: plan.inverse, affectedRanges: plan.affectedRanges }],
        apply: () => registry.getMutation<QueryLoadMutationPayload>('query.load.sheet-table')({ id: 'query.load.sheet-table', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges }, context),
      });
      return;
    case 'query.load.pivot-source':
      context.applyMutation({
        id: 'query.load.pivot-source', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges,
        inverse: [{ id: 'query.load.pivot-source', unitId: context.workbook.unitId, sheetId, params: plan.inverse, affectedRanges: plan.affectedRanges }],
        apply: () => registry.getMutation<QueryLoadMutationPayload>('query.load.pivot-source')({ id: 'query.load.pivot-source', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges }, context),
      });
      return;
    case 'query.load.workbook-table':
      context.applyMutation({
        id: 'query.load.workbook-table', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges,
        inverse: [{ id: 'query.load.workbook-table', unitId: context.workbook.unitId, sheetId, params: plan.inverse, affectedRanges: plan.affectedRanges }],
        apply: () => registry.getMutation<QueryLoadMutationPayload>('query.load.workbook-table')({ id: 'query.load.workbook-table', unitId: context.workbook.unitId, sheetId, params: plan.payload, affectedRanges: plan.affectedRanges }, context),
      });
      return;
  }
}

function executeLoad(registry: CommandRegistry, params: QueryLoadCommandPayload, context: CommandContext): CommandResult {
  const plan = buildQueryLoadPlan(context.workbook, params);
  applyQueryLoadMutation(registry, plan, context);
  return { operationId: context.operationId, mutationCount: 1, affectedRanges: plan.affectedRanges };
}

export function registerQueryCommands(registry: CommandRegistry): void {
  registerQueryMutations(registry);
  registry.registerCommand<QueryLoadParams>({ id: 'query.load', execute: (params, context) => executeLoad(registry, params, context) });
  registry.registerCommand<QueryDefinitionReplaceParams>({
    id: 'query.definition.replace',
    execute: (params, context) => {
      const next = serializeQueryDefinition(params.definition);
      const previous = context.workbook.getQueryDefinition(next.id) ?? null;
      context.applyMutation({
        id: 'query.definition.replace', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId,
        params: { queryId: next.id, definition: next }, affectedRanges: [],
        inverse: [{ id: 'query.definition.replace', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { queryId: next.id, definition: previous }, affectedRanges: [] }],
        apply: () => context.workbook.setQueryDefinition(next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
  registry.registerCommand<QueryRefreshParams>({ id: 'query.refresh', execute: (params, context) => executeLoad(registry, params, context) });
}
