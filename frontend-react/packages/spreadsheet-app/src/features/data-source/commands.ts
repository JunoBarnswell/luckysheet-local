import type { CommandContext, CommandRegistry, CommandResult, CommandRuntime } from '@react-sheets/command-runtime';
import {
  DEFAULT_DATA_BLOCK_ROW_COUNT,
  normalizeDataSourceManifest,
  type DataSourceManifest,
  type RangeRef,
  type SheetDataRegion,
} from '@react-sheets/core-model';
import {
  validateDataSourceManifest,
  validateDataSourceMutationParams,
  type DataSourceMutationId,
} from '@react-sheets/protocol';

export interface DataSourceAddCommandParams {
  sheetId: string;
  source: DataSourceManifest;
}

export interface DataSourceUpdateCommandParams {
  sheetId: string;
  source: DataSourceManifest;
}

export interface DataSourceRemoveCommandParams {
  sheetId: string;
  sourceId: string;
}

export interface DataRegionAddCommandParams {
  sheetId: string;
  region: SheetDataRegion;
}

export interface DataRegionRemoveCommandParams {
  sheetId: string;
  regionId: string;
}

export const DATA_SOURCE_COMMAND_IDS = [
  'dataSource.add',
  'dataSource.update',
  'dataSource.remove',
  'dataRegion.add',
  'dataRegion.remove',
] as const;

export const DATA_SOURCE_MUTATION_IDS = DATA_SOURCE_COMMAND_IDS;

type DataSourceCommandParams =
  | DataSourceAddCommandParams
  | DataSourceUpdateCommandParams
  | DataSourceRemoveCommandParams
  | DataRegionAddCommandParams
  | DataRegionRemoveCommandParams;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function requireSheetId(value: unknown, label: string): string {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || value.sheetId.trim().length === 0) {
    throw new Error(`${label} requires a sheetId`);
  }
  return value.sheetId;
}

function validateRangeMetadata(value: unknown, label: string): asserts value is RangeRef {
  if (!isRecord(value)
    || Object.keys(value).some((key) => !['sheetId', 'startRow', 'endRow', 'startColumn', 'endColumn'].includes(key))
    || typeof value.sheetId !== 'string'
    || !Number.isSafeInteger(value.startRow) || Number(value.startRow) < 0
    || !Number.isSafeInteger(value.endRow) || Number(value.endRow) < Number(value.startRow)
    || !Number.isSafeInteger(value.startColumn) || Number(value.startColumn) < 0
    || !Number.isSafeInteger(value.endColumn) || Number(value.endColumn) < Number(value.startColumn)) {
    throw new Error(`${label} is invalid`);
  }
}

/**
 * Validate a manifest against both the wire contract and the workbook-local
 * invariants. This function never inspects or loads block bytes.
 */
export function normalizeDataSourceForCommand(
  workbook: CommandContext['workbook'],
  sheetId: string,
  value: unknown,
): DataSourceManifest {
  validateDataSourceManifest(value);
  const source = normalizeDataSourceManifest(structuredClone(value));
  if (source.blockRowCount !== DEFAULT_DATA_BLOCK_ROW_COUNT) {
    throw new Error(`Data source blockRowCount must be ${DEFAULT_DATA_BLOCK_ROW_COUNT}`);
  }
  if (source.sourceSheetId !== undefined) {
    workbook.getSheet(source.sourceSheetId);
    if (source.sourceSheetId !== sheetId) throw new Error('Data source sheetId does not match the command sheetId');
  }
  if (source.sourceRange !== undefined) {
    validateRangeMetadata(source.sourceRange, 'Data source sourceRange');
    workbook.getSheet(source.sourceRange.sheetId);
    if (source.sourceSheetId !== source.sourceRange.sheetId) {
      throw new Error('Data source sourceRange must target sourceSheetId');
    }
  }
  if ((source.kind === 'worksheet-range' || source.kind === 'sheet-table')
    && (!source.sourceSheetId || !source.sourceRange)) {
    throw new Error(`${source.kind} data sources require sourceSheetId and sourceRange`);
  }
  const blockIds = new Set<string>();
  for (const block of source.blocks) {
    if (blockIds.has(block.id)) throw new Error(`Duplicate data block: ${block.id}`);
    blockIds.add(block.id);
    if (block.startRow + block.rowCount > source.rowCount) {
      throw new Error(`Data block exceeds source rowCount: ${block.id}`);
    }
    if (!/^[A-Fa-f0-9]{64}$/.test(block.checksum)) throw new Error(`Data block checksum is invalid: ${block.id}`);
    if (block.byteLength <= 0) throw new Error(`Data block byteLength must be positive: ${block.id}`);
  }
  return source;
}

function normalizeRegionForCommand(
  workbook: CommandContext['workbook'],
  sheetId: string,
  value: unknown,
): SheetDataRegion {
  validateDataSourceMutationParams('dataRegion.add', { region: value });
  if (!isRecord(value)) throw new Error('Sheet data region is required');
  const region = structuredClone(value) as unknown as SheetDataRegion;
  if (region.range.sheetId !== sheetId) throw new Error('Sheet data region range must target the command sheet');
  validateRangeMetadata(region.range, 'Sheet data region range');
  if (region.headerRow < region.range.startRow || region.headerRow > region.range.endRow) {
    throw new Error('Sheet data region headerRow must be inside its range');
  }
  workbook.getSheet(sheetId);
  workbook.getDataSource(region.sourceId);
  return region;
}

function sourceRanges(workbook: CommandContext['workbook'], source: DataSourceManifest, fallbackSheetId: string): RangeRef[] {
  if (source.sourceRange) return [structuredClone(source.sourceRange)];
  if (source.sourceSheetId) {
    const sheet = workbook.getSheet(source.sourceSheetId);
    return [{
      sheetId: source.sourceSheetId,
      startRow: 0,
      endRow: Math.max(0, sheet.rowCount - 1),
      startColumn: 0,
      endColumn: Math.max(0, sheet.columnCount - 1),
    }];
  }
  // A chunked source without a worksheet owner has no cell range. The
  // mutation still carries its owning operation sheet for the protocol.
  void fallbackSheetId;
  return [];
}

function regionRanges(region: SheetDataRegion): RangeRef[] {
  return [structuredClone(region.range)];
}

function findRegion(workbook: CommandContext['workbook'], sheetId: string, regionId: string): SheetDataRegion {
  const region = workbook.getSheet(sheetId).dataRegions.find((entry) => entry.id === regionId);
  if (!region) throw new Error(`Unknown sheet data region: ${regionId}`);
  return structuredClone(region);
}

function mutationParams(id: DataSourceMutationId, params: unknown): void {
  validateDataSourceMutationParams(id, params);
}

function applyRegionAdd(workbook: CommandContext['workbook'], region: SheetDataRegion): void {
  const sheet = workbook.getSheet(region.range.sheetId);
  if (sheet.dataRegions.some((entry) => entry.id === region.id)) throw new Error(`Sheet data region already exists: ${region.id}`);
  workbook.getDataSource(region.sourceId);
  sheet.dataRegions.push(structuredClone(region));
}

function applyRegionRemove(workbook: CommandContext['workbook'], sheetId: string, regionId: string): void {
  const sheet = workbook.getSheet(sheetId);
  const index = sheet.dataRegions.findIndex((entry) => entry.id === regionId);
  if (index < 0) throw new Error(`Unknown sheet data region: ${regionId}`);
  sheet.dataRegions.splice(index, 1);
}

function applyDataSourceMutation(workbook: CommandContext['workbook'], id: DataSourceMutationId, params: unknown, sheetId: string): void {
  mutationParams(id, params);
  switch (id) {
    case 'dataSource.add':
      if ((params as { source: DataSourceManifest }).source.sourceSheetId !== undefined
        && (params as { source: DataSourceManifest }).source.sourceSheetId !== sheetId) {
        throw new Error('Data source sheetId does not match the mutation sheetId');
      }
      workbook.addDataSource((params as { source: DataSourceManifest }).source);
      return;
    case 'dataSource.update':
      if ((params as { source: DataSourceManifest }).source.sourceSheetId !== undefined
        && (params as { source: DataSourceManifest }).source.sourceSheetId !== sheetId) {
        throw new Error('Data source sheetId does not match the mutation sheetId');
      }
      workbook.updateDataSource((params as { source: DataSourceManifest }).source);
      return;
    case 'dataSource.remove':
      workbook.removeDataSource((params as { sourceId: string }).sourceId);
      return;
    case 'dataRegion.add':
      if ((params as { region: SheetDataRegion }).region.range.sheetId !== sheetId) throw new Error('Sheet data region targets another sheet');
      applyRegionAdd(workbook, (params as { region: SheetDataRegion }).region);
      return;
    case 'dataRegion.remove':
      applyRegionRemove(workbook, sheetId, (params as { regionId: string }).regionId);
      return;
  }
}

function applyDataSourceAdd(
  context: CommandContext,
  sheetId: string,
  source: DataSourceManifest,
  affectedRanges: RangeRef[],
): void {
  context.applyMutation({
    id: 'dataSource.add',
    unitId: context.workbook.unitId,
    sheetId,
    params: { source },
    affectedRanges,
    inverse: [{ id: 'dataSource.remove', unitId: context.workbook.unitId, sheetId, params: { sourceId: source.id }, affectedRanges }],
    apply: () => applyDataSourceMutation(context.workbook, 'dataSource.add', { source }, sheetId),
  });
}

function applyDataSourceUpdate(
  context: CommandContext,
  sheetId: string,
  source: DataSourceManifest,
  previous: DataSourceManifest,
  affectedRanges: RangeRef[],
): void {
  context.applyMutation({
    id: 'dataSource.update',
    unitId: context.workbook.unitId,
    sheetId,
    params: { source },
    affectedRanges,
    inverse: [{ id: 'dataSource.update', unitId: context.workbook.unitId, sheetId, params: { source: previous }, affectedRanges }],
    apply: () => applyDataSourceMutation(context.workbook, 'dataSource.update', { source }, sheetId),
  });
}

function applyDataSourceRemove(
  context: CommandContext,
  sheetId: string,
  source: DataSourceManifest,
  affectedRanges: RangeRef[],
): void {
  context.applyMutation({
    id: 'dataSource.remove',
    unitId: context.workbook.unitId,
    sheetId,
    params: { sourceId: source.id },
    affectedRanges,
    inverse: [{ id: 'dataSource.add', unitId: context.workbook.unitId, sheetId, params: { source }, affectedRanges }],
    apply: () => applyDataSourceMutation(context.workbook, 'dataSource.remove', { sourceId: source.id }, sheetId),
  });
}

function applyDataRegionAdd(
  context: CommandContext,
  sheetId: string,
  region: SheetDataRegion,
  affectedRanges: RangeRef[],
): void {
  context.applyMutation({
    id: 'dataRegion.add',
    unitId: context.workbook.unitId,
    sheetId,
    params: { region },
    affectedRanges,
    inverse: [{ id: 'dataRegion.remove', unitId: context.workbook.unitId, sheetId, params: { regionId: region.id }, affectedRanges }],
    apply: () => applyDataSourceMutation(context.workbook, 'dataRegion.add', { region }, sheetId),
  });
}

function applyDataRegionRemove(
  context: CommandContext,
  sheetId: string,
  region: SheetDataRegion,
  affectedRanges: RangeRef[],
): void {
  context.applyMutation({
    id: 'dataRegion.remove',
    unitId: context.workbook.unitId,
    sheetId,
    params: { regionId: region.id },
    affectedRanges,
    inverse: [{ id: 'dataRegion.add', unitId: context.workbook.unitId, sheetId, params: { region }, affectedRanges }],
    apply: () => applyDataSourceMutation(context.workbook, 'dataRegion.remove', { regionId: region.id }, sheetId),
  });
}

function sourceCommandParams(value: unknown, label: string): { sheetId: string; source: DataSourceManifest } {
  const sheetId = requireSheetId(value, label);
  if (!isRecord(value) || !('source' in value)) throw new Error(`${label} requires a source`);
  return { sheetId, source: value.source as DataSourceManifest };
}

function removeSourceCommandParams(value: unknown): DataSourceRemoveCommandParams {
  const sheetId = requireSheetId(value, 'dataSource.remove');
  if (!isRecord(value) || typeof value.sourceId !== 'string' || value.sourceId.trim().length === 0) {
    throw new Error('dataSource.remove requires sourceId');
  }
  return { sheetId, sourceId: value.sourceId };
}

function addRegionCommandParams(value: unknown): DataRegionAddCommandParams {
  const sheetId = requireSheetId(value, 'dataRegion.add');
  if (!isRecord(value) || !('region' in value)) throw new Error('dataRegion.add requires a region');
  return { sheetId, region: value.region as SheetDataRegion };
}

function removeRegionCommandParams(value: unknown): DataRegionRemoveCommandParams {
  const sheetId = requireSheetId(value, 'dataRegion.remove');
  if (!isRecord(value) || typeof value.regionId !== 'string' || value.regionId.trim().length === 0) {
    throw new Error('dataRegion.remove requires regionId');
  }
  return { sheetId, regionId: value.regionId };
}

function result(context: CommandContext, affectedRanges: RangeRef[]): CommandResult {
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function sourceSchema(value: unknown): boolean {
  try {
    validateDataSourceMutationParams('dataSource.add', value);
    return true;
  } catch {
    return false;
  }
}

function sourceUpdateSchema(value: unknown): boolean {
  return sourceSchema(value);
}

function sourceRemoveSchema(value: unknown): boolean {
  try {
    validateDataSourceMutationParams('dataSource.remove', value);
    return true;
  } catch {
    return false;
  }
}

function regionAddSchema(value: unknown): boolean {
  try {
    validateDataSourceMutationParams('dataRegion.add', value);
    return true;
  } catch {
    return false;
  }
}

function regionRemoveSchema(value: unknown): boolean {
  try {
    validateDataSourceMutationParams('dataRegion.remove', value);
    return true;
  } catch {
    return false;
  }
}

function registerMutationContracts(registry: CommandRegistry): void {
  registry.registerMutation<{ source: DataSourceManifest }>({
    id: 'dataSource.add',
    handler: (item, context) => applyDataSourceMutation(context.workbook, 'dataSource.add', item.params, item.sheetId),
    metadata: {
      schema: { name: 'DataSourceAddMutation', validate: sourceSchema },
      permission: { capability: 'data-source.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.source.sourceRange ? [structuredClone(params.source.sourceRange)] : [], mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['dataSource.remove'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<{ source: DataSourceManifest }>({
    id: 'dataSource.update',
    handler: (item, context) => applyDataSourceMutation(context.workbook, 'dataSource.update', item.params, item.sheetId),
    metadata: {
      schema: { name: 'DataSourceUpdateMutation', validate: sourceUpdateSchema },
      permission: { capability: 'data-source.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.source.sourceRange ? [structuredClone(params.source.sourceRange)] : [], mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['dataSource.update'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<{ sourceId: string }>({
    id: 'dataSource.remove',
    handler: (item, context) => applyDataSourceMutation(context.workbook, 'dataSource.remove', item.params, item.sheetId),
    metadata: {
      schema: { name: 'DataSourceRemoveMutation', validate: sourceRemoveSchema },
      permission: { capability: 'data-source.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['dataSource.add'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<{ region: SheetDataRegion }>({
    id: 'dataRegion.add',
    handler: (item, context) => applyDataSourceMutation(context.workbook, 'dataRegion.add', item.params, item.sheetId),
    metadata: {
      schema: { name: 'DataRegionAddMutation', validate: regionAddSchema },
      permission: { capability: 'data-source.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.region.range)], mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['dataRegion.remove'], minCount: 1, maxCount: 1 },
    },
  });
  registry.registerMutation<{ regionId: string }>({
    id: 'dataRegion.remove',
    handler: (item, context) => applyDataSourceMutation(context.workbook, 'dataRegion.remove', item.params, item.sheetId),
    metadata: {
      schema: { name: 'DataRegionRemoveMutation', validate: regionRemoveSchema },
      permission: { capability: 'data-source.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['dataRegion.add'], minCount: 1, maxCount: 1 },
    },
  });
}

export function registerDataSourceCommands(runtime: CommandRuntime): string[] {
  registerMutationContracts(runtime.registry);

  runtime.registry.registerCommand<DataSourceAddCommandParams>({
    id: 'dataSource.add',
    execute(input, context): CommandResult {
      const params = sourceCommandParams(input, 'dataSource.add');
      const source = normalizeDataSourceForCommand(context.workbook, params.sheetId, params.source);
      if (context.workbook.dataModel.sources.has(source.id)) throw new Error(`Data source already exists: ${source.id}`);
      const affectedRanges = sourceRanges(context.workbook, source, params.sheetId);
      applyDataSourceAdd(context, params.sheetId, source, affectedRanges);
      return result(context, affectedRanges);
    },
  });

  runtime.registry.registerCommand<DataSourceUpdateCommandParams>({
    id: 'dataSource.update',
    execute(input, context): CommandResult {
      const params = sourceCommandParams(input, 'dataSource.update');
      const previous = context.workbook.getDataSource(params.source.id);
      const source = normalizeDataSourceForCommand(context.workbook, params.sheetId, params.source);
      const affectedRanges = sourceRanges(context.workbook, source, params.sheetId);
      applyDataSourceUpdate(context, params.sheetId, source, previous, affectedRanges);
      return result(context, affectedRanges);
    },
  });

  runtime.registry.registerCommand<DataSourceRemoveCommandParams>({
    id: 'dataSource.remove',
    execute(input, context): CommandResult {
      const params = removeSourceCommandParams(input);
      const source = context.workbook.getDataSource(params.sourceId);
      const affectedRanges = sourceRanges(context.workbook, source, params.sheetId);
      applyDataSourceRemove(context, params.sheetId, source, affectedRanges);
      return result(context, affectedRanges);
    },
  });

  runtime.registry.registerCommand<DataRegionAddCommandParams>({
    id: 'dataRegion.add',
    execute(input, context): CommandResult {
      const params = addRegionCommandParams(input);
      const region = normalizeRegionForCommand(context.workbook, params.sheetId, params.region);
      if (context.workbook.getSheet(params.sheetId).dataRegions.some((entry) => entry.id === region.id)) {
        throw new Error(`Sheet data region already exists: ${region.id}`);
      }
      const affectedRanges = regionRanges(region);
      applyDataRegionAdd(context, params.sheetId, region, affectedRanges);
      return result(context, affectedRanges);
    },
  });

  runtime.registry.registerCommand<DataRegionRemoveCommandParams>({
    id: 'dataRegion.remove',
    execute(input, context): CommandResult {
      const params = removeRegionCommandParams(input);
      const region = findRegion(context.workbook, params.sheetId, params.regionId);
      const affectedRanges = regionRanges(region);
      applyDataRegionRemove(context, params.sheetId, region, affectedRanges);
      return result(context, affectedRanges);
    },
  });

  return [...DATA_SOURCE_COMMAND_IDS];
}

export function registerDataSourceFeature(runtime: CommandRuntime): {
  id: string;
  version: string;
  dependencies: string[];
  commandIds: string[];
  mutationIds: string[];
  permissions: string[];
} {
  return {
    id: 'data-source',
    version: '1.0.0',
    dependencies: ['sheet-features'],
    commandIds: registerDataSourceCommands(runtime),
    mutationIds: [...DATA_SOURCE_MUTATION_IDS],
    permissions: ['data-source.write'],
  };
}

export type { DataSourceCommandParams };
