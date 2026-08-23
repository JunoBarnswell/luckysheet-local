import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { RangeRef, SparklineGroup, SparklineModel } from '@react-sheets/core-model';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from '../../command-helpers';

export interface SparklineInsertParams {
  sheetId: string;
  sparkline: SparklineModel;
}

export interface SparklineUpdateParams {
  sheetId: string;
  sparklineId: string;
  patch: Partial<SparklineModel>;
}

export interface SparklineGroupCreateParams {
  sheetId: string;
  group: SparklineGroup;
}

export interface SparklineGroupUpdateParams {
  sheetId: string;
  groupId: string;
  patch: Partial<SparklineGroup>;
}

export interface SparklineInsertDialogParams {
  sheetId: string;
  sparklineId: string;
  dataRange: RangeRef;
  location: { row: number; column: number };
  type: SparklineModel['type'];
  groupId?: string;
}

export function registerSparklineCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<SparklineInsertParams>(runtime, 'sparkline.add', (params, context) => {
    context.workbook.getSheet(params.sheetId).sparklines.push(structuredClone(params.sparkline));
  });
  registerMutationHandler<{ sheetId: string; sparklineId: string }>(runtime, 'sparkline.remove', (params, context) => {
    removeById(context.workbook.getSheet(params.sheetId).sparklines, params.sparklineId);
  });
  registerMutationHandler<SparklineUpdateParams>(runtime, 'sparkline.update', (params, context) => {
    const sparkline = context.workbook.getSheet(params.sheetId).sparklines.find((entry) => entry.id === params.sparklineId);
    if (sparkline) Object.assign(sparkline, structuredClone(params.patch));
  });
  registerMutationHandler<SparklineGroupCreateParams>(runtime, 'sparkline.group.create', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    sheet.sparklineGroups.push(structuredClone(params.group));
    for (const sparklineId of params.group.sparklineIds) {
      const sparkline = sheet.sparklines.find((entry) => entry.id === sparklineId);
      if (sparkline) {
        sparkline.groupId = params.group.id;
        sparkline.showAxis = params.group.showAxis;
        sparkline.showMarkers = params.group.showMarkers;
      }
    }
  });
  registerMutationHandler<{ sheetId: string; groupId: string }>(runtime, 'sparkline.group.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    removeById(sheet.sparklineGroups, params.groupId);
    for (const sparkline of sheet.sparklines) {
      if (sparkline.groupId !== params.groupId) continue;
      delete sparkline.groupId;
      delete sparkline.showAxis;
      delete sparkline.showMarkers;
    }
  });
  registerMutationHandler<SparklineGroupUpdateParams>(runtime, 'sparkline.group.update', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const group = sheet.sparklineGroups.find((entry) => entry.id === params.groupId);
    if (!group) return;
    Object.assign(group, structuredClone(params.patch));
    for (const sparkline of sheet.sparklines) {
      if (sparkline.groupId !== params.groupId) continue;
      if (params.patch.showAxis != null) sparkline.showAxis = params.patch.showAxis;
      if (params.patch.showMarkers != null) sparkline.showMarkers = params.patch.showMarkers;
    }
  });

  runtime.registry.registerCommand<SparklineInsertParams>({
    id: 'sparkline.insert',
    execute: (params, context) => {
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineInsertParams, { sheetId: string; sparklineId: string }>(context, {
        id: 'sparkline.add',
        sheetId: params.sheetId,
        params,
        inverseId: 'sparkline.remove',
        inverseParams: { sheetId: params.sheetId, sparklineId: params.sparkline.id },
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineInsertParams>('sparkline.add')({
          id: 'sparkline.add',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.insert');

  runtime.registry.registerCommand<SparklineInsertDialogParams>({
    id: 'sparkline.insertDataLocation',
    execute: (params, context) => {
      const sparkline: SparklineModel = {
        id: params.sparklineId,
        sheetId: params.sheetId,
        anchor: { row: params.location.row, column: params.location.column },
        sourceRange: structuredClone(params.dataRange),
        type: params.type,
        color: '#2563eb',
        negativeColor: '#ef4444',
        groupId: params.groupId,
        showAxis: false,
        showMarkers: false,
      };
      return runtime.registry.getCommand<SparklineInsertParams>('sparkline.insert').execute({ sheetId: params.sheetId, sparkline }, context);
    },
  });
  commandIds.push('sparkline.insertDataLocation');

  runtime.registry.registerCommand<SparklineUpdateParams>({
    id: 'sparkline.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const sparkline = sheet.sparklines.find((entry) => entry.id === params.sparklineId);
      const previous = sparkline ? structuredClone(sparkline) : undefined;
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineUpdateParams>(context, {
        id: 'sparkline.update',
        sheetId: params.sheetId,
        params,
        inverseParams: previous
          ? { sheetId: params.sheetId, sparklineId: params.sparklineId, patch: previous }
          : params,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineUpdateParams>('sparkline.update')({
          id: 'sparkline.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: sparkline ? 1 : 0, affectedRanges };
    },
  });
  commandIds.push('sparkline.update');

  runtime.registry.registerCommand<SparklineGroupCreateParams>({
    id: 'sparkline.group.create',
    execute: (params, context) => {
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupCreateParams, { sheetId: string; groupId: string }>(context, {
        id: 'sparkline.group.create',
        sheetId: params.sheetId,
        params,
        inverseId: 'sparkline.group.remove',
        inverseParams: { sheetId: params.sheetId, groupId: params.group.id },
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupCreateParams>('sparkline.group.create')({
          id: 'sparkline.group.create',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.group.create');

  runtime.registry.registerCommand<SparklineGroupUpdateParams>({
    id: 'sparkline.group.update',
    execute: (params, context) => {
      const group = context.workbook.getSheet(params.sheetId).sparklineGroups.find((entry) => entry.id === params.groupId);
      const previous = group ? structuredClone(group) : undefined;
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupUpdateParams>(context, {
        id: 'sparkline.group.update',
        sheetId: params.sheetId,
        params,
        inverseParams: previous
          ? { sheetId: params.sheetId, groupId: params.groupId, patch: previous }
          : params,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupUpdateParams>('sparkline.group.update')({
          id: 'sparkline.group.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: group ? 1 : 0, affectedRanges };
    },
  });
  commandIds.push('sparkline.group.update');

  runtime.registry.registerCommand<{ sheetId: string; sparklineId: string }>({
    id: 'sparkline.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const sparkline = sheet.sparklines.find((entry) => entry.id === params.sparklineId);
      const previous = sparkline ? structuredClone(sparkline) : undefined;
      const affectedRanges = sheetRange(params.sheetId);
      if (!previous) return { operationId: context.operationId, mutationCount: 0, affectedRanges };
      applyTrackedMutation<{ sheetId: string; sparklineId: string }, SparklineInsertParams>(context, {
        id: 'sparkline.remove',
        sheetId: params.sheetId,
        params,
        inverseId: 'sparkline.add',
        inverseParams: { sheetId: params.sheetId, sparkline: previous },
        affectedRanges,
        apply: () => runtime.registry.getMutation<{ sheetId: string; sparklineId: string }>('sparkline.remove')({
          id: 'sparkline.remove',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.remove');

  return commandIds;
}

export const SPARKLINE_MUTATION_IDS = [
  'sparkline.add',
  'sparkline.remove',
  'sparkline.update',
  'sparkline.group.create',
  'sparkline.group.remove',
  'sparkline.group.update',
] as const;
