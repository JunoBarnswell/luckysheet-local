import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
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

export interface SparklineGroupReplaceParams {
  sheetId: string;
  group: SparklineGroup;
}

interface SparklineMemberState {
  sparklineId: string;
  type: SparklineModel['type'];
  groupId?: string;
  showAxis?: boolean;
  showMarkers?: boolean;
}

interface SparklineGroupState {
  group: SparklineGroup;
  index: number;
}

/** Complete state for the groups/members touched by one atomic transaction. */
interface SparklineGroupStateParams {
  sheetId: string;
  groupIds: string[];
  groups: SparklineGroupState[];
  members: SparklineMemberState[];
}

function validateGroupMembers(sheet: ReturnType<CommandContext['workbook']['getSheet']>, group: SparklineGroup): void {
  const ids = new Set<string>();
  for (const sparklineId of group.sparklineIds) {
    if (ids.has(sparklineId)) throw new Error(`Sparkline group contains duplicate member: ${sparklineId}`);
    ids.add(sparklineId);
    if (!sheet.sparklines.some((entry) => entry.id === sparklineId)) throw new Error(`Unknown sparkline member: ${sparklineId}`);
  }
}

function captureGroupState(
  sheet: ReturnType<CommandContext['workbook']['getSheet']>,
  groupIds: Iterable<string>,
  additionalMemberIds: Iterable<string> = [],
): SparklineGroupStateParams {
  const ids = [...new Set(groupIds)];
  const groups = sheet.sparklineGroups
    .map((group, index) => ({ group: structuredClone(group), index }))
    .filter((entry) => ids.includes(entry.group.id));
  const memberIds = new Set<string>([...groups.flatMap((entry) => entry.group.sparklineIds), ...additionalMemberIds]);
  const members = sheet.sparklines
    .filter((sparkline) => memberIds.has(sparkline.id))
    .map((sparkline) => ({
      sparklineId: sparkline.id,
      type: sparkline.type,
      groupId: sparkline.groupId,
      showAxis: sparkline.showAxis,
      showMarkers: sparkline.showMarkers,
    }));
  return { sheetId: sheet.id, groupIds: ids, groups, members };
}

function memberStateFromSparkline(sparkline: SparklineModel): SparklineMemberState {
  return {
    sparklineId: sparkline.id,
    type: sparkline.type,
    groupId: sparkline.groupId,
    showAxis: sparkline.showAxis,
    showMarkers: sparkline.showMarkers,
  };
}

function applyGroupState(params: SparklineGroupStateParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const affected = new Set(params.groupIds);
  // Remove only the groups participating in this transition.  Their exact
  // positions are restored from the state payload below, keeping redo and
  // remote replay deterministic.
  for (let index = sheet.sparklineGroups.length - 1; index >= 0; index -= 1) {
    if (affected.has(sheet.sparklineGroups[index]!.id)) sheet.sparklineGroups.splice(index, 1);
  }
  const orderedGroups = [...params.groups].sort((left, right) => left.index - right.index);
  for (const entry of orderedGroups) {
    const index = Math.max(0, Math.min(entry.index, sheet.sparklineGroups.length));
    sheet.sparklineGroups.splice(index, 0, structuredClone(entry.group));
  }
  for (const state of params.members) {
    const sparkline = sheet.sparklines.find((entry) => entry.id === state.sparklineId);
    if (!sparkline) continue;
    sparkline.type = state.type;
    if (state.groupId === undefined) delete sparkline.groupId;
    else sparkline.groupId = state.groupId;
    if (state.showAxis === undefined) delete sparkline.showAxis;
    else sparkline.showAxis = state.showAxis;
    if (state.showMarkers === undefined) delete sparkline.showMarkers;
    else sparkline.showMarkers = state.showMarkers;
  }
}

function transitionForGroup(
  sheet: ReturnType<CommandContext['workbook']['getSheet']>,
  nextGroup: SparklineGroup,
): { before: SparklineGroupStateParams; after: SparklineGroupStateParams } {
  validateGroupMembers(sheet, nextGroup);
  const current = sheet.sparklineGroups.find((entry) => entry.id === nextGroup.id);
  const selectedIds = new Set(nextGroup.sparklineIds);
  const oldGroups = sheet.sparklineGroups.filter((group) => group.id === nextGroup.id || group.sparklineIds.some((id) => selectedIds.has(id)));
  const groupIds = [...new Set([nextGroup.id, ...oldGroups.map((group) => group.id)])];
  const before = captureGroupState(sheet, groupIds, selectedIds);
  const afterGroups = sheet.sparklineGroups
    .filter((group) => !groupIds.includes(group.id))
    .concat(oldGroups
      .filter((group) => group.id !== nextGroup.id)
      .map((group) => ({ ...structuredClone(group), sparklineIds: group.sparklineIds.filter((id) => !selectedIds.has(id)) })))
    .concat(structuredClone(nextGroup));
  const oldMemberState = new Map(before.members.map((state) => [state.sparklineId, state]));
  const memberIds = new Set([...before.members.map((state) => state.sparklineId), ...nextGroup.sparklineIds]);
  const afterMembers = [...memberIds].map((sparklineId) => {
    const previous = oldMemberState.get(sparklineId);
    const sparkline = sheet.sparklines.find((entry) => entry.id === sparklineId);
    if (!sparkline) throw new Error(`Unknown sparkline member: ${sparklineId}`);
    if (selectedIds.has(sparklineId)) return { sparklineId, type: nextGroup.type, groupId: nextGroup.id, showAxis: nextGroup.showAxis, showMarkers: nextGroup.showMarkers };
    const retained = oldGroups.find((group) => group.id !== nextGroup.id && group.sparklineIds.includes(sparklineId));
    if (retained) return previous ? { ...previous } : memberStateFromSparkline(sparkline);
    return { sparklineId, type: previous?.type ?? sparkline.type };
  });
  const groupStateById = new Map(afterGroups.map((group) => [group.id, group]));
  const after: SparklineGroupStateParams = {
    sheetId: sheet.id,
    groupIds,
    groups: [...groupStateById.values()].map((group) => ({
      group,
      index: sheet.sparklineGroups.findIndex((entry) => entry.id === group.id) >= 0
        ? sheet.sparklineGroups.findIndex((entry) => entry.id === group.id)
        : sheet.sparklineGroups.length,
    })),
    members: afterMembers,
  };
  // Replacement should retain the target group's original position; newly
  // created groups are appended after existing groups.
  const targetIndex = current ? sheet.sparklineGroups.findIndex((entry) => entry.id === current.id) : sheet.sparklineGroups.length;
  const targetState = after.groups.find((entry) => entry.group.id === nextGroup.id);
  if (targetState) targetState.index = targetIndex;
  return { before, after };
}

function transitionForGroupRemoval(
  sheet: ReturnType<CommandContext['workbook']['getSheet']>,
  groupId: string,
): { before: SparklineGroupStateParams; after: SparklineGroupStateParams } {
  const group = sheet.sparklineGroups.find((entry) => entry.id === groupId);
  if (!group) throw new Error(`Unknown sparkline group: ${groupId}`);
  const before = captureGroupState(sheet, [groupId]);
  const after: SparklineGroupStateParams = {
    sheetId: sheet.id,
    groupIds: [groupId],
    groups: [],
    members: before.members.map((state) => ({ sparklineId: state.sparklineId, type: state.type })),
  };
  return { before, after };
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
  runtime.registry.registerMutation<SparklineGroupStateParams>('sparkline.group.add', (item, context) => {
    applyGroupState(item.params, context);
  });
  runtime.registry.registerMutation<SparklineGroupStateParams>('sparkline.group.remove', (item, context) => {
    applyGroupState(item.params, context);
  });
  runtime.registry.registerMutation<SparklineGroupStateParams>('sparkline.group.replace', (item, context) => {
    applyGroupState(item.params, context);
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
      const sheet = context.workbook.getSheet(params.sheetId);
      if (sheet.sparklineGroups.some((group) => group.id === params.group.id)) throw new Error(`Sparkline group already exists: ${params.group.id}`);
      const { before, after } = transitionForGroup(sheet, structuredClone(params.group));
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupStateParams>(context, {
        id: 'sparkline.group.add',
        sheetId: params.sheetId,
        params: after,
        inverseId: 'sparkline.group.remove',
        inverseParams: before,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupStateParams>('sparkline.group.add')({
          id: 'sparkline.group.add',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: after,
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
      const sheet = context.workbook.getSheet(params.sheetId);
      const group = sheet.sparklineGroups.find((entry) => entry.id === params.groupId);
      if (!group) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const nextGroup: SparklineGroup = {
        ...structuredClone(group),
        ...structuredClone(params.patch),
        sheetId: params.sheetId,
        sparklineIds: params.patch.sparklineIds ? [...params.patch.sparklineIds] : [...group.sparklineIds],
      };
      const { before, after } = transitionForGroup(sheet, nextGroup);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupStateParams>(context, {
        id: 'sparkline.group.replace',
        sheetId: params.sheetId,
        params: after,
        inverseId: 'sparkline.group.replace',
        inverseParams: before,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupStateParams>('sparkline.group.replace')({
          id: 'sparkline.group.replace',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: after,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.group.update');

  runtime.registry.registerCommand<SparklineGroupReplaceParams>({
    id: 'sparkline.group.replace',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const current = sheet.sparklineGroups.find((group) => group.id === params.group.id);
      if (!current) throw new Error(`Unknown sparkline group: ${params.group.id}`);
      const { before, after } = transitionForGroup(sheet, structuredClone(params.group));
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupStateParams>(context, {
        id: 'sparkline.group.replace',
        sheetId: params.sheetId,
        params: after,
        inverseId: 'sparkline.group.replace',
        inverseParams: before,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupStateParams>('sparkline.group.replace')({
          id: 'sparkline.group.replace',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: after,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.group.replace');

  runtime.registry.registerCommand<{ sheetId: string; groupId: string }>({
    id: 'sparkline.group.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const { before, after } = transitionForGroupRemoval(sheet, params.groupId);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<SparklineGroupStateParams>(context, {
        id: 'sparkline.group.remove',
        sheetId: params.sheetId,
        params: after,
        inverseId: 'sparkline.group.add',
        inverseParams: before,
        affectedRanges,
        apply: () => runtime.registry.getMutation<SparklineGroupStateParams>('sparkline.group.remove')({
          id: 'sparkline.group.remove',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: after,
          affectedRanges,
        }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('sparkline.group.remove');

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
  'sparkline.group.add',
  'sparkline.group.remove',
  'sparkline.group.replace',
] as const;
