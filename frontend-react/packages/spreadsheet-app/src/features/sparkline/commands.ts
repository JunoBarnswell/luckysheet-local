import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type { RangeRef, SparklineGroup, SparklineModel } from '@react-sheets/core-model';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function removeById<T extends { id: string }>(items: T[], id: string): T | undefined {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return undefined;
  return items.splice(index, 1)[0];
}

export interface SparklineInsertParams {
  sheetId: string;
  sparkline: SparklineModel;
  /** Optional group state restored after the member is inserted during undo. */
  groupState?: SparklineGroupStateParams;
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

interface SparklineRemoveParams {
  sheetId: string;
  sparklineId: string;
  /** Group state after this sparkline has been detached. */
  groupState?: SparklineGroupStateParams;
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
  const groupIds = new Set(params.groupIds);
  if (groupIds.size !== params.groupIds.length) throw new Error('Sparkline group state contains duplicate group ids');
  for (const entry of params.groups) {
    if (entry.group.sheetId !== params.sheetId) throw new Error(`Sparkline group ${entry.group.id} targets another sheet`);
    if (!groupIds.has(entry.group.id)) throw new Error(`Sparkline group ${entry.group.id} is outside the transaction`);
    validateGroupMembers(sheet, entry.group);
  }
  for (const member of params.members) {
    if (!sheet.sparklines.some((entry) => entry.id === member.sparklineId)) throw new Error(`Unknown sparkline member: ${member.sparklineId}`);
    if (member.groupId !== undefined && !groupIds.has(member.groupId) && !sheet.sparklineGroups.some((group) => group.id === member.groupId)) {
      throw new Error(`Unknown sparkline group: ${member.groupId}`);
    }
  }
  const affected = groupIds;
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
    if (!sparkline) throw new Error(`Unknown sparkline member: ${state.sparklineId}`);
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
  if (nextGroup.sheetId !== sheet.id) throw new Error(`Sparkline group ${nextGroup.id} targets another sheet`);
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
  color?: string;
  negativeColor?: string;
  highlightMax?: boolean;
  highlightMin?: boolean;
  highlightFirst?: boolean;
  highlightLast?: boolean;
  highlightNegative?: boolean;
}

const SPARKLINE_TYPES = new Set<SparklineModel['type']>(['line', 'column', 'win-loss']);

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const isRange = (value: unknown): value is RangeRef => isRecord(value)
  && isNonEmptyString(value.sheetId)
  && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
  && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn)
  && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
  && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
const isSparkline = (value: unknown): value is SparklineModel => isRecord(value)
  && isNonEmptyString(value.id) && isNonEmptyString(value.sheetId)
  && isRecord(value.anchor) && Number.isSafeInteger(value.anchor.row) && Number.isSafeInteger(value.anchor.column)
  && Number(value.anchor.row) >= 0 && Number(value.anchor.column) >= 0
  && isRange(value.sourceRange) && SPARKLINE_TYPES.has(value.type as SparklineModel['type'])
  && (value.groupId === undefined || isNonEmptyString(value.groupId));
const isSparklineGroup = (value: unknown): value is SparklineGroup => isRecord(value)
  && isNonEmptyString(value.id) && isNonEmptyString(value.sheetId)
  && SPARKLINE_TYPES.has(value.type as SparklineModel['type'])
  && Array.isArray(value.sparklineIds) && value.sparklineIds.every(isNonEmptyString)
  && (value.showAxis === undefined || typeof value.showAxis === 'boolean')
  && (value.showMarkers === undefined || typeof value.showMarkers === 'boolean');
const isMemberState = (value: unknown): value is SparklineMemberState => isRecord(value)
  && isNonEmptyString(value.sparklineId) && SPARKLINE_TYPES.has(value.type as SparklineModel['type'])
  && (value.groupId === undefined || isNonEmptyString(value.groupId))
  && (value.showAxis === undefined || typeof value.showAxis === 'boolean')
  && (value.showMarkers === undefined || typeof value.showMarkers === 'boolean');
const isGroupState = (value: unknown): value is SparklineGroupStateParams => isRecord(value)
  && isNonEmptyString(value.sheetId)
  && Array.isArray(value.groupIds) && value.groupIds.every(isNonEmptyString)
  && Array.isArray(value.groups)
  && value.groups.every((entry) => isRecord(entry) && isSparklineGroup(entry.group) && Number.isSafeInteger(entry.index) && Number(entry.index) >= 0)
  && Array.isArray(value.members) && value.members.every(isMemberState);
const isSparklineInsert = (value: unknown): value is SparklineInsertParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isSparkline(value.sparkline)
  && value.sparkline.sheetId === value.sheetId
  && (value.groupState === undefined || isGroupState(value.groupState));
const isSparklineRemove = (value: unknown): value is SparklineRemoveParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.sparklineId)
  && (value.groupState === undefined || isGroupState(value.groupState));
const isSparklineUpdate = (value: unknown): value is SparklineUpdateParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.sparklineId) && isRecord(value.patch);
const isGroupStateParams = (value: unknown): value is SparklineGroupStateParams => isGroupState(value);

function sparklineAffectedRanges(value: unknown): ReturnType<typeof sheetRange> {
  if (isRecord(value) && isNonEmptyString(value.sheetId)) return sheetRange(value.sheetId);
  return [];
}

function applySparklineAdd(context: CommandContext, params: SparklineInsertParams): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  if (sheet.sparklines.some((entry) => entry.id === params.sparkline.id)) throw new Error(`Sparkline already exists: ${params.sparkline.id}`);
  if (params.sparkline.sheetId !== params.sheetId) throw new Error(`Sparkline ${params.sparkline.id} targets another sheet`);
  if (params.sparkline.groupId !== undefined && !params.groupState) {
    validateSparklineState(context, params.sparkline);
    throw new Error('Insert the sparkline before assigning it to a group');
  }
  const sparkline = structuredClone(params.sparkline);
  if (params.groupState) {
    delete sparkline.groupId;
    delete sparkline.showAxis;
    delete sparkline.showMarkers;
  }
  validateSparklineState(context, sparkline);
  sheet.sparklines.push(sparkline);
  if (params.groupState) applyGroupState(params.groupState, context);
}

function applySparklineRemove(context: CommandContext, params: SparklineRemoveParams): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  const sparkline = sheet.sparklines.find((entry) => entry.id === params.sparklineId);
  if (!sparkline) throw new Error(`Unknown sparkline: ${params.sparklineId}`);
  if (sparkline.groupId !== undefined && !params.groupState) throw new Error(`Sparkline ${params.sparklineId} requires a group transition when removed`);
  if (params.groupState) applyGroupState(params.groupState, context);
  if (!removeById(sheet.sparklines, params.sparklineId)) throw new Error(`Unknown sparkline: ${params.sparklineId}`);
}

function sameSparklineValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => sameSparklineValue(value, right[index]));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const keys = [...new Set([...Object.keys(leftRecord), ...Object.keys(rightRecord)])]
    .filter((key) => leftRecord[key] !== undefined || rightRecord[key] !== undefined)
    .sort();
  return keys.every((key) => sameSparklineValue(leftRecord[key], rightRecord[key]));
}

function applySparklineUpdate(context: CommandContext, params: SparklineUpdateParams): void {
  const sparkline = context.workbook.getSheet(params.sheetId).sparklines.find((entry) => entry.id === params.sparklineId);
  if (!sparkline) throw new Error(`Unknown sparkline: ${params.sparklineId}`);
  const next = { ...structuredClone(sparkline), ...structuredClone(params.patch), sheetId: params.sheetId };
  validateSparklineState(context, next);
  Object.assign(sparkline, next);
}

function executeSparklineInsert(params: SparklineInsertParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'sparkline.add',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{ id: 'sparkline.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, sparklineId: params.sparkline.id }, affectedRanges }],
    apply: () => applySparklineAdd(context, params),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function validateSparklineState(context: CommandContext, sparkline: SparklineModel): void {
  if (!sparkline.id.trim()) throw new Error('Sparkline id is required');
  const target = context.workbook.getSheet(sparkline.sheetId);
  if (!SPARKLINE_TYPES.has(sparkline.type)) throw new Error(`Unsupported sparkline type: ${String(sparkline.type)}`);
  if (sparkline.anchor.row < 0 || sparkline.anchor.column < 0) throw new Error('Sparkline anchor is invalid');
  const source = sparkline.sourceRange;
  const sourceSheet = context.workbook.getSheet(source.sheetId);
  if (source.startRow < 0 || source.endRow < source.startRow || source.startColumn < 0 || source.endColumn < source.startColumn) {
    throw new Error('Sparkline source range is invalid');
  }
  if (source.endRow >= sourceSheet.rowCount || source.endColumn >= sourceSheet.columnCount) throw new Error('Sparkline source range exceeds worksheet bounds');
  if (sparkline.groupId !== undefined) {
    const group = target.sparklineGroups.find((entry) => entry.id === sparkline.groupId);
    if (!group) throw new Error(`Unknown sparkline group: ${sparkline.groupId}`);
    if (!group.sparklineIds.includes(sparkline.id)) throw new Error(`Sparkline ${sparkline.id} is not a member of group ${sparkline.groupId}`);
  }
}

export function registerSparklineCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  runtime.registry.registerMutation<SparklineInsertParams>({
      id: 'sparkline.add',
      handler: (item, context) => {
    applySparklineAdd(context, item.params);
  },
      metadata: {
    schema: { name: 'SparklineInsertParams', validate: isSparklineInsert },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<SparklineRemoveParams>({
      id: 'sparkline.remove',
      handler: (item, context) => {
    applySparklineRemove(context, item.params);
  },
      metadata: {
    schema: { name: 'SparklineRemoveParams', validate: isSparklineRemove },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.add'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<SparklineUpdateParams>({
      id: 'sparkline.update',
      handler: (item, context) => {
    applySparklineUpdate(context, item.params);
  },
      metadata: {
    schema: { name: 'SparklineUpdateParams', validate: isSparklineUpdate },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.update'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<SparklineGroupStateParams>({
      id: 'sparkline.group.add',
      handler: (item, context) => {
    applyGroupState(item.params, context);
  },
      metadata: {
    schema: { name: 'SparklineGroupStateParams', validate: isGroupStateParams },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.group.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<SparklineGroupStateParams>({
      id: 'sparkline.group.remove',
      handler: (item, context) => {
    applyGroupState(item.params, context);
  },
      metadata: {
    schema: { name: 'SparklineGroupStateParams', validate: isGroupStateParams },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.group.add'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<SparklineGroupStateParams>({
      id: 'sparkline.group.replace',
      handler: (item, context) => {
    applyGroupState(item.params, context);
  },
      metadata: {
    schema: { name: 'SparklineGroupStateParams', validate: isGroupStateParams },
    permission: { capability: 'sparkline.edit' },
    affectedRanges: { resolve: sparklineAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['sparkline.group.replace'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerCommand<SparklineInsertParams>({
    id: 'sparkline.insert',
    execute: (params, context) => executeSparklineInsert(params, context),
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
        highlightMax: params.highlightMax,
        highlightMin: params.highlightMin,
        highlightFirst: params.highlightFirst,
        highlightLast: params.highlightLast,
        highlightNegative: params.highlightNegative,
        groupId: params.groupId,
        showAxis: false,
        showMarkers: false,
      };
      return executeSparklineInsert({ sheetId: params.sheetId, sparkline }, context);
    },
  });
  commandIds.push('sparkline.insertDataLocation');

  runtime.registry.registerCommand<SparklineUpdateParams>({
    id: 'sparkline.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const sparkline = sheet.sparklines.find((entry) => entry.id === params.sparklineId);
      if (!sparkline) throw new Error(`Unknown sparkline: ${params.sparklineId}`);
      const previous = structuredClone(sparkline);
      const affectedRanges = sheetRange(params.sheetId);
      const next = { ...structuredClone(sparkline), ...structuredClone(params.patch), sheetId: params.sheetId };
      validateSparklineState(context, next);
      if (sameSparklineValue(previous, next)) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      context.applyMutation({
        id: 'sparkline.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sparkline.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, sparklineId: params.sparklineId, patch: previous }, affectedRanges }],
        apply: () => applySparklineUpdate(context, params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
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
      context.applyMutation({
        id: 'sparkline.group.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: after,
        affectedRanges,
        inverse: [{ id: 'sparkline.group.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: before, affectedRanges }],
        apply: () => applyGroupState(after, context),
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
      if (!group) throw new Error(`Unknown sparkline group: ${params.groupId}`);
      const nextGroup: SparklineGroup = {
        ...structuredClone(group),
        ...structuredClone(params.patch),
        sheetId: params.sheetId,
        sparklineIds: params.patch.sparklineIds ? [...params.patch.sparklineIds] : [...group.sparklineIds],
      };
      const { before, after } = transitionForGroup(sheet, nextGroup);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'sparkline.group.replace',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: after,
        affectedRanges,
        inverse: [{ id: 'sparkline.group.replace', unitId: context.workbook.unitId, sheetId: params.sheetId, params: before, affectedRanges }],
        apply: () => applyGroupState(after, context),
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
      context.applyMutation({
        id: 'sparkline.group.replace',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: after,
        affectedRanges,
        inverse: [{ id: 'sparkline.group.replace', unitId: context.workbook.unitId, sheetId: params.sheetId, params: before, affectedRanges }],
        apply: () => applyGroupState(after, context),
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
      context.applyMutation({
        id: 'sparkline.group.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: after,
        affectedRanges,
        inverse: [{ id: 'sparkline.group.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: before, affectedRanges }],
        apply: () => applyGroupState(after, context),
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
      if (!sparkline) throw new Error(`Unknown sparkline: ${params.sparklineId}`);
      const previous = structuredClone(sparkline);
      const affectedRanges = sheetRange(params.sheetId);
      const groups = sheet.sparklineGroups.filter((group) => group.sparklineIds.includes(params.sparklineId));
      const groupTransition = groups.length > 0
        ? (() => {
          const before = captureGroupState(sheet, groups.map((group) => group.id));
          const after = {
            sheetId: sheet.id,
            groupIds: before.groupIds,
            groups: before.groups.map((entry) => ({ ...entry, group: { ...entry.group, sparklineIds: entry.group.sparklineIds.filter((id) => id !== params.sparklineId) } })),
            members: before.members.filter((member) => member.sparklineId !== params.sparklineId),
          } satisfies SparklineGroupStateParams;
          return { before, after };
        })()
        : undefined;
      const removeParams: SparklineRemoveParams = { sheetId: params.sheetId, sparklineId: params.sparklineId, groupState: groupTransition?.after };
      context.applyMutation({
        id: 'sparkline.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: removeParams,
        affectedRanges,
        inverse: [{ id: 'sparkline.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, sparkline: previous, groupState: groupTransition?.before }, affectedRanges }],
        apply: () => applySparklineRemove(context, removeParams),
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
