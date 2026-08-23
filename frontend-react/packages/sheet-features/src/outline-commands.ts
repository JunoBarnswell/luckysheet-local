import type { OutlineGroup, OutlineModel, RangeRef } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isOutlineMutation(value: unknown): value is { sheetId: string; outline: OutlineModel } {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || !isRecord(value.outline) || !Array.isArray(value.outline.groups)) return false;
  return value.outline.groups.every((group) => isRecord(group)
    && typeof group.id === 'string'
    && (group.axis === 'row' || group.axis === 'column')
    && Number.isInteger(group.start)
    && Number.isInteger(group.end)
    && Number.isInteger(group.level)
    && typeof group.collapsed === 'boolean');
}

function outlineAffectedRanges(value: { sheetId: string; outline: OutlineModel }): RangeRef[] {
  return value.outline.groups.map((group) => group.axis === 'row'
    ? { sheetId: value.sheetId, startRow: group.start, endRow: group.end, startColumn: 0, endColumn: 0 }
    : { sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: group.start, endColumn: group.end });
}

function emptyOutline(): OutlineModel {
  return { groups: [] };
}

function outlineOf(sheet: { outline?: OutlineModel }): OutlineModel {
  return sheet.outline ?? emptyOutline();
}

function groupRange(group: OutlineGroup, sheet: { id: string; rowCount: number; columnCount: number }): RangeRef {
  return group.axis === 'row'
    ? {
      sheetId: sheet.id,
      startRow: group.start,
      endRow: group.end,
      startColumn: 0,
      endColumn: Math.max(0, sheet.columnCount - 1),
    }
    : {
      sheetId: sheet.id,
      startRow: 0,
      endRow: Math.max(0, sheet.rowCount - 1),
      startColumn: group.start,
      endColumn: group.end,
    };
}

function validateGroup(group: OutlineGroup, sheet: { rowCount: number; columnCount: number }): void {
  if (!group.id.trim()) throw new Error('Outline group id is required');
  if (group.axis !== 'row' && group.axis !== 'column') throw new Error('Outline axis is invalid');
  if (group.start < 0 || group.end < group.start) throw new Error('Outline group range is invalid');
  if (group.level < 1 || group.level > 3 || !Number.isInteger(group.level)) throw new Error('Outline level must be 1, 2, or 3');
  const limit = group.axis === 'row' ? sheet.rowCount : sheet.columnCount;
  if (group.end >= limit) throw new Error('Outline group exceeds the worksheet boundary');
}

function validateOutline(outline: OutlineModel, sheet: { rowCount: number; columnCount: number }): void {
  const ids = new Set<string>();
  for (const group of outline.groups) {
    validateGroup(group, sheet);
    if (ids.has(group.id)) throw new Error(`Outline group already exists: ${group.id}`);
    ids.add(group.id);
  }
}

export interface OutlineGroupParams {
  sheetId: string;
  group: OutlineGroup;
}

export interface OutlineToggleParams {
  sheetId: string;
  groupId: string;
  collapsed: boolean;
}

export function registerOutlineCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation({
    id: 'outline.set',
    handler: (item, context) => {
    if (!isOutlineMutation(item.params)) throw new Error('Invalid outline.set mutation payload');
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    validateOutline(params.outline, sheet);
    sheet.outline = structuredClone(params.outline);
    },
    metadata: {
      schema: { name: 'OutlineMutation', validate: isOutlineMutation },
      permission: { capability: 'sheet.outline.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: outlineAffectedRanges, mode: 'declared' },
      inverseIds: ['outline.set'],
    },
  });

  runtime.registry.registerCommand<OutlineGroupParams>({
    id: 'outline.group.add',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      validateGroup(params.group, sheet);
      const previous = structuredClone(outlineOf(sheet));
      if (previous.groups.some((group) => group.id === params.group.id)) throw new Error(`Outline group already exists: ${params.group.id}`);
      const outline = structuredClone(outlineOf(sheet));
      outline.groups.push(structuredClone(params.group));
      const affectedRanges: RangeRef[] = [groupRange(params.group, sheet)];
      context.applyMutation({
        id: 'outline.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, outline },
        affectedRanges,
        inverse: [{ id: 'outline.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, outline: previous }, affectedRanges }],
        apply: () => { sheet.outline = outline; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; groupId: string }>({
    id: 'outline.group.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = structuredClone(outlineOf(sheet));
      const removed = previous.groups.find((group) => group.id === params.groupId);
      if (!removed) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const outline = structuredClone(outlineOf(sheet));
      outline.groups = outline.groups.filter((group) => group.id !== params.groupId);
      const affectedRanges: RangeRef[] = [groupRange(removed, sheet)];
      context.applyMutation({
        id: 'outline.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, outline },
        affectedRanges,
        inverse: [{ id: 'outline.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, outline: previous }, affectedRanges }],
        apply: () => { sheet.outline = outline; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<OutlineToggleParams>({
    id: 'outline.group.toggle',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = structuredClone(outlineOf(sheet));
      const outline = structuredClone(outlineOf(sheet));
      const group = outline.groups.find((entry) => entry.id === params.groupId);
      if (!group) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      group.collapsed = params.collapsed;
      const affectedRanges: RangeRef[] = [groupRange(group, sheet)];
      context.applyMutation({
        id: 'outline.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, outline },
        affectedRanges,
        inverse: [{ id: 'outline.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, outline: previous }, affectedRanges }],
        apply: () => { sheet.outline = outline; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; level: 1 | 2 | 3 }>({
    id: 'outline.showLevel',
    execute: (params, context) => {
      if (!Number.isInteger(params.level) || params.level < 1 || params.level > 3) throw new Error('Outline level must be 1, 2, or 3');
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = structuredClone(outlineOf(sheet));
      const outline = structuredClone(outlineOf(sheet));
      for (const group of outline.groups) {
        group.collapsed = group.level > params.level;
      }
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'outline.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, outline },
        affectedRanges,
        inverse: [{ id: 'outline.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, outline: previous }, affectedRanges }],
        apply: () => { sheet.outline = outline; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

}
