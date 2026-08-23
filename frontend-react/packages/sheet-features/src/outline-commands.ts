import type { OutlineGroup, OutlineModel, RangeRef } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

function emptyOutline(): OutlineModel {
  return { groups: [] };
}

function outlineOf(sheet: { outline?: OutlineModel }): OutlineModel {
  return sheet.outline ?? emptyOutline();
}

function rowRange(group: OutlineGroup, columnCount: number, sheetId: string): RangeRef {
  return {
    sheetId,
    startRow: group.start,
    endRow: group.end,
    startColumn: 0,
    endColumn: Math.max(0, columnCount - 1),
  };
}

function validateGroup(group: OutlineGroup, sheet: { rowCount: number; columnCount: number }): void {
  if (!group.id.trim()) throw new Error('Outline group id is required');
  if (group.start < 0 || group.end < group.start) throw new Error('Outline group range is invalid');
  if (group.level < 1 || group.level > 3 || !Number.isInteger(group.level)) throw new Error('Outline level must be 1, 2, or 3');
  const limit = group.axis === 'row' ? sheet.rowCount : sheet.columnCount;
  if (group.end >= limit) throw new Error('Outline group exceeds the worksheet boundary');
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
  runtime.registry.registerMutation('outline.set', (item, context) => {
    const params = item.params as { sheetId: string; outline: OutlineModel };
    context.workbook.getSheet(params.sheetId).outline = structuredClone(params.outline);
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
      const affectedRanges: RangeRef[] = [rowRange(params.group, sheet.columnCount, params.sheetId)];
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
      const affectedRanges: RangeRef[] = [rowRange(removed, sheet.columnCount, params.sheetId)];
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
      const affectedRanges: RangeRef[] = [rowRange(group, sheet.columnCount, params.sheetId)];
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
