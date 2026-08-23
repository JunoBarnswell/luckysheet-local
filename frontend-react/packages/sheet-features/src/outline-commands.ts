import type { OutlineGroup, OutlineModel, RangeRef } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { computeOutlineHiddenRows } from './data-features';

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
      const previous = structuredClone(outlineOf(sheet));
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
        if (group.axis === 'row') group.collapsed = group.level > params.level;
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

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'outline.applyHiddenRows',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const hidden = computeOutlineHiddenRows(sheet);
      const previous = [...sheet.hiddenRows];
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'rows.hidden.restore',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, indices: [...hidden] },
        affectedRanges,
        inverse: [{ id: 'rows.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => {
          sheet.hiddenRows.clear();
          for (const row of hidden) sheet.hiddenRows.add(row);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}
