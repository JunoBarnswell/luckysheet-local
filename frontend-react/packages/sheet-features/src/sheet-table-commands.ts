import type { RangeRef, SheetTableModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

export interface AddSheetTableParams extends SheetTableModel {}

export function registerSheetTableCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation('sheetTable.add', (item, context) => {
    const table = item.params as SheetTableModel;
    context.workbook.getSheet(table.sheetId).sheetTables.push(structuredClone(table));
  });
  runtime.registry.registerMutation('sheetTable.remove', (item, context) => {
    const params = item.params as { sheetId: string; tableId: string };
    const sheet = context.workbook.getSheet(params.sheetId);
    const index = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
    if (index >= 0) sheet.sheetTables.splice(index, 1);
  });
  runtime.registry.registerMutation('sheetTable.update', (item, context) => {
    const table = item.params as SheetTableModel;
    const sheet = context.workbook.getSheet(table.sheetId);
    const index = sheet.sheetTables.findIndex((entry) => entry.id === table.id);
    if (index >= 0) sheet.sheetTables[index] = structuredClone(table);
    else sheet.sheetTables.push(structuredClone(table));
  });

  runtime.registry.registerCommand<AddSheetTableParams>({
    id: 'sheetTable.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [structuredClone(params.range)];
      context.applyMutation({
        id: 'sheetTable.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheetTable.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, tableId: params.id }, affectedRanges }],
        apply: () => {
          context.workbook.getSheet(params.sheetId).sheetTables.push(structuredClone(params));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; tableId: string }>({
    id: 'sheetTable.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.sheetTables[index]!);
      const affectedRanges: RangeRef[] = [structuredClone(previous.range)];
      context.applyMutation({
        id: 'sheetTable.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheetTable.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => {
          const idx = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
          if (idx >= 0) sheet.sheetTables.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; tableId: string }>({
    id: 'sheetTable.convertToRange',
    execute: (params, context) => {
      return runtime.execute('sheetTable.remove', params);
    },
  });
}
