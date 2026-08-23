import type { RangeRef, SheetTableModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import {
  findSheetTableForFilter,
  planTotalRowToggle,
  snapshotTotalRowCells,
} from './sheet-table-features';

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

  runtime.registry.registerCommand<{ sheetId: string; tableId: string; enabled: boolean }>({
    id: 'sheetTable.toggleTotalRow',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };

      const previousTable = structuredClone(sheet.sheetTables[index]!);
      if (previousTable.hasTotalRow === params.enabled) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }

      const plan = planTotalRowToggle(previousTable, params.enabled);
      const cellSnapshots = snapshotTotalRowCells(sheet, plan.totalRow, plan.startColumn, plan.endColumn);
      const cellRange: RangeRef = {
        sheetId: params.sheetId,
        startRow: plan.totalRow,
        endRow: plan.totalRow,
        startColumn: plan.startColumn,
        endColumn: plan.endColumn,
      };
      const affectedRanges: RangeRef[] = [structuredClone(previousTable.range), cellRange];
      const previousFilter = sheet.filter ? structuredClone(sheet.filter) : undefined;
      const linkedTable = findSheetTableForFilter(sheet);
      let mutationCount = 0;

      context.applyMutation({
        id: 'sheetTable.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: plan.nextTable,
        affectedRanges: [structuredClone(plan.nextTable.range)],
        inverse: [{
          id: 'sheetTable.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: previousTable,
          affectedRanges: [structuredClone(previousTable.range)],
        }],
        apply: () => {
          sheet.sheetTables[index] = structuredClone(plan.nextTable);
        },
      });
      mutationCount += 1;

      if (plan.clearTotalRow) {
        context.applyMutation({
          id: 'range.clear',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { range: cellRange, mode: 'contents' as const },
          affectedRanges: [cellRange],
          inverse: cellSnapshots.map((entry) => ({
            id: 'cell.restore',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { row: entry.row, column: entry.column, previous: entry.previous },
            affectedRanges: [cellRange],
          })),
          apply: () => {
            for (let column = plan.startColumn; column <= plan.endColumn; column++) {
              sheet.cells.delete(plan.totalRow, column);
            }
          },
        });
        mutationCount += 1;
      } else {
        context.applyMutation({
          id: 'range.set',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: {
            sheetId: params.sheetId,
            startRow: plan.totalRow,
            startColumn: plan.startColumn,
            values: plan.values,
          },
          affectedRanges: [cellRange],
          inverse: cellSnapshots.map((entry) => ({
            id: 'cell.restore',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { row: entry.row, column: entry.column, previous: entry.previous },
            affectedRanges: [cellRange],
          })),
          apply: () => {
            const rowValues = plan.values[0] ?? [];
            rowValues.forEach((value, columnOffset) => {
              sheet.cells.set(plan.totalRow, plan.startColumn + columnOffset, { ...value });
            });
          },
        });
        mutationCount += 1;
      }

      if (previousFilter && linkedTable?.id === previousTable.id) {
        const nextFilter = {
          ...previousFilter,
          range: structuredClone(plan.nextTable.range),
        };
        context.applyMutation({
          id: 'filter.set',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, filter: nextFilter },
          affectedRanges: [structuredClone(plan.nextTable.range)],
          inverse: [{
            id: 'filter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, filter: previousFilter },
            affectedRanges: [structuredClone(previousFilter.range)],
          }],
          apply: () => {
            sheet.filter = structuredClone(nextFilter);
          },
        });
        mutationCount += 1;
      }

      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
}
