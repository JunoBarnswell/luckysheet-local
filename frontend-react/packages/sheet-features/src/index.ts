import type { CellData, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';

export interface SetCellValueParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
}

export interface SetRangeValuesParams {
  sheetId: string;
  startRow: number;
  startColumn: number;
  values: CellData[][];
}

export interface AddSheetParams {
  id: string;
  name: string;
}

function cellRange(params: SetCellValueParams): RangeRef[] {
  return [{
    sheetId: params.sheetId,
    startRow: params.row,
    endRow: params.row,
    startColumn: params.column,
    endColumn: params.column,
  }];
}

function restoreCell(workbook: WorkbookModel, item: MutationInfo<{ row: number; column: number; previous?: CellData }>): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}

export function registerSheetCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation('sheet.add', (item, context) => {
    const params = item.params as AddSheetParams;
    context.workbook.addSheet(params.id, params.name);
  });
  runtime.registry.registerMutation('sheet.remove', (item, context) => {
    context.workbook.removeSheet(item.sheetId);
  });
  runtime.registry.registerCommand<AddSheetParams>({
    id: 'sheet.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.add',
        unitId: context.workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.remove', unitId: context.workbook.unitId, sheetId: params.id, params: {}, affectedRanges }],
        apply: () => context.workbook.addSheet(params.id, params.name),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerMutation('cell.set', (item, context) => {
    const params = item.params as SetCellValueParams;
    context.workbook.getSheet(params.sheetId).cells.set(params.row, params.column, { ...params.value });
  });
  runtime.registry.registerMutation('cell.restore', (item, context) => restoreCell(context.workbook, item as MutationInfo<{ row: number; column: number; previous?: CellData }>));

  runtime.registry.registerCommand<SetCellValueParams>({
    id: 'sheet.cell.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const affectedRanges = cellRange(params);
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: params.row, column: params.column, previous },
          affectedRanges,
        }],
        apply: () => sheet.cells.set(params.row, params.column, { ...params.value }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<SetRangeValuesParams>({
    id: 'sheet.range.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [];
      for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
        const rowValues = params.values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
          const row = params.startRow + rowOffset;
          const column = params.startColumn + columnOffset;
          previous.push({ row, column, value: sheet.cells.get(row, column) });
          affectedRanges.push({ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column });
        }
      }
      context.applyMutation({
        id: 'range.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: item.row, column: item.column, previous: item.value },
          affectedRanges: [{ sheetId: params.sheetId, startRow: item.row, endRow: item.row, startColumn: item.column, endColumn: item.column }],
        })),
        apply: () => {
          for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
            const rowValues = params.values[rowOffset] ?? [];
            for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
              const value = rowValues[columnOffset];
              if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, { ...value });
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerMutation('range.set', (item, context) => {
    const params = item.params as SetRangeValuesParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
      const rowValues = params.values[rowOffset] ?? [];
      for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
        const value = rowValues[columnOffset];
        if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, { ...value });
      }
    }
  });
}
