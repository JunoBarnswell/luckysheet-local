import type {
  CellData,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  FreezeModel,
  MergeSpan,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { shiftFormula } from './clipboard';

export * from './clipboard';

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

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  mode?: 'all' | 'contents' | 'formats';
}

export interface AddSheetParams {
  id: string;
  name: string;
  rowCount?: number;
  columnCount?: number;
}

export interface RenameSheetParams {
  sheetId: string;
  name: string;
}

export interface InsertRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface DeleteRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface ResizeRowParams {
  sheetId: string;
  row: number;
  height: number;
}

export interface InsertColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface DeleteColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface ResizeColumnParams {
  sheetId: string;
  column: number;
  width: number;
}

export interface SetMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface RemoveMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface SetFreezeParams {
  sheetId: string;
  freeze: FreezeModel;
}

export interface SetRangeStyleParams {
  sheetId: string;
  range: RangeRef;
  style: Partial<CellStyle>;
}

export interface SortRangeParams {
  sheetId: string;
  range: RangeRef;
  sortColumn: number;
  ascending: boolean;
  hasHeader?: boolean;
}

export interface AutoFillParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
}

export interface AddConditionalFormatParams {
  sheetId: string;
  rule: ConditionalFormatRule;
}

export interface AddDataValidationParams {
  sheetId: string;
  rule: DataValidationRule;
}

function cellRange(params: SetCellValueParams): RangeRef[] {
  return [
    {
      sheetId: params.sheetId,
      startRow: params.row,
      endRow: params.row,
      startColumn: params.column,
      endColumn: params.column,
    },
  ];
}

function restoreCell(
  workbook: WorkbookModel,
  item: MutationInfo<{ row: number; column: number; previous?: CellData }>,
): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}

export function registerSheetCommands(runtime: CommandRuntime): void {
  // 1. Sheet mutations & commands
  runtime.registry.registerMutation('sheet.add', (item, context) => {
    const params = item.params as AddSheetParams;
    context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount);
  });
  runtime.registry.registerMutation('sheet.remove', (item, context) => {
    context.workbook.removeSheet(item.sheetId);
  });
  runtime.registry.registerMutation('sheet.rename', (item, context) => {
    const params = item.params as RenameSheetParams;
    context.workbook.getSheet(params.sheetId).name = params.name;
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
        inverse: [
          {
            id: 'sheet.remove',
            unitId: context.workbook.unitId,
            sheetId: params.id,
            params: {},
            affectedRanges,
          },
        ],
        apply: () =>
          context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RenameSheetParams>({
    id: 'sheet.rename',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousName = sheet.name;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.rename',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.rename',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, name: previousName },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.name = params.name;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 2. Cell mutations & commands
  runtime.registry.registerMutation('cell.set', (item, context) => {
    const params = item.params as SetCellValueParams;
    context.workbook
      .getSheet(params.sheetId)
      .cells.set(params.row, params.column, { ...params.value });
  });
  runtime.registry.registerMutation('cell.restore', (item, context) =>
    restoreCell(
      context.workbook,
      item as MutationInfo<{ row: number; column: number; previous?: CellData }>,
    ),
  );

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
        inverse: [
          {
            id: 'cell.restore',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { row: params.row, column: params.column, previous },
            affectedRanges,
          },
        ],
        apply: () => sheet.cells.set(params.row, params.column, { ...params.value }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 3. Range set
  runtime.registry.registerMutation('range.set', (item, context) => {
    const params = item.params as SetRangeValuesParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
      const rowValues = params.values[rowOffset] ?? [];
      for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
        const value = rowValues[columnOffset];
        if (value) {
          sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, {
            ...value,
          });
        }
      }
    }
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
          affectedRanges.push({
            sheetId: params.sheetId,
            startRow: row,
            endRow: row,
            startColumn: column,
            endColumn: column,
          });
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
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => {
          for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
            const rowValues = params.values[rowOffset] ?? [];
            for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
              const value = rowValues[columnOffset];
              if (value)
                sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, {
                  ...value,
                });
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 4. Range Clear
  runtime.registry.registerCommand<ClearRangeParams>({
    id: 'sheet.range.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [params.range];

      for (let r = params.range.startRow; r <= params.range.endRow; r++) {
        for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
          const cell = sheet.cells.get(r, c);
          if (cell) {
            previous.push({ row: r, column: c, value: structuredClone(cell) });
          }
        }
      }

      context.applyMutation({
        id: 'range.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: item.row, column: item.column, previous: item.value },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => {
          for (let r = params.range.startRow; r <= params.range.endRow; r++) {
            for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
              if (params.mode === 'formats') {
                const cell = sheet.cells.get(r, c);
                if (cell) {
                  cell.style = undefined;
                  cell.styleId = undefined;
                  cell.numberFormat = undefined;
                }
              } else {
                sheet.cells.delete(r, c);
              }
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 5. Range Style batch application
  runtime.registry.registerCommand<SetRangeStyleParams>({
    id: 'sheet.style.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [params.range];

      for (let r = params.range.startRow; r <= params.range.endRow; r++) {
        for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
          const cell = sheet.cells.get(r, c);
          previous.push({ row: r, column: c, value: cell ? structuredClone(cell) : undefined });
        }
      }

      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: item.row, column: item.column, previous: item.value },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => {
          for (let r = params.range.startRow; r <= params.range.endRow; r++) {
            for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
              let cell = sheet.cells.get(r, c);
              if (!cell) {
                cell = { value: null };
                sheet.cells.set(r, c, cell);
              }
              cell.style = { ...(cell.style ?? {}), ...params.style };
              if (params.style.numberFormat !== undefined) {
                cell.numberFormat = params.style.numberFormat;
              }
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 6. Merge commands & mutations
  runtime.registry.registerMutation('merge.set', (item, context) => {
    const params = item.params as SetMergeParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    const span: MergeSpan = {
      range: params.range,
      anchor: { row: params.range.startRow, column: params.range.startColumn },
    };
    sheet.merges.push(span);
  });
  runtime.registry.registerMutation('merge.remove', (item, context) => {
    const params = item.params as RemoveMergeParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    const idx = sheet.merges.findIndex(
      (m) =>
        m.range.startRow === params.range.startRow &&
        m.range.startColumn === params.range.startColumn,
    );
    if (idx >= 0) sheet.merges.splice(idx, 1);
  });

  runtime.registry.registerCommand<SetMergeParams>({
    id: 'sheet.merge.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const span: MergeSpan = {
        range: params.range,
        anchor: { row: params.range.startRow, column: params.range.startColumn },
      };
      const affectedRanges = [params.range];

      context.applyMutation({
        id: 'merge.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: params.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.push(span);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RemoveMergeParams>({
    id: 'sheet.merge.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const idx = sheet.merges.findIndex(
        (m) =>
          m.range.startRow === params.range.startRow &&
          m.range.startColumn === params.range.startColumn,
      );
      const affectedRanges = [params.range];
      if (idx < 0)
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previousSpan = sheet.merges[idx]!;

      context.applyMutation({
        id: 'merge.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: previousSpan.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 7. Freeze commands
  runtime.registry.registerMutation('freeze.set', (item, context) => {
    const params = item.params as SetFreezeParams;
    context.workbook.getSheet(params.sheetId).freeze = { ...params.freeze };
  });

  runtime.registry.registerCommand<SetFreezeParams>({
    id: 'sheet.freeze.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const prevFreeze = { ...sheet.freeze };
      const affectedRanges: RangeRef[] = [];

      context.applyMutation({
        id: 'freeze.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'freeze.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, freeze: prevFreeze },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.freeze = { ...params.freeze };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 8. Row and Column resizing
  runtime.registry.registerMutation('row.resize', (item, context) => {
    const params = item.params as ResizeRowParams;
    context.workbook.getSheet(params.sheetId).rowHeights[params.row] = params.height;
  });
  runtime.registry.registerMutation('column.resize', (item, context) => {
    const params = item.params as ResizeColumnParams;
    context.workbook.getSheet(params.sheetId).columnWidths[params.column] = params.width;
  });

  runtime.registry.registerCommand<ResizeRowParams>({
    id: 'sheet.row.resize',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const prevHeight = sheet.rowHeights[params.row] ?? 32;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'row.resize',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'row.resize',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: params.row, height: prevHeight },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.rowHeights[params.row] = params.height;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<ResizeColumnParams>({
    id: 'sheet.column.resize',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const prevWidth = sheet.columnWidths[params.column] ?? 128;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'column.resize',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'column.resize',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, column: params.column, width: prevWidth },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.columnWidths[params.column] = params.width;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 9. Range Sorting
  runtime.registry.registerCommand<SortRangeParams>({
    id: 'sheet.sort',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const { range, sortColumn, ascending, hasHeader } = params;
      const startR = hasHeader ? range.startRow + 1 : range.startRow;
      const rowCount = range.endRow - startR + 1;
      if (rowCount <= 1)
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };

      // Collect rows
      const rows: Array<{ originalIndex: number; cells: Array<CellData | undefined> }> = [];
      for (let r = startR; r <= range.endRow; r++) {
        const rowCells: Array<CellData | undefined> = [];
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          rowCells.push(sheet.cells.get(r, c));
        }
        rows.push({ originalIndex: r, cells: rowCells });
      }

      // Sort
      const colOffset = sortColumn - range.startColumn;
      rows.sort((a, b) => {
        const valA = a.cells[colOffset]?.value;
        const valB = b.cells[colOffset]?.value;
        if (valA === valB) return 0;
        if (valA === null || valA === undefined) return 1;
        if (valB === null || valB === undefined) return -1;
        if (typeof valA === 'number' && typeof valB === 'number') {
          return ascending ? valA - valB : valB - valA;
        }
        return ascending
          ? String(valA).localeCompare(String(valB))
          : String(valB).localeCompare(String(valA));
      });

      // Build values matrix
      const values: CellData[][] = rows.map((r) => r.cells.map((c) => c ?? { value: null }));
      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: startR,
        startColumn: range.startColumn,
        values,
      });
    },
  });

  // 10. AutoFill Sequence / Formula shift
  runtime.registry.registerCommand<AutoFillParams>({
    id: 'sheet.autofill',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const { sourceRange, targetRange } = params;
      const values: CellData[][] = [];

      const sourceRowCount = sourceRange.endRow - sourceRange.startRow + 1;
      const sourceColCount = sourceRange.endColumn - sourceRange.startColumn + 1;

      for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
        const rowList: CellData[] = [];
        const sourceR = sourceRange.startRow + ((r - targetRange.startRow) % sourceRowCount);
        const rowOffset = r - sourceR;

        for (let c = targetRange.startColumn; c <= targetRange.endColumn; c++) {
          const sourceC =
            sourceRange.startColumn + ((c - targetRange.startColumn) % sourceColCount);
          const colOffset = c - sourceC;
          const sourceCell = sheet.cells.get(sourceR, sourceC);

          if (!sourceCell) {
            rowList.push({ value: null });
            continue;
          }

          if (sourceCell.formula) {
            const shifted = shiftFormula(sourceCell.formula, rowOffset, colOffset);
            rowList.push({ ...sourceCell, formula: shifted, value: null });
          } else if (typeof sourceCell.value === 'number') {
            // Sequence extension
            const step = rowOffset !== 0 ? Math.floor((r - sourceRange.endRow) / sourceRowCount) + 1 : 0;
            rowList.push({ ...sourceCell, value: sourceCell.value + step });
          } else {
            rowList.push(structuredClone(sourceCell));
          }
        }
        values.push(rowList);
      }

      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: targetRange.startRow,
        startColumn: targetRange.startColumn,
        values,
      });
    },
  });
}
