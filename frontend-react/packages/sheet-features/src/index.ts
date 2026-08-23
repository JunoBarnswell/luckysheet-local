import type {
  BandedRule,
  CellData,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  FilterModel,
  FreezeModel,
  MergeSpan,
  RangeRef,
  WorkbookTableModel,
  WorkbookModel,
  WorksheetModel,
  StructuralTransformParams,
  DefinedNameModel,
} from '@react-sheets/core-model';
import { StructuralTransform, normalizeDefinedNameModel } from '@react-sheets/core-model';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { shiftFormula } from './clipboard';
import { registerEditingCommands, rewriteFormulasForSheetRename } from './editing';
import { registerDataToolCommands, normalizeConditionalFormatRule, normalizeDataValidationRule } from './data-features';
import { registerSheetTableCommands } from './sheet-table-commands';
import { registerOutlineCommands } from './outline-commands';

function snapshotCellRegion(
  sheet: WorksheetModel,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): Array<{ row: number; column: number; cell: CellData }> {
  const extracted: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
      extracted.push({ row, column, cell: structuredClone(cell) });
    }
  });
  return extracted;
}

function applyStructuralTransform(workbook: WorkbookModel, params: StructuralTransformParams): void {
  StructuralTransform.apply(workbook, params);
}

export * from './clipboard';
export * from './data-features';
export * from './editing';
export * from './sheet-table-features';
export * from './sheet-table-commands';
export * from './outline-commands';
export * from './outline-features';


export interface SetCellValueParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
}

export interface AddTableParams extends WorkbookTableModel {}

export interface SetRangeValuesParams {
  sheetId: string;
  startRow: number;
  startColumn: number;
  values: CellData[][];
}

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  mode?: 'all' | 'contents' | 'formats' | 'notes' | 'hyperlinks';
}

interface ClearRangeRestoreParams {
  sheetId: string;
  range: RangeRef;
  cells: Array<{ row: number; column: number; value?: CellData }>;
  notes: Array<{ row: number; column: number; note: import('@react-sheets/core-model').CellNote }>;
  comments: import('@react-sheets/core-model').CommentThread[];
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

export interface RenameWorkbookParams { name: string; }

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
  registerEditingCommands(runtime);
  registerDataToolCommands(runtime);
  registerSheetTableCommands(runtime);
  registerOutlineCommands(runtime);

  runtime.registry.registerMutation('workbook.renamed', (item, context) => {
    const params = item.params as RenameWorkbookParams;
    context.workbook.name = params.name;
  });
  runtime.registry.registerCommand<RenameWorkbookParams>({
    id: 'workbook.rename',
    execute: (params, context) => {
      const name = params.name.trim();
      if (!name) throw new Error('Workbook name is required');
      const previous = context.workbook.name;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'workbook.renamed',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.activeSheetId,
        params: { name },
        affectedRanges,
        inverse: [{ id: 'workbook.renamed', unitId: context.workbook.unitId, sheetId: context.workbook.activeSheetId, params: { name: previous }, affectedRanges }],
        apply: () => { context.workbook.name = name; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

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
  runtime.registry.registerMutation('sheet.restore', (item, context) => {
    const restored = item.params as { sheet: WorksheetModel };
    if (!context.workbook.sheets.has(restored.sheet.id)) {
      context.workbook.sheets.set(restored.sheet.id, restored.sheet);
    }
  });

  runtime.registry.registerCommand<{ id: string }>({
    id: 'sheet.remove',
    execute: (paramsInput, context) => {
      const params = paramsInput as { id: string };
      const workbook = context.workbook;
      if (workbook.getSheets().length <= 1) {
        throw new Error('A workbook must keep at least one worksheet');
      }
      const target = workbook.getSheet(params.id);
      const clone = target.cloneSheet();
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.remove',
        unitId: workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.restore',
            unitId: workbook.unitId,
            sheetId: params.id,
            params: { sheet: clone },
            affectedRanges,
          },
        ],
        apply: () => workbook.removeSheet(params.id),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
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
      const formulaRewrites = previousName !== params.name
        ? rewriteFormulasForSheetRename(context.workbook, params.sheetId, previousName, params.name)
        : [];
      if (formulaRewrites.length > 0) {
        for (const item of formulaRewrites) {
          if (item.previous) context.workbook.getSheet(item.sheetId).cells.set(item.row, item.column, item.previous);
        }
      }
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
          ...formulaRewrites.map((item) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: item.sheetId,
            params: { row: item.row, column: item.column, previous: item.previous },
            affectedRanges: [] as RangeRef[],
          })),
        ],
        apply: () => {
          sheet.name = params.name;
          if (previousName !== params.name) {
            rewriteFormulasForSheetRename(context.workbook, params.sheetId, previousName, params.name);
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('table.add', (item, context) => {
    context.workbook.addTable(item.params as WorkbookTableModel);
  });
  runtime.registry.registerMutation('table.remove', (item, context) => {
    context.workbook.removeTable(item.params as string);
  });
  runtime.registry.registerCommand<AddTableParams>({
    id: 'table.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.add',
        unitId: context.workbook.unitId,
        sheetId: params.sourceSheetId ?? context.workbook.activeSheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'table.remove', unitId: context.workbook.unitId, sheetId: params.sourceSheetId ?? context.workbook.activeSheetId, params: params.id, affectedRanges }],
        apply: () => context.workbook.addTable(params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ tableId: string; sheetId: string }>({
    id: 'table.remove',
    execute: (params, context) => {
      const previous = structuredClone(context.workbook.getTable(params.tableId));
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: params.tableId,
        affectedRanges,
        inverse: [{ id: 'table.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => context.workbook.removeTable(params.tableId),
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

  runtime.registry.registerMutation('range.clear', (item, context) => {
    const params = item.params as ClearRangeParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
      for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
        const current = sheet.cells.get(row, column);
        if (params.mode === 'notes') {
          if (current?.note || current?.comment) {
            const next = { ...current };
            delete next.note;
            delete next.comment;
            sheet.cells.set(row, column, next);
          }
          sheet.notes.delete(`${row}:${column}`);
          continue;
        }
        if (params.mode === 'hyperlinks') {
          if (current?.hyperlink !== undefined || current?.hyperlinkDetail !== undefined) {
            const next = { ...current };
            delete next.hyperlink;
            delete next.hyperlinkDetail;
            sheet.cells.set(row, column, next);
          }
          continue;
        }
        if (!current) continue;
        if (params.mode === 'formats') {
          const next = { ...current };
          delete next.style;
          delete next.styleId;
          delete next.numberFormat;
          delete next.displayValue;
          sheet.cells.set(row, column, next);
        } else if (params.mode === 'contents') {
          const next = { ...current };
          next.value = null;
          delete next.formula;
          delete next.displayValue;
          sheet.cells.set(row, column, next);
        } else {
          sheet.cells.delete(row, column);
        }
      }
    }
    if (params.mode === undefined || params.mode === 'all') {
      sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) =>
        thread.row < params.range.startRow
        || thread.row > params.range.endRow
        || thread.column < params.range.startColumn
        || thread.column > params.range.endColumn));
      for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
        for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
          sheet.notes.delete(`${row}:${column}`);
        }
      }
    }
  });

  runtime.registry.registerMutation('range.clear.restore', (item, context) => {
    const params = item.params as ClearRangeRestoreParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
      for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
        sheet.cells.delete(row, column);
        sheet.notes.delete(`${row}:${column}`);
      }
    }
    sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) =>
      thread.row < params.range.startRow
      || thread.row > params.range.endRow
      || thread.column < params.range.startColumn
      || thread.column > params.range.endColumn));
    for (const item of params.cells) {
      if (item.value) sheet.cells.set(item.row, item.column, structuredClone(item.value));
    }
    for (const item of params.notes) sheet.notes.set(`${item.row}:${item.column}`, structuredClone(item.note));
    sheet.commentThreads.push(...structuredClone(params.comments));
  });

  runtime.registry.registerMutation('style.set', (item, context) => {
    const params = item.params as SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string };
    const sheet = context.workbook.getSheet(params.sheetId);
    const ranges = 'range' in params ? [params.range] : params.ranges;
    const style = 'numberFormat' in params && !('style' in params)
      ? { numberFormat: params.numberFormat }
      : (params as SetRangeStyleParams).style;
    for (const range of ranges) {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const current = sheet.cells.get(row, column) ?? { value: null as CellData['value'] };
          const next = { ...current, style: { ...(current.style ?? {}), ...style } };
          if (style.numberFormat !== undefined) next.numberFormat = style.numberFormat;
          sheet.cells.set(row, column, next);
        }
      }
    }
  });

  // 4. Range Clear
  runtime.registry.registerCommand<ClearRangeParams>({
    id: 'sheet.range.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const notes: ClearRangeRestoreParams['notes'] = [];
      const comments: ClearRangeRestoreParams['comments'] = [];
      const range = {
        ...params.range,
        startRow: Math.min(params.range.startRow, params.range.endRow),
        endRow: Math.max(params.range.startRow, params.range.endRow),
        startColumn: Math.min(params.range.startColumn, params.range.endColumn),
        endColumn: Math.max(params.range.startColumn, params.range.endColumn),
      };
      const affectedRanges: RangeRef[] = [range];

      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          previous.push({ row: r, column: c, value: structuredClone(sheet.cells.get(r, c)) });
          const note = sheet.notes.get(`${r}:${c}`);
          if (note) notes.push({ row: r, column: c, note: structuredClone(note) });
        }
      }
      comments.push(...structuredClone(sheet.commentThreads.filter((thread) =>
        thread.row >= range.startRow
        && thread.row <= range.endRow
        && thread.column >= range.startColumn
        && thread.column <= range.endColumn)));

      context.applyMutation({
        id: 'range.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range },
        affectedRanges,
        inverse: [{
          id: 'range.clear.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, range, cells: previous, notes, comments },
          affectedRanges,
        }],
        apply: () => runtime.registry.getMutation('range.clear')({ id: 'range.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, range }, affectedRanges }, context),
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
        apply: () => runtime.registry.getMutation('style.set')({ id: 'style.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }, context),
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
    execute: (params, context) => runtime.execute('data.sort.rows', {
      sheetId: params.sheetId,
      range: params.range,
      criteria: [{ column: params.sortColumn, ascending: params.ascending }],
      hasHeader: params.hasHeader,
    }),
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

  // 11. 结构操作:行/列插入与删除（统一走 StructuralTransform）
  runtime.registry.registerMutation('rows.inserted', (item, context) => {
    const params = item.params as { sheetId: string; at: number; count: number };
    applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count });
  });
  runtime.registry.registerMutation('rows.deleted', (item, context) => {
    const params = item.params as { sheetId: string; at: number; count: number };
    applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count });
  });
  runtime.registry.registerMutation('columns.inserted', (item, context) => {
    const params = item.params as { sheetId: string; at: number; count: number };
    applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count });
  });
  runtime.registry.registerMutation('columns.deleted', (item, context) => {
    const params = item.params as { sheetId: string; at: number; count: number };
    applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count });
  });
  runtime.registry.registerMutation('row.hidden', (item, context) => {
    const params = item.params as { sheetId: string; index: number };
    context.workbook.getSheet(params.sheetId).hiddenRows.add(params.index);
  });
  runtime.registry.registerMutation('row.unhidden', (item, context) => {
    const params = item.params as { sheetId: string; index: number };
    context.workbook.getSheet(params.sheetId).hiddenRows.delete(params.index);
  });
  runtime.registry.registerMutation('rows.unhidden.all', (item, context) => {
    context.workbook.getSheet(item.sheetId).hiddenRows.clear();
  });
  runtime.registry.registerMutation('rows.hidden.restore', (item, context) => {
    const params = item.params as { sheetId: string; indices: number[] };
    const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
    hiddenRows.clear();
    for (const index of params.indices) hiddenRows.add(index);
  });
  runtime.registry.registerMutation('column.hidden', (item, context) => {
    const params = item.params as { sheetId: string; index: number };
    context.workbook.getSheet(params.sheetId).hiddenColumns.add(params.index);
  });
  runtime.registry.registerMutation('column.unhidden', (item, context) => {
    const params = item.params as { sheetId: string; index: number };
    context.workbook.getSheet(params.sheetId).hiddenColumns.delete(params.index);
  });
  runtime.registry.registerMutation('columns.unhidden.all', (item, context) => {
    context.workbook.getSheet(item.sheetId).hiddenColumns.clear();
  });
  runtime.registry.registerMutation('columns.hidden.restore', (item, context) => {
    const params = item.params as { sheetId: string; indices: number[] };
    const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
    hiddenColumns.clear();
    for (const index of params.indices) hiddenColumns.add(index);
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.row.hide',
    execute: (params, context) => {
      const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
      if (hiddenRows.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: params.index, endRow: params.index, startColumn: 0, endColumn: Math.max(0, context.workbook.getSheet(params.sheetId).columnCount - 1) }];
      context.applyMutation({
        id: 'row.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'row.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenRows.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.column.hide',
    execute: (params, context) => {
      const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
      if (hiddenColumns.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, context.workbook.getSheet(params.sheetId).rowCount - 1), startColumn: params.index, endColumn: params.index }];
      context.applyMutation({
        id: 'column.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'column.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenColumns.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.rows.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenRows].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'rows.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'rows.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenRows.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.columns.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenColumns].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'columns.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'columns.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenColumns.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: params.at + params.count - 1,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, params.at, end, 0, Math.max(0, sheet.columnCount - 1));
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: end,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: [] as RangeRef[],
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: params.at + params.count - 1,
      }];
      context.applyMutation({
        id: 'columns.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, 0, Math.max(0, sheet.rowCount - 1), params.at, end);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: end,
      }];
      context.applyMutation({
        id: 'columns.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: [] as RangeRef[],
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 12. 多列排序 / 转置 / 翻转 / 拆分
  runtime.registry.registerCommand<{
    sheetId: string;
    range: RangeRef;
    criteria: Array<{ column: number; ascending: boolean }>;
    hasHeader: boolean;
  }>({
    id: 'sheet.sort.multi',
    execute: (params, context) => runtime.execute('data.sort.rows', params),
  });

  runtime.registry.registerCommand<{ sheetId: string; row: number; column: number; delimiter: string; maxColumns?: number }>({
    id: 'sheet.splitColumn',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const cell = sheet.cells.get(params.row, params.column);
      const text = cell?.value == null ? '' : String(cell.value);
      const maxColumns = Math.max(2, params.maxColumns ?? 4);
      const parts = text.split(params.delimiter).slice(0, maxColumns);
      if (parts.length <= 1 && parts[0] === text) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const baseStyle = cell?.style ? structuredClone(cell.style) : undefined;
      const values: CellData[][] = [parts.map((part) => ({ value: coerceText(part, cell), style: baseStyle ? structuredClone(baseStyle) : undefined }))];
      while (values[0]!.length < 1) values[0]!.push({ value: null });
      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: params.row,
        startColumn: params.column,
        values,
      });
    },
  });

  // 13. 筛选 / 条件格式 / 数据验证 / 色带 / 名称
  runtime.registry.registerMutation('filter.set', (item, context) => {
    const params = item.params as { sheetId: string; filter: FilterModel };
    context.workbook.getSheet(params.sheetId).filter = structuredClone(params.filter);
  });
  runtime.registry.registerMutation('filter.remove', (item, context) => {
    const params = item.params as { sheetId: string };
    const sheet = context.workbook.getSheet(params.sheetId);
    sheet.filter = undefined;
  });
  runtime.registry.registerCommand<{ sheetId: string; filter: FilterModel }>({
    id: 'sheet.filter.set',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).filter;
      const affectedRanges: RangeRef[] = [structuredClone(params.filter.range)];
      context.applyMutation({
        id: 'filter.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: previous ? 'filter.set' : 'filter.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: previous ? { sheetId: params.sheetId, filter: structuredClone(previous) } : { sheetId: params.sheetId },
            affectedRanges,
          },
        ],
        apply: () => {
          context.workbook.getSheet(params.sheetId).filter = structuredClone(params.filter);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.filter.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.filter;
      if (!previous) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [structuredClone(previous.range)];
      context.applyMutation({
        id: 'filter.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'filter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, filter: structuredClone(previous) },
            affectedRanges,
          },
        ],
        apply: () => {
          context.workbook.getSheet(params.sheetId).filter = undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('cf.add', (item, context) => {
    const params = item.params as AddConditionalFormatParams;
    const sheet = context.workbook.getSheet(params.rule.sheetId);
    const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.rule.id);
    if (index >= 0) sheet.conditionalFormats[index] = structuredClone(params.rule);
    else sheet.conditionalFormats.push(structuredClone(params.rule));
  });
  runtime.registry.registerMutation('cf.remove', (item, context) => {
    const params = item.params as { sheetId: string; ruleId: string };
    const sheet = context.workbook.getSheet(params.sheetId);
    const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
    if (index >= 0) sheet.conditionalFormats.splice(index, 1);
  });
  runtime.registry.registerMutation('cf.clear', (item, context) => {
    const params = item.params as { sheetId: string };
    context.workbook.getSheet(params.sheetId).conditionalFormats.length = 0;
  });
  runtime.registry.registerCommand<AddConditionalFormatParams>({
    id: 'sheet.cf.add',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const normalizedRule = normalizeConditionalFormatRule(params.rule, sheet.conditionalFormats.length + 1);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'cf.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'cf.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id },
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(normalizedRule.sheetId);
          const index = target.conditionalFormats.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) target.conditionalFormats[index] = structuredClone(normalizedRule);
          else target.conditionalFormats.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.cf.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.conditionalFormats[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'cf.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'cf.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddConditionalFormatParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.conditionalFormats.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.cf.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (sheet.conditionalFormats.length === 0) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const previous = structuredClone(sheet.conditionalFormats);
      const affectedRanges: RangeRef[] = previous.flatMap((rule) => structuredClone(rule.ranges));
      context.applyMutation({
        id: 'cf.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((rule) => ({
          id: 'cf.add' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, rule } satisfies AddConditionalFormatParams,
          affectedRanges: [] as RangeRef[],
        })),
        apply: () => {
          context.workbook.getSheet(params.sheetId).conditionalFormats.length = 0;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('dv.add', (item, context) => {
    const params = item.params as AddDataValidationParams;
    const sheet = context.workbook.getSheet(params.rule.sheetId);
    const index = sheet.dataValidations.findIndex((rule) => rule.id === params.rule.id);
    if (index >= 0) sheet.dataValidations[index] = structuredClone(params.rule);
    else sheet.dataValidations.push(structuredClone(params.rule));
  });
  runtime.registry.registerMutation('dv.remove', (item, context) => {
    const params = item.params as { sheetId: string; ruleId: string };
    const sheet = context.workbook.getSheet(params.sheetId);
    const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
    if (index >= 0) sheet.dataValidations.splice(index, 1);
  });
  runtime.registry.registerCommand<AddDataValidationParams>({
    id: 'sheet.dv.add',
    execute: (params, context) => {
      const normalizedRule = normalizeDataValidationRule(params.rule);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'dv.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'dv.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(normalizedRule.sheetId);
          const index = sheet.dataValidations.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) sheet.dataValidations[index] = structuredClone(normalizedRule);
          else sheet.dataValidations.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.dv.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.dataValidations[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'dv.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'dv.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddDataValidationParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.dataValidations.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.dataValidations.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('banded.set', (item, context) => {
    const params = item.params as { sheetId: string; rule: BandedRule | null };
    const sheet = context.workbook.getSheet(params.sheetId);
    sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
  });
  runtime.registry.registerCommand<{ sheetId: string; rule: BandedRule | null }>({
    id: 'sheet.banded.set',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).bandedRule;
      const affectedRanges: RangeRef[] = params.rule ? [structuredClone(params.rule.range)] : [];
      context.applyMutation({
        id: 'banded.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'banded.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous ? structuredClone(previous) : null },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(params.sheetId);
          sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('name.set', (item, context) => {
    const params = item.params as { model: DefinedNameModel };
    context.workbook.setDefinedName(params.model);
  });
  runtime.registry.registerMutation('name.remove', (item, context) => {
    const params = item.params as { name: string; scope?: 'workbook' | 'sheet'; sheetId?: string };
    context.workbook.removeDefinedName(params.name, params.scope ?? 'workbook', params.sheetId);
  });
  runtime.registry.registerCommand<{
    name: string;
    value?: string;
    formula?: string;
    scope?: 'workbook' | 'sheet';
    sheetId?: string;
    hidden?: boolean;
    comment?: string;
  }>({
    id: 'workbook.name.set',
    execute: (params, context) => {
      const model: DefinedNameModel = {
        name: params.name,
        formula: params.formula ?? params.value ?? '',
        scope: params.scope ?? 'workbook',
        ...(params.sheetId ? { sheetId: params.sheetId } : {}),
        ...(params.hidden === undefined ? {} : { hidden: params.hidden }),
        ...(params.comment === undefined ? {} : { comment: params.comment }),
      };
      // Validate before opening a mutation so invalid scope/name input cannot
      // create a history entry or leave the legacy formula view half updated.
      const normalized = normalizeDefinedNameModel(model);
      const previous = context.workbook.getDefinedNameExact(normalized.name, normalized.scope, normalized.sheetId);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.set',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.activeSheetId,
        params: { model: normalized },
        affectedRanges,
        inverse: previous !== undefined
          ? [{ id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.activeSheetId, params: { model: previous }, affectedRanges }]
          : [{ id: 'name.remove', unitId: context.workbook.unitId, sheetId: context.workbook.activeSheetId, params: { name: normalized.name, scope: normalized.scope, sheetId: normalized.sheetId }, affectedRanges }],
        apply: () => {
          context.workbook.setDefinedName(normalized);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ name: string; scope?: 'workbook' | 'sheet'; sheetId?: string }>({
    id: 'workbook.name.remove',
    execute: (params, context) => {
      const previous = context.workbook.getDefinedNameExact(params.name, params.scope ?? 'workbook', params.sheetId);
      if (previous === undefined || previous.scope !== (params.scope ?? 'workbook') || previous.sheetId !== params.sheetId) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.remove',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.activeSheetId,
        params,
        affectedRanges,
        inverse: [
          { id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.activeSheetId, params: { model: previous }, affectedRanges },
        ],
        apply: () => {
          context.workbook.removeDefinedName(params.name, previous.scope, previous.sheetId);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}

function coerceText(text: string, previousCell: CellData | undefined): CellData['value'] {
  if (typeof previousCell?.value === 'number') {
    const numeric = Number(text.replace(/[$,%]/g, ''));
    if (Number.isFinite(numeric) && text.trim() !== '') return numeric;
  }
  return text;
}
