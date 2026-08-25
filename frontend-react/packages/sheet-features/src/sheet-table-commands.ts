import type { AutoFilterModel, RangeRef, SheetTableModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import {
  planTotalRowToggle,
  snapshotTotalRowCells,
  validateFilterOwnership,
  validateSheetTableModel,
} from './sheet-table-features';

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && Number.isInteger(value.startRow) && Number.isInteger(value.endRow)
    && Number.isInteger(value.startColumn) && Number.isInteger(value.endColumn)
    && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
    && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
}

function isSheetTable(value: unknown): value is SheetTableModel {
  if (!isRecord(value) || typeof value.id !== 'string' || typeof value.sheetId !== 'string' || typeof value.name !== 'string' || !isRange(value.range)) return false;
  return typeof value.hasHeaderRow === 'boolean'
    && typeof value.hasTotalRow === 'boolean'
    && typeof value.showBandedRows === 'boolean'
    && typeof value.showBandedColumns === 'boolean'
    && typeof value.showFirstColumn === 'boolean'
    && typeof value.showLastColumn === 'boolean'
    && typeof value.showFilterButton === 'boolean'
    && ['none', 'rows', 'columns', 'both'].includes(String(value.autoExpand))
    && Array.isArray(value.columns)
    && value.columns.every((column) => isRecord(column) && typeof column.id === 'string' && typeof column.name === 'string');
}

function isSheetTableRemove(value: unknown): value is { sheetId: string; tableId: string; range: RangeRef } {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.tableId === 'string' && isRange(value.range);
}

function tableRange(value: SheetTableModel): RangeRef[] {
  return [structuredClone(value.range)];
}

interface TableAutoFilterParams {
  sheetId: string;
  tableId: string;
  autoFilter?: AutoFilterModel;
}

function isTableAutoFilter(value: unknown): value is TableAutoFilterParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && typeof value.tableId === 'string'
    && (value.autoFilter === undefined || (isRecord(value.autoFilter) && isRange(value.autoFilter.range) && isRecord(value.autoFilter.columns)));
}

function removedTableRange(value: { range: RangeRef }): RangeRef[] {
  return [structuredClone(value.range)];
}

export interface AddSheetTableParams extends SheetTableModel {}

export function registerSheetTableCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<AddSheetTableParams>({
    id: 'sheetTable.add',
    handler: (item, context) => {
      if (!isSheetTable(item.params)) throw new Error('Invalid sheetTable.add mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const table = validateSheetTableModel(item.params, sheet);
      if (sheet.sheetTables.some((entry) => entry.id === table.id || entry.name.toLocaleLowerCase() === table.name.toLocaleLowerCase())) {
        throw new Error(`Sheet Table already exists: ${table.name}`);
      }
      if (sheet.sheetTables.some((entry) => entry.range.startRow <= table.range.endRow
        && entry.range.endRow >= table.range.startRow && entry.range.startColumn <= table.range.endColumn
        && entry.range.endColumn >= table.range.startColumn)) throw new Error('Sheet Tables cannot overlap');
      sheet.sheetTables.push(structuredClone(table));
    },
    metadata: {
      schema: { name: 'SheetTableModel', validate: isSheetTable },
      permission: { capability: 'sheet.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: tableRange, mode: 'exact' },
      inverseIds: ['sheetTable.remove'],
    },
  });
  runtime.registry.registerMutation<TableAutoFilterParams>({
    id: 'sheetTable.autoFilter.set',
    handler: (item, context) => {
      if (!isTableAutoFilter(item.params)) throw new Error('Invalid sheetTable.autoFilter.set mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const table = sheet.sheetTables.find((entry) => entry.id === item.params.tableId);
      if (!table) throw new Error(`Sheet Table not found: ${item.params.tableId}`);
      table.autoFilter = item.params.autoFilter
        ? validateFilterOwnership(sheet, item.params.autoFilter, { kind: 'table', tableId: table.id })
        : undefined;
    },
    metadata: {
      schema: { name: 'SheetTableAutoFilterSet', validate: isTableAutoFilter },
      permission: { capability: 'sheet.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.autoFilter ? [structuredClone(params.autoFilter.range)] : [], mode: 'declared' },
      inverseIds: ['sheetTable.autoFilter.set'],
    },
  });
  runtime.registry.registerMutation({
    id: 'sheetTable.remove',
    handler: (item, context) => {
      if (!isSheetTableRemove(item.params)) throw new Error('Invalid sheetTable.remove mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
      if (index >= 0) sheet.sheetTables.splice(index, 1);
    },
    metadata: {
      schema: { name: 'SheetTableRemove', validate: isSheetTableRemove },
      permission: { capability: 'sheet.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: removedTableRange, mode: 'exact' },
      inverseIds: ['sheetTable.add'],
    },
  });
  runtime.registry.registerMutation({
    id: 'sheetTable.update',
    handler: (item, context) => {
      if (!isSheetTable(item.params)) throw new Error('Invalid sheetTable.update mutation payload');
      const sheet = context.workbook.getSheet(item.params.sheetId);
      const table = validateSheetTableModel(item.params, sheet);
      const index = sheet.sheetTables.findIndex((entry) => entry.id === table.id);
      if (index < 0) throw new Error(`Sheet Table not found: ${table.id}`);
      if (sheet.sheetTables.some((entry) => entry.id !== table.id
        && entry.range.startRow <= table.range.endRow && entry.range.endRow >= table.range.startRow
        && entry.range.startColumn <= table.range.endColumn && entry.range.endColumn >= table.range.startColumn)) {
        throw new Error('Sheet Tables cannot overlap');
      }
      sheet.sheetTables[index] = structuredClone(table);
    },
    metadata: {
      schema: { name: 'SheetTableModel', validate: isSheetTable },
      permission: { capability: 'sheet.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: tableRange, mode: 'exact' },
      inverseIds: ['sheetTable.update'],
    },
  });

  runtime.registry.registerCommand<AddSheetTableParams>({
    id: 'sheetTable.add',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const table = validateSheetTableModel(params, sheet);
      if (sheet.sheetTables.some((entry) => entry.id === table.id || entry.name.toLocaleLowerCase() === table.name.toLocaleLowerCase())) {
        throw new Error(`Sheet Table already exists: ${table.name}`);
      }
      const intersects = (left: RangeRef, right: RangeRef): boolean => left.startRow <= right.endRow
        && left.endRow >= right.startRow && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
      if (sheet.sheetTables.some((entry) => intersects(entry.range, table.range))) throw new Error('Sheet Tables cannot overlap');
      const affectedRanges: RangeRef[] = [structuredClone(table.range)];
      context.applyMutation({
        id: 'sheetTable.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: table,
        affectedRanges,
        inverse: [{ id: 'sheetTable.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, tableId: params.id, range: params.range }, affectedRanges }],
        apply: () => {
          context.workbook.getSheet(params.sheetId).sheetTables.push(structuredClone(table));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<AddSheetTableParams>({
    id: 'sheetTable.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.sheetTables.findIndex((entry) => entry.id === params.id);
      if (index < 0) throw new Error(`Sheet Table not found: ${params.id}`);
      const previous = structuredClone(sheet.sheetTables[index]!);
      const next = validateSheetTableModel(params, sheet);
      const overlaps = sheet.sheetTables.some((entry) => entry.id !== next.id
        && entry.range.startRow <= next.range.endRow && entry.range.endRow >= next.range.startRow
        && entry.range.startColumn <= next.range.endColumn && entry.range.endColumn >= next.range.startColumn);
      if (overlaps) throw new Error('Sheet Tables cannot overlap');
      const affectedRanges = [structuredClone(next.range)];
      context.applyMutation({
        id: 'sheetTable.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: next,
        affectedRanges,
        inverse: [{ id: 'sheetTable.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges: [structuredClone(previous.range)] }],
        apply: () => { sheet.sheetTables[index] = structuredClone(next); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<TableAutoFilterParams>({
    id: 'sheetTable.autoFilter.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const table = sheet.sheetTables.find((entry) => entry.id === params.tableId);
      if (!table) throw new Error(`Sheet Table not found: ${params.tableId}`);
      const previous = table.autoFilter ? structuredClone(table.autoFilter) : undefined;
      const next = params.autoFilter
        ? validateFilterOwnership(sheet, params.autoFilter, { kind: 'table', tableId: table.id })
        : undefined;
      const affectedRanges = [structuredClone(table.range)];
      const mutationParams = { ...params, autoFilter: next };
      context.applyMutation({
        id: 'sheetTable.autoFilter.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: mutationParams,
        affectedRanges,
        inverse: [{
          id: 'sheetTable.autoFilter.set',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, tableId: params.tableId, autoFilter: previous },
          affectedRanges,
        }],
        apply: () => {
          table.autoFilter = next ? structuredClone(next) : undefined;
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
        params: { ...params, range: previous.range },
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
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.sheetTables[index]!);
      const affectedRanges = [structuredClone(previous.range)];
      context.applyMutation({
        id: 'sheetTable.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range: previous.range },
        affectedRanges,
        inverse: [{ id: 'sheetTable.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => {
          const target = sheet.sheetTables.findIndex((table) => table.id === params.tableId);
          if (target >= 0) sheet.sheetTables.splice(target, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
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
      const linkedTable = sheet.sheetTables.find((table) => table.id === previousTable.id && table.autoFilter);
      const previousFilter = linkedTable?.autoFilter ? structuredClone(linkedTable.autoFilter) : sheet.autoFilter ? structuredClone(sheet.autoFilter) : undefined;
      let mutationCount = 0;

      // A total row is a real worksheet row. Insert/delete it through the
      // structural command so data immediately below the table is shifted and
      // all formula/object/range participants are updated atomically. The old
      // implementation merely expanded the table range and overwrote the
      // next row, which was silent data loss.
      if (params.enabled) {
        const result = runtime.execute('sheet.rows.insert', {
          sheetId: params.sheetId,
          at: plan.totalRow,
          count: 1,
        });
        mutationCount += result.mutationCount;
      } else {
        const result = runtime.execute('sheet.rows.delete', {
          sheetId: params.sheetId,
          at: plan.totalRow,
          count: 1,
        });
        mutationCount += result.mutationCount;
      }

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
        // The row deletion above already removed the total-row cells. Keep
        // the snapshot solely as an integrity check for the command plan; no
        // second clear/write is allowed to create a stale-cell fallback.
        void cellSnapshots;
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
            params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.previous },
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
        if (linkedTable) {
          context.applyMutation({
            id: 'sheetTable.autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, tableId: linkedTable.id, autoFilter: nextFilter },
            affectedRanges: [structuredClone(plan.nextTable.range)],
            inverse: [{
              id: 'sheetTable.autoFilter.set',
              unitId: context.workbook.unitId,
              sheetId: params.sheetId,
              params: { sheetId: params.sheetId, tableId: linkedTable.id, autoFilter: previousFilter },
              affectedRanges: [structuredClone(previousFilter.range)],
            }],
            apply: () => {
              const replacement = sheet.sheetTables.find((candidate) => candidate.id === linkedTable.id);
              if (replacement) replacement.autoFilter = structuredClone(nextFilter);
            },
          });
        } else {
          context.applyMutation({
            id: 'autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, autoFilter: nextFilter },
            affectedRanges: [structuredClone(plan.nextTable.range)],
            inverse: [{
              id: 'autoFilter.set',
              unitId: context.workbook.unitId,
              sheetId: params.sheetId,
              params: { sheetId: params.sheetId, autoFilter: previousFilter },
              affectedRanges: [structuredClone(previousFilter.range)],
            }],
            apply: () => {
              sheet.autoFilter = structuredClone(nextFilter);
            },
          });
        }
        mutationCount += 1;
      }

      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
}
