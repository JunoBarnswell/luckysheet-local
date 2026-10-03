import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { WorkbookSession } from '@react-sheets/spreadsheet-app';
import { getWorkbookObjectPort } from '../../../spreadsheet-app/src/workbook-object-port';
import { subtotalFunctionNumber } from '@react-sheets/formula-engine';
import { SdkError } from '../error';
import { parseCellAddress } from '../workbook/cell-address';
import { MAX_OBJECT_RANGE_CELLS } from '../workbook/value';
import type { DataActions, DataActionResult, DataRangeAddress } from './contract';

/** Owns selection-derived plans. Dispatch remains the single transaction boundary. */
export class DataDomain {
  private disposed = false;
  private readonly releaseLifetime: () => void;
  readonly actions: DataActions;
  constructor(private readonly session: WorkbookSession) {
    this.releaseLifetime = getWorkbookObjectPort(session).subscribeDisposed(() => this.dispose());
    this.actions = this.actionsFor();
  }
  /** Bind a consumer lifetime to this same domain, without another plan or state. */
  actionsFor(assertCurrent?: () => void): DataActions {
    const execute = (operation: string, plan: () => CommandDescriptor) => this.execute(operation, plan, assertCurrent);
    return Object.freeze({
      quickSort: (ascending: boolean) => execute('data.quickSort', () => {
        const context = this.session.getDataRegionContext();
        if (typeof ascending !== 'boolean' || context.range.endRow <= context.range.startRow) throw this.invalid('data.quickSort', '排序需要多行数据及有效顺序。');
        return { commandId: 'data.sort.quick', params: { sheetId: context.range.sheetId, range: context.range,
          sortColumn: context.activeColumn, ascending, hasHeader: context.header.kind === 'present', dataRegionContext: context } };
      }),
      toggleFilter: () => execute('data.toggleFilter', () => this.filterPlan(false)),
      clearFilter: () => execute('data.clearFilter', () => this.filterPlan(true)),
      textToColumns: (input: Parameters<DataActions['textToColumns']>[0] = {}) => execute('data.textToColumns', () => {
        const range = this.range();
        const column = this.session.getSelection().activeCell.column;
        const available = this.session.getSelectedSheet().columnCount - column;
        return { commandId: 'data.textToColumns', params: { sheetId: range.sheetId,
          range: { ...range, startColumn: column, endColumn: column }, delimiter: input.delimiter ?? ',', maxColumns: input.maxColumns ?? Math.min(8, available) } };
      }),
      removeDuplicates: (input: Parameters<DataActions['removeDuplicates']>[0] = {}) => execute('data.removeDuplicates', () => {
        const range = this.range();
        if (range.endRow <= range.startRow) throw this.invalid('data.removeDuplicates', '删除重复项需要多行数据。');
        const columns = input.columns ? [...input.columns] : Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, index) => range.startColumn + index);
        return { commandId: 'data.removeDuplicates', params: { sheetId: range.sheetId, range, columns, hasHeader: input.hasHeader ?? true } };
      }),
      subtotal: (input: Parameters<DataActions['subtotal']>[0] = {}) => execute('data.subtotal', () => {
        if (!input || typeof input !== 'object' || Array.isArray(input)
          || Object.keys(input).some(key => !['range', 'groupColumn', 'valueColumn', 'functionName', 'excludeHiddenRows'].includes(key))) throw this.invalid('data.subtotal', '小计选项无效。');
        const range = this.range(input.range);
        if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) throw this.invalid('data.subtotal', '分类汇总需要至少两列及多行数据。');
        const functionName = input.functionName === undefined ? 'SUM' : input.functionName;
        const groupColumn = input.groupColumn === undefined ? range.startColumn : input.groupColumn;
        const valueColumn = input.valueColumn === undefined ? range.startColumn + 1 : input.valueColumn;
        if (subtotalFunctionNumber(functionName) === undefined || [groupColumn, valueColumn].some(column => !Number.isSafeInteger(column) || column < range.startColumn || column > range.endColumn)
          || input.excludeHiddenRows !== undefined && typeof input.excludeHiddenRows !== 'boolean') throw this.invalid('data.subtotal', '小计函数、列或隐藏行选项无效。');
        return { commandId: 'data.subtotal', params: { sheetId: range.sheetId, range,
          groupColumn, valueColumn, functionName, ...(input.excludeHiddenRows === undefined ? {} : { excludeHiddenRows: input.excludeHiddenRows }) } };
      }),
    });
  }
  private range(target?: DataRangeAddress) {
    if (target !== undefined) {
      try {
        if (!target || typeof target !== 'object' || typeof target.sheetId !== 'string' || !target.sheetId
          || typeof target.address !== 'string' || Object.keys(target).some(key => !['sheetId', 'address'].includes(key))) throw new Error('Explicit worksheet and address are required.');
        const parts = target.address.split(':');
        if (parts.length > 2) throw new Error('A single rectangular address is required.');
        const start = parseCellAddress(parts[0]!), end = parseCellAddress(parts[1] ?? parts[0]!);
        const sheet = getWorkbookObjectPort(this.session).readWorksheet(target.sheetId);
        if (sheet.kind !== 'worksheet' || end.row < start.row || end.column < start.column || end.row >= sheet.rowCount || end.column >= sheet.columnCount
          || (end.row - start.row + 1) * (end.column - start.column + 1) > MAX_OBJECT_RANGE_CELLS) throw new Error('Range exceeds the canonical worksheet or operation budget.');
        return { sheetId: sheet.id, startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column };
      } catch (cause) {
        throw new SdkError('INVALID_ARGUMENT', 'data.subtotal', cause instanceof Error ? cause.message : '区域无效。', '请提供当前工作簿内有效工作表和矩形地址。', { cause, object: { workbookId: getWorkbookObjectPort(this.session).unitId } });
      }
    }
    const selected = this.session.getPrimaryRange();
    return selected.startRow !== selected.endRow || selected.startColumn !== selected.endColumn
      ? { ...selected } : this.session.getCurrentRegion();
  }
  private filterPlan(clear: boolean): CommandDescriptor {
    const context = this.session.getDataRegionContext();
    if (!clear && context.owner.kind === 'sheet-table') {
      const tableId = context.owner.tableId;
      const table = this.session.getSelectedSheet().sheetTables.find(candidate => candidate.id === tableId);
      if (!table) throw new SdkError('CONTRACT_INVALID', 'data.toggleFilter', '筛选所属表格不存在。', '请重新加载工作簿。');
      return { commandId: 'sheetTable.update', params: { ...structuredClone(table), showFilterButton: !table.showFilterButton, autoFilter: undefined } };
    }
    return { commandId: clear ? 'sheet.autoFilter.clearCriteria' : 'sheet.autoFilter.toggle',
      params: { sheetId: context.range.sheetId, range: context.range, dataRegionContext: context } };
  }
  private invalid(operation: string, message: string): SdkError {
    return new SdkError('INVALID_ARGUMENT', operation, message, '请检查所选区域和参数后重试。', { object: { workbookId: getWorkbookObjectPort(this.session).unitId } });
  }
  private async execute(operation: string, plan: () => CommandDescriptor, assertCurrent?: () => void): Promise<DataActionResult> {
    try {
      assertCurrent?.();
      if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', operation, '数据会话已关闭。', '请重新打开工作簿。');
      const result = await this.session.dispatch(plan());
      assertCurrent?.();
      if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', operation, '数据会话已关闭。', '请重新打开工作簿。');
      if (result.status === 'committed') return { status: 'applied' };
      const code = result.error.code === 'PERMISSION_DENIED' ? 'FORBIDDEN' : 'REQUEST_REJECTED';
      return { status: 'rejected', error: new SdkError(code, operation, result.error.message, '请检查权限、数据来源及工作簿状态后重试。', { cause: result.error, object: { workbookId: getWorkbookObjectPort(this.session).unitId } }) };
    } catch (cause) {
      const error = cause instanceof SdkError && cause.operation === operation && cause.object ? cause : new SdkError(cause instanceof SdkError ? cause.code : 'REQUEST_REJECTED', operation,
        cause instanceof Error ? cause.message : '数据操作失败。', cause instanceof SdkError ? cause.recovery : '请检查数据来源及工作簿状态后重试。',
        { cause, ...(cause instanceof SdkError && cause.status !== undefined ? { status: cause.status } : {}), object: { workbookId: getWorkbookObjectPort(this.session).unitId } });
      if (!this.disposed && error.code !== 'RUNTIME_DISPOSED') this.session.notify(error.message);
      return { status: 'rejected', error };
    }
  }
  dispose(): void { if (this.disposed) return; this.disposed = true; this.releaseLifetime(); }
}
