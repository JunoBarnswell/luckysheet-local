import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { WorkbookSession } from '@react-sheets/spreadsheet-app';
import { SdkError } from '../error';
import type { DataActions, DataActionResult } from './contract';

/** Owns selection-derived plans. Dispatch remains the single transaction boundary. */
export class DataDomain {
  private disposed = false;
  readonly actions: DataActions;
  constructor(private readonly session: WorkbookSession) {
    this.actions = Object.freeze({
      quickSort: (ascending: boolean) => this.execute('data.quickSort', () => {
        const context = this.session.getDataRegionContext();
        if (typeof ascending !== 'boolean' || context.range.endRow <= context.range.startRow) throw this.invalid('data.quickSort', '排序需要多行数据及有效顺序。');
        return { commandId: 'data.sort.quick', params: { sheetId: context.range.sheetId, range: context.range,
          sortColumn: context.activeColumn, ascending, hasHeader: context.header.kind === 'present', dataRegionContext: context } };
      }),
      toggleFilter: () => this.execute('data.toggleFilter', () => this.filterPlan(false)),
      clearFilter: () => this.execute('data.clearFilter', () => this.filterPlan(true)),
      textToColumns: (input: Parameters<DataActions['textToColumns']>[0] = {}) => this.execute('data.textToColumns', () => {
        const range = this.range();
        const column = this.session.getSelection().activeCell.column;
        const available = this.session.getSelectedSheet().columnCount - column;
        return { commandId: 'data.textToColumns', params: { sheetId: range.sheetId,
          range: { ...range, startColumn: column, endColumn: column }, delimiter: input.delimiter ?? ',', maxColumns: input.maxColumns ?? Math.min(8, available) } };
      }),
      removeDuplicates: (input: Parameters<DataActions['removeDuplicates']>[0] = {}) => this.execute('data.removeDuplicates', () => {
        const range = this.range();
        if (range.endRow <= range.startRow) throw this.invalid('data.removeDuplicates', '删除重复项需要多行数据。');
        const columns = input.columns ? [...input.columns] : Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, index) => range.startColumn + index);
        return { commandId: 'data.removeDuplicates', params: { sheetId: range.sheetId, range, columns, hasHeader: input.hasHeader ?? true } };
      }),
      subtotal: (input: Parameters<DataActions['subtotal']>[0] = {}) => this.execute('data.subtotal', () => {
        const range = this.range();
        if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) throw this.invalid('data.subtotal', '分类汇总需要至少两列及多行数据。');
        return { commandId: 'data.subtotal', params: { sheetId: range.sheetId, range,
          groupColumn: input.groupColumn ?? range.startColumn, valueColumn: input.valueColumn ?? range.startColumn + 1, functionName: input.functionName ?? 'SUM' } };
      }),
    });
  }
  private range() {
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
    return new SdkError('INVALID_ARGUMENT', operation, message, '请检查所选区域和参数后重试。');
  }
  private async execute(operation: string, plan: () => CommandDescriptor): Promise<DataActionResult> {
    try {
      if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', operation, '数据会话已关闭。', '请重新打开工作簿。');
      const result = await this.session.dispatch(plan());
      if (result.status === 'committed') return { status: 'applied' };
      const code = result.error.code === 'PERMISSION_DENIED' ? 'FORBIDDEN' : 'REQUEST_REJECTED';
      return { status: 'rejected', error: new SdkError(code, operation, result.error.message, '请检查权限、数据来源及工作簿状态后重试。', { cause: result.error }) };
    } catch (cause) {
      const error = cause instanceof SdkError ? cause : new SdkError('REQUEST_REJECTED', operation,
        cause instanceof Error ? cause.message : '数据操作失败。', '请检查数据来源及工作簿状态后重试。', { cause });
      if (!this.disposed) this.session.notify(error.message);
      return { status: 'rejected', error };
    }
  }
  dispose(): void { this.disposed = true; }
}
