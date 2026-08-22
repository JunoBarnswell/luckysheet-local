import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from './index';

test('CommandRuntime executes a registered command and tracks history', () => {
  const workbook = new WorkbookModel('unit-1', 'Runtime');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation('cell.set', (item, context) => {
    const params = item.params as { row: number; column: number; value: string };
    context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
  });
  runtime.registry.registerMutation('cell.restore', (item, context) => {
    const params = item.params as { row: number; column: number; previous?: { value: string } };
    if (params.previous) context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, params.previous);
    else context.workbook.getSheet(item.sheetId).cells.delete(params.row, params.column);
  });
  runtime.registry.registerCommand({
    id: 'cell.set',
    execute: (params: { row: number; column: number; value: string }, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      const previous = sheet.cells.get(params.row, params.column);
      const range = [{ sheetId: 'sheet-1', startRow: params.row, endRow: params.row, startColumn: params.column, endColumn: params.column }];
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params,
        affectedRanges: range,
        inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: 'sheet-1', params: { row: params.row, column: params.column, previous }, affectedRanges: range }],
        apply: () => sheet.cells.set(params.row, params.column, { value: params.value }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: range };
    },
  });
  const result = runtime.execute('cell.set', { row: 1, column: 1, value: 'A' });
  assert.equal(result.mutationCount, 1);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.equal(runtime.undo(), true);
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1), undefined);
  assert.equal(runtime.redo(), true);
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1)?.value, 'A');
});
