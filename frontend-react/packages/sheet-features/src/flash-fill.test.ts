import test from 'node:test';
import assert from 'node:assert/strict';
import { planFlashFill } from './flash-fill';
import { registerSheetCommands } from './index';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

const range = (sheetId: string, startRow: number, endRow: number, column: number) => ({ sheetId, startRow, endRow, startColumn: column, endColumn: column });

test('Flash Fill plans from canonical page cells and writes only blank targets', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('flash-fill-canonical');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    for (const [row, value] of [['Ada Lovelace', 'Ada'], ['Grace Hopper', 'Grace'], ['Katherine Johnson', undefined], ['Mary Jackson', undefined]] as const) {
      await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: ['Ada Lovelace', 'Grace Hopper', 'Katherine Johnson', 'Mary Jackson'].indexOf(row), column: 0, value: { value: row } });
      if (value !== undefined) await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: ['Ada', 'Grace'].indexOf(value), column: 1, value: { value } });
    }
    const plan = planFlashFill(sheet, { sheetId: sheet.id, sourceRange: range(sheet.id, 0, 3, 0), targetRange: range(sheet.id, 0, 3, 1) });
    assert.deepEqual(plan.operation, { kind: 'token', delimiter: ' ', index: 0 });
    assert.deepEqual(plan.writes.map((write) => [write.row, write.column, write.after?.value]), [[2, 1, 'Katherine'], [3, 1, 'Mary']]);
    assert.equal(sheet.cells.get(2, 1), undefined);
  } finally {
    close();
  }
});

test('Flash Fill rejects formulas and overlapping ranges before producing writes', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('flash-fill-rejection');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 0, column: 0, value: { value: null, formula: '=1+1' } });
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 0, value: { value: 'B' } });
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 0, column: 1, value: { value: 'A' } });
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 1, value: { value: 'B' } });
    assert.throws(() => planFlashFill(sheet, { sheetId: sheet.id, sourceRange: range(sheet.id, 0, 1, 0), targetRange: range(sheet.id, 0, 1, 1) }), /formula cells/);
    assert.throws(() => planFlashFill(sheet, { sheetId: sheet.id, sourceRange: range(sheet.id, 0, 1, 0), targetRange: range(sheet.id, 0, 1, 0) }), /must be different/);
  } finally {
    close();
  }
});
