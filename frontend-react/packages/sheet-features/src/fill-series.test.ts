import test from 'node:test';
import assert from 'node:assert/strict';
import { planFill } from './fill-series';
import { registerSheetCommands } from './index';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

const range = (sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number) => ({ sheetId, startRow, endRow, startColumn, endColumn });

test('series planner reads canonical page values, preserves seeds, and emits typed writes', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('fill-series-canonical');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 0, column: 0, value: { value: 1 } });
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 0, value: { value: 3 } });
    const plan = planFill(sheet, {
      sheetId: sheet.id,
      sourceRange: range(sheet.id, 0, 1, 0, 0),
      targetRange: range(sheet.id, 0, 4, 0, 0),
      direction: 'down', mode: 'series',
    });
    assert.deepEqual(plan.writes.map((write) => [write.row, write.column, write.cell?.value]), [[2, 0, 5], [3, 0, 7], [4, 0, 9]]);
    assert.equal(sheet.cells.get(0, 0)?.value, 1);
    assert.equal(sheet.cells.get(1, 0)?.value, 3);
  } finally {
    close();
  }
});

test('series planner rejects malformed geometry and non-numeric seeds before producing writes', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('fill-series-rejection');
  try {
    registerSheetCommands(runtime);
    const sheet = workbook.getSheet('sheet-1');
    await runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 0, column: 0, value: { value: 'seed' } });
    assert.throws(() => planFill(sheet, {
      sheetId: sheet.id,
      sourceRange: range(sheet.id, 0, 0, 0, 0),
      targetRange: range(sheet.id, 0, 2, 0, 0),
      direction: 'down', mode: 'series',
    }), /numeric|seed|series/i);
    assert.throws(() => planFill(sheet, {
      sheetId: sheet.id,
      sourceRange: range('other-sheet', 0, 0, 0, 0),
      targetRange: range(sheet.id, 0, 2, 0, 0),
      direction: 'down', mode: 'series',
    }), /sheet|range/i);
  } finally {
    close();
  }
});
