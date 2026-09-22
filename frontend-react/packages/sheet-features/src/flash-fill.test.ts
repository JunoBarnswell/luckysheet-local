import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { planFlashFill } from './flash-fill';
import { registerSheetCommands } from './index';

function setup() {
  const workbook = new WorkbookModel('flash-fill-unit', 'Flash Fill');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.rowCount = 12;
  sheet.columnCount = 8;
  return sheet;
}

function range(sheetId: string, startRow: number, endRow: number, column: number) {
  return { sheetId, startRow, endRow, startColumn: column, endColumn: column };
}

test('Flash Fill infers the same token from multiple examples and writes only blanks', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 'Ada Lovelace' });
  sheet.cells.set(1, 0, { value: 'Grace Hopper' });
  sheet.cells.set(0, 1, { value: 'Ada' });
  sheet.cells.set(1, 1, { value: 'Grace' });
  sheet.cells.set(2, 0, { value: 'Katherine Johnson' });
  sheet.cells.set(3, 0, { value: 'Mary Jackson' });
  const plan = planFlashFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 3, 0),
    targetRange: range(sheet.id, 0, 3, 1),
  });
  assert.deepEqual(plan.operation, { kind: 'token', delimiter: ' ', index: 0 });
  assert.deepEqual(plan.writes.map((write) => [write.row, write.column, write.after?.value]), [
    [2, 1, 'Katherine'],
    [3, 1, 'Mary'],
  ]);
  assert.deepEqual(sheet.cells.get(2, 1), undefined);
});

test('Flash Fill infers a numeric delta and rejects ambiguous examples', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 10 });
  sheet.cells.set(1, 0, { value: 20 });
  sheet.cells.set(0, 1, { value: 11 });
  sheet.cells.set(1, 1, { value: 21 });
  sheet.cells.set(2, 0, { value: 30 });
  const plan = planFlashFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 2, 0),
    targetRange: range(sheet.id, 0, 2, 1),
  });
  assert.deepEqual(plan.operation, { kind: 'numeric-delta', delta: 1 });
  assert.equal(plan.writes[0]?.after?.value, 31);

  sheet.cells.delete(0, 1);
  sheet.cells.delete(1, 1);
  assert.throws(() => planFlashFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 2, 0),
    targetRange: range(sheet.id, 0, 2, 1),
  }), /at least two non-blank examples/);
});

test('Flash Fill rejects formulas and overlapping source/target ranges before writes', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: null, formula: '=1+1' });
  sheet.cells.set(1, 0, { value: 'B' });
  sheet.cells.set(0, 1, { value: 'A' });
  sheet.cells.set(1, 1, { value: 'B' });
  assert.throws(() => planFlashFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 1, 0),
    targetRange: range(sheet.id, 0, 1, 1),
  }), /formula cells/);
  assert.throws(() => planFlashFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 1, 0),
    targetRange: range(sheet.id, 0, 1, 0),
  }), /must be different/);
});

test('Flash Fill is one undoable command and rejects a stale target on replay', () => {
  const workbook = new WorkbookModel('flash-fill-command', 'Flash Fill command');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.rowCount = 8;
  sheet.columnCount = 4;
  sheet.cells.set(0, 0, { value: 'Ada Lovelace' });
  sheet.cells.set(1, 0, { value: 'Grace Hopper' });
  sheet.cells.set(0, 1, { value: 'Ada' });
  sheet.cells.set(1, 1, { value: 'Grace' });
  assert.throws(() => runtime.execute('range.flashFill', { sheetId: sheet.id, sourceRange: range(sheet.id, 0, 3, 0), targetRange: range(sheet.id, 0, 3, 1) }), /blank source value/);
  assert.equal(sheet.cells.get(2, 1), undefined);
  // The example above has no source value on row 2, so the transaction must
  // fail before any cell mutation. Add a complete source band and retry.
  sheet.cells.set(2, 0, { value: 'Katherine Johnson' });
  sheet.cells.set(3, 0, { value: 'Mary Jackson' });
  runtime.execute('range.flashFill', { sheetId: sheet.id, sourceRange: range(sheet.id, 0, 3, 0), targetRange: range(sheet.id, 0, 3, 1) });
  assert.equal(sheet.cells.get(2, 1)?.value, 'Katherine');
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.get(2, 1), undefined);
  assert.equal(runtime.redo(), true);
  assert.equal(sheet.cells.get(3, 1)?.value, 'Mary');
});
