import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMatrix, WorkbookModel } from './index';

test('CellMatrix keeps empty logical space sparse', () => {
  const matrix = new CellMatrix();
  matrix.set(100_000, 4, { value: 'tail' });
  assert.equal(matrix.get(0, 0), undefined);
  assert.equal(matrix.get(100_000, 4)?.value, 'tail');
  assert.deepEqual(Object.keys(matrix.toJSON()), ['100000']);
});

test('WorkbookSnapshotV1 round-trips model state', () => {
  const workbook = new WorkbookModel('unit-1', 'Research');
  workbook.getSheet('sheet-1').cells.set(1, 2, { value: 42, formula: '=40+2' });
  const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
  assert.equal(restored.getSheet('sheet-1').cells.get(1, 2)?.formula, '=40+2');
});
