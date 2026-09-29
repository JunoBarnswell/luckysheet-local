import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { MAX_COLUMN_INDEX, MAX_ROW_INDEX, RangeIndex } from '@react-sheets/formula-engine';
import { StructuralTransform, WorkbookModel, type StructuralTransformParams } from './index';

function apply(workbook: WorkbookModel, params: StructuralTransformParams) {
  const index = new RangeIndex(workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name })));
  return StructuralTransform.apply(workbook, params, index);
}

describe('planned structural cell relocations', () => {
  for (const axis of ['row', 'column'] as const) {
    for (const count of [1, 3]) {
      it(`inserts ${count} ${axis}(s) at an occupied tail without growing the extent twice`, () => {
        const workbook = new WorkbookModel(`extent-${axis}-${count}`, 'Extent ownership');
        const sheet = workbook.getSheet('sheet-1');
        sheet.rowCount = 8;
        sheet.columnCount = 8;
        const row = axis === 'row' ? 7 : 1;
        const column = axis === 'column' ? 7 : 1;
        sheet.cells.set(row, column, { value: 'tail' });
        const before = workbook.snapshot();

        apply(workbook, { kind: axis === 'row' ? 'insert-rows' : 'insert-columns', sheetId: sheet.id, at: 2, count });

        assert.equal(sheet.rowCount, axis === 'row' ? 8 + count : 8);
        assert.equal(sheet.columnCount, axis === 'column' ? 8 + count : 8);
        assert.equal(sheet.cells.get(axis === 'row' ? row + count : row, axis === 'column' ? column + count : column)?.value, 'tail');
        apply(workbook, { kind: axis === 'row' ? 'delete-rows' : 'delete-columns', sheetId: sheet.id, at: 2, count });
        assert.deepEqual(workbook.snapshot(), before);
      });
    }

    it(`allows the last legal ${axis} insertion and rejects an actual overflow without partial changes`, () => {
      const workbook = new WorkbookModel(`extent-limit-${axis}`, 'Extent limit');
      const sheet = workbook.getSheet('sheet-1');
      const maximum = axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
      if (axis === 'row') sheet.rowCount = maximum;
      else sheet.columnCount = maximum;
      const row = axis === 'row' ? maximum - 1 : 0;
      const column = axis === 'column' ? maximum - 1 : 0;
      sheet.cells.set(row, column, { value: 'last legal cell' });
      const params = { kind: axis === 'row' ? 'insert-rows' : 'insert-columns', sheetId: sheet.id, at: 0, count: 1 } as const;

      apply(workbook, params);

      assert.equal(axis === 'row' ? sheet.rowCount : sheet.columnCount, maximum + 1);
      assert.equal(sheet.cells.get(axis === 'row' ? maximum : 0, axis === 'column' ? maximum : 0)?.value, 'last legal cell');
      const beforeFailure = workbook.snapshot();
      assert.throws(() => apply(workbook, params), /outside worksheet bounds/);
      assert.deepEqual(workbook.snapshot(), beforeFailure);
    });

    it(`keeps surviving payload identity during ${axis} cell shifts and snapshots only removed cells`, () => {
      const workbook = new WorkbookModel(`cell-payload-${axis}`, 'Cell payload');
      const sheet = workbook.getSheet('sheet-1');
      sheet.rowCount = 8;
      sheet.columnCount = 8;
      sheet.cells.set(3, 3, { value: 'first', style: { bold: true } });
      sheet.cells.set(axis === 'row' ? 4 : 3, axis === 'column' ? 4 : 3, { value: 'second' });
      const first = sheet.cells.get(3, 3);
      const second = sheet.cells.get(axis === 'row' ? 4 : 3, axis === 'column' ? 4 : 3);
      const range = axis === 'row'
        ? { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 3, endColumn: 3 }
        : { sheetId: sheet.id, startRow: 3, endRow: 3, startColumn: 1, endColumn: 1 };
      const result = apply(workbook, { kind: 'cell-shift', sheetId: sheet.id, axis, operation: 'insert', sourceRange: range });
      assert.equal(sheet.cells.get(axis === 'row' ? 4 : 3, axis === 'column' ? 4 : 3), first);
      assert.equal(sheet.cells.get(axis === 'row' ? 5 : 3, axis === 'column' ? 5 : 3), second);
      assert.deepEqual(result.removedCells, []);
      apply(workbook, { kind: 'cell-shift', sheetId: sheet.id, axis, operation: 'delete', sourceRange: range });
      assert.equal(sheet.cells.get(3, 3), first);
      assert.equal(sheet.cells.get(axis === 'row' ? 4 : 3, axis === 'column' ? 4 : 3), second);
      assert.equal(sheet.rowCount, 8);
      assert.equal(sheet.columnCount, 8);

      const removed = apply(workbook, { kind: axis === 'row' ? 'delete-rows' : 'delete-columns', sheetId: sheet.id, at: 3, count: 1 });
      assert.equal(removed.removedCells.length, 1);
      const snapshot = removed.removedCells[0]!.cell;
      assert.notEqual(snapshot, first);
      assert.notEqual(snapshot.style, first?.style);
      snapshot.style!.bold = false;
      assert.equal(first?.style?.bold, true);
      assert.equal(sheet.cells.get(3, 3), second);
    });
  }

  it('rejects an unwritable surviving cell before deleting any source cells or metadata', () => {
    for (const axis of ['row', 'column'] as const) for (const operation of ['insert', 'delete'] as const) {
      for (const family of ['axis', 'cell'] as const) {
        const workbook = new WorkbookModel(`cell-preflight-${axis}-${operation}-${family}`, 'Cell preflight');
        const sheet = workbook.getSheet('sheet-1');
        sheet.rowCount = 8;
        sheet.columnCount = 8;
        sheet.cells.set(2, 2, { value: 'early valid cell' });
        sheet.cells.set(3, 3, { value: 'invalid survivor', style: { fontFamily: 'Calibri' } });
        sheet.cells.get(3, 3)!.style!.fontFamily = 'invalid\u0000font';
        sheet.review.setNote(2, 2, { id: 'keep-note', author: 'user', text: 'keep', createdAt: 'now', visible: true });
        const range = axis === 'row'
          ? { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 2, endColumn: 3 }
          : { sheetId: sheet.id, startRow: 2, endRow: 3, startColumn: 1, endColumn: 1 };
        const params: StructuralTransformParams = family === 'axis'
          ? { kind: axis === 'row' ? operation === 'insert' ? 'insert-rows' : 'delete-rows'
            : operation === 'insert' ? 'insert-columns' : 'delete-columns', sheetId: sheet.id, at: 1, count: 1 }
          : { kind: 'cell-shift', sheetId: sheet.id, axis, operation, sourceRange: range };
        const before = workbook.snapshot();

        assert.throws(() => apply(workbook, params), /Font family contains control characters/);

        assert.deepEqual(workbook.snapshot(), before);
      }
    }
  });
});
