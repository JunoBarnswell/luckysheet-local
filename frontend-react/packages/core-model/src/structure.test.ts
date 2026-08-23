import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CellMatrix, WorkbookModel } from './index';

function seedWorkbook(): { workbook: WorkbookModel; sheetId: string } {
  const workbook = new WorkbookModel('unit-test', 'Structural');
  const sheet = workbook.addSheet('s1', 'Sheet1');
  sheet.cells.set(0, 0, { value: 'A0' });
  sheet.cells.set(1, 0, { value: 'A1' });
  sheet.cells.set(2, 0, { value: 'A2' });
  return { workbook, sheetId: sheet.id };
}

describe('structural operations', () => {
  it('shiftRows moves cells below the insertion point', () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { value: 'top' });
    matrix.set(5, 1, { value: 'bottom' });
    matrix.shiftRows(3, 2, 1);
    assert.equal(matrix.get(0, 0)?.value, 'top');
    assert.equal(matrix.get(7, 1)?.value, 'bottom');
    matrix.shiftRows(3, 2, -1);
    assert.equal(matrix.get(5, 1)?.value, 'bottom');
  });

  it('worksheet insertRows keeps merges and freeze consistent', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.merges.push({ range: { sheetId: 's1', startRow: 2, endRow: 3, startColumn: 1, endColumn: 1 }, anchor: { row: 2, column: 1 } });
    sheet.freeze = { xSplit: 0, ySplit: 2, startRow: 2, startColumn: 0 };
    sheet.insertRows(2, 3);
    assert.equal(sheet.cells.get(2, 0)?.value ?? null, null);
    assert.equal(sheet.cells.get(5, 0)?.value, 'A2');
    assert.equal(sheet.merges[0]!.range.startRow, 5);
    assert.equal(sheet.freeze.ySplit, 5);
    assert.equal(sheet.rowCount, 1003);
  });

  it('deleteRows removes region and returns extracted cells for undo', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.cells.set(1, 0, { value: 'gone' });
    const removed = sheet.deleteRows(1, 1);
    assert.ok(removed.some((entry) => entry.cell.value === 'gone'));
    assert.equal(sheet.cells.get(1, 0)?.value, 'A2');
    // 恢复
    sheet.insertRows(1, 1);
    for (const entry of removed) sheet.cells.set(entry.row, entry.column, entry.cell);
    assert.equal(sheet.cells.get(1, 0)?.value, 'gone');
  });

  it('insertColumns shifts widths and hidden columns', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.columnWidths[2] = 200;
    sheet.hiddenColumns.add(3);
    sheet.insertColumns(1, 2);
    assert.equal(sheet.columnWidths[4], 200);
    assert.ok(sheet.hiddenColumns.has(5));
    assert.ok(!sheet.hiddenColumns.has(3));
  });
});
