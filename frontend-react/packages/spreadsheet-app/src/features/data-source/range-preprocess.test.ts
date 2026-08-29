import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { preprocessRange } from './range-preprocess';

test('range preprocessing expands one populated seed to its continuous occupied region', () => {
  const workbook = new WorkbookModel('range-preprocess', 'Range preprocess');
  const sheet = workbook.getSheet('sheet-1');
  for (let row = 2; row <= 5; row += 1) {
    for (let column = 1; column <= 3; column += 1) sheet.cells.set(row, column, { value: row === 2 ? `H${column}` : row * 10 + column });
  }
  sheet.cells.set(20, 20, { value: 'unrelated' });
  const result = preprocessRange({
    sheet,
    seed: { sheetId: sheet.id, startRow: 3, endRow: 3, startColumn: 2, endColumn: 2 },
    mode: 'all',
  });
  assert.deepEqual(result.range, { sheetId: sheet.id, startRow: 2, endRow: 5, startColumn: 1, endColumn: 3 });
  assert.equal(result.headerRow, 2);
  assert.equal(result.source, 'continuous-region');
});

test('range preprocessing resolves a table owner before occupied-cell expansion', () => {
  const workbook = new WorkbookModel('range-table', 'Range table');
  const sheet = workbook.getSheet('sheet-1');
  sheet.sheetTables.push({
    id: 'table-1', sheetId: sheet.id, name: 'Orders', range: { sheetId: sheet.id, startRow: 4, endRow: 10, startColumn: 2, endColumn: 6 },
    hasHeaderRow: true, hasTotalRow: false, styleName: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false,
    showBandedRows: true, showBandedColumns: false, showFilterButton: true, autoExpand: 'none', columns: [],
  });
  const result = preprocessRange({ sheet, seed: { sheetId: sheet.id, startRow: 6, endRow: 6, startColumn: 3, endColumn: 3 }, mode: 'all' });
  assert.deepEqual(result.range, sheet.sheetTables[0]!.range);
  assert.equal(result.source, 'sheet-table');
});

test('range preprocessing rejects empty, cross-sheet, and out-of-bounds seeds', () => {
  const workbook = new WorkbookModel('range-reject', 'Range reject');
  const sheet = workbook.getSheet('sheet-1');
  assert.throws(() => preprocessRange({ sheet, seed: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, mode: 'all' }), /RANGE_PREPROCESS_EMPTY/);
  assert.throws(() => preprocessRange({ sheet, seed: { sheetId: 'other', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, mode: 'none' }), /SHEET_MISMATCH/);
  assert.throws(() => preprocessRange({ sheet, seed: { sheetId: sheet.id, startRow: sheet.rowCount, endRow: sheet.rowCount, startColumn: 0, endColumn: 0 }, mode: 'none' }), /OUT_OF_BOUNDS/);
});
