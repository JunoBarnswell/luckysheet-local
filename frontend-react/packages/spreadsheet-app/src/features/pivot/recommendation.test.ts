import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { recommendPivotTables } from './recommendation';

test('Recommended PivotTables derives row, column and value placements from field types', () => {
  const workbook = new WorkbookModel('pivot-recommendations', 'Pivot Recommendations');
  const sheet = workbook.getSheet('sheet-1');
  ['Region', 'Quarter', 'Revenue'].forEach((value, column) => sheet.cells.set(0, column, { value }));
  [['East', 'Q1', 10], ['West', 'Q2', 20]].forEach((row, rowIndex) => row.forEach((value, column) => sheet.cells.set(rowIndex + 1, column, { value })));
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };
  const candidates = recommendPivotTables(workbook, sheet.id, range);
  assert.ok(candidates.length >= 2);
  assert.equal(candidates[0]?.layout.rows.length, 1);
  assert.equal(candidates[0]?.layout.values[0]?.summarizeBy, 'sum');
  assert.deepEqual(candidates[0]?.source.range, range);
});

test('Recommended PivotTables rejects a header-only region', () => {
  const workbook = new WorkbookModel('pivot-recommendations-reject', 'Pivot Recommendations Reject');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Revenue' });
  assert.throws(() => recommendPivotTables(workbook, sheet.id, { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }), /PIVOT_SOURCE_INVALID/);
});
