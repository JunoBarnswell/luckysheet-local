import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { recommendCharts } from './recommendation';

test('Recommended Charts derives ordered real candidates from typed selection data', () => {
  const workbook = new WorkbookModel('chart-recommendations', 'Chart Recommendations');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Region' });
  sheet.cells.set(0, 1, { value: 'Revenue' });
  const rows: ReadonlyArray<readonly [string, number]> = [['East', 10], ['West', 20], ['North', 15]];
  rows.forEach((row, index) => {
    sheet.cells.set(index + 1, 0, { value: row[0] });
    sheet.cells.set(index + 1, 1, { value: row[1] });
  });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
  const candidates = recommendCharts(workbook, range);
  assert.equal(candidates[0]?.chartType, 'column');
  assert.equal(candidates[0]?.subtype, 'clustered');
  assert.ok(candidates.some((candidate) => candidate.chartType === 'pie'));
  assert.deepEqual(candidates[0]?.source.ranges[0], range);
});

test('Recommended Charts rejects a non-numeric selection without fabricating candidates', () => {
  const workbook = new WorkbookModel('chart-recommendations-reject', 'Chart Recommendations Reject');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Name' });
  sheet.cells.set(1, 0, { value: 'Alpha' });
  assert.throws(() => recommendCharts(workbook, { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }), /INVALID_CHART_SOURCE/);
});
