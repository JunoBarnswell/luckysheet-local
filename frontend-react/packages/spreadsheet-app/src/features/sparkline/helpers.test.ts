import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import { extractSparklineValues } from './helpers';

describe('sparkline helpers', () => {
  it('extractSparklineValues reads numeric source cells in range order', () => {
    const sheet = new WorksheetModel('s1', 'Sheet1');
    sheet.cells.set(0, 0, { value: 1 });
    sheet.cells.set(0, 1, { value: 3 });
    sheet.cells.set(0, 2, { value: 2 });
    const values = extractSparklineValues(sheet, {
      id: 'spark-1',
      sheetId: 's1',
      anchor: { row: 0, column: 3 },
      sourceRange: { sheetId: 's1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      type: 'line',
      color: '#2563eb',
    });
    assert.deepEqual(values, [1, 3, 2]);
  });

  it('extractSparklineValues resolves a source on another worksheet', () => {
    const workbook = new WorkbookModel('sparkline-helper-cross-sheet', 'Sparkline Helper');
    const source = workbook.addSheet('source-2', 'Source 2');
    source.cells.set(0, 0, { value: 4 });
    source.cells.set(0, 1, { value: 6 });
    const values = extractSparklineValues(workbook, {
      id: 'spark-cross',
      sheetId: 'sheet-1',
      anchor: { row: 0, column: 3 },
      sourceRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      type: 'line',
      color: '#2563eb',
    });
    assert.deepEqual(values, [4, 6]);
  });
});
