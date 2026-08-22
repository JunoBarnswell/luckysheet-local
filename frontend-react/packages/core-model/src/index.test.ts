import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMatrix, WorkbookModel } from './index';

test('CellMatrix keeps empty logical space sparse', () => {
  const matrix = new CellMatrix();
  matrix.set(100_000, 4, { value: 'tail' });
  assert.equal(matrix.get(0, 0), undefined);
  assert.equal(matrix.get(100_000, 4)?.value, 'tail');
  assert.deepEqual(Object.keys(matrix.toJSON()), ['100000']);
  assert.equal(matrix.has(100_000, 4), true);
  assert.equal(matrix.has(0, 0), false);
  assert.equal(matrix.count(), 1);

  // clone & delete
  const cloned = matrix.clone();
  assert.equal(cloned.get(100_000, 4)?.value, 'tail');
  matrix.delete(100_000, 4);
  assert.equal(matrix.has(100_000, 4), false);
  assert.equal(matrix.count(), 0);
  assert.equal(cloned.has(100_000, 4), true);
});

test('WorksheetModel handles merges and anchors properly', () => {
  const workbook = new WorkbookModel('unit-merge', 'Merge Test');
  const sheet = workbook.getSheet('sheet-1');
  sheet.merges.push({
    range: { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 },
    anchor: { row: 1, column: 1 },
  });

  assert.equal(sheet.isMergeAnchor(1, 1), true);
  assert.equal(sheet.isMergeAnchor(2, 2), false);
  assert.equal(sheet.isMergeAnchor(0, 0), true); // unmerged cells are their own anchor
  assert.ok(sheet.isMerged(2, 2));
  assert.equal(sheet.isMerged(0, 0), undefined);
});

test('WorkbookModel manages multiple sheets and preserves activeSheetId', () => {
  const workbook = new WorkbookModel('unit-sheets', 'MultiSheet');
  assert.equal(workbook.getSheets().length, 1);
  assert.equal(workbook.activeSheetId, 'sheet-1');

  const sheet2 = workbook.addSheet('sheet-2', 'Financials', 500, 50);
  assert.equal(sheet2.name, 'Financials');
  assert.equal(workbook.getSheets().length, 2);
  assert.equal(workbook.getSheetByName('financials')?.id, 'sheet-2');

  workbook.activeSheetId = 'sheet-2';
  workbook.removeSheet('sheet-2');
  assert.equal(workbook.getSheets().length, 1);
  assert.equal(workbook.activeSheetId, 'sheet-1');

  // Removing the only remaining sheet must throw
  assert.throws(() => workbook.removeSheet('sheet-1'), /must keep at least one worksheet/);
});

test('WorkbookSnapshotV1 round-trips complete model state including Pro features and metadata', () => {
  const workbook = new WorkbookModel('unit-full', 'Full Test');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(1, 2, {
    value: 42,
    formula: '=40+2',
    style: { bold: true, background: '#fef08a', textColor: '#854d0e', numberFormat: '$#,##0' },
  });
  sheet.merges.push({
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 },
    anchor: { row: 0, column: 0 },
  });
  sheet.freeze = { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 };
  sheet.rowHeights[0] = 40;
  sheet.columnWidths[0] = 160;
  sheet.hiddenRows.add(5);
  sheet.charts.push({
    id: 'chart-1',
    sheetId: 'sheet-1',
    type: 'column',
    title: 'Revenue',
    sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }],
    bounds: { x: 50, y: 50, width: 400, height: 250 },
  });
  sheet.shapes.push({
    id: 'shape-1',
    sheetId: 'sheet-1',
    type: 'rectangle',
    bounds: { x: 10, y: 10, width: 100, height: 50 },
    fill: '#3b82f6',
    stroke: '#1d4ed8',
  });
  sheet.sparklines.push({
    id: 'spark-1',
    sheetId: 'sheet-1',
    anchor: { row: 2, column: 5 },
    sourceRange: [{ sheetId: 'sheet-1', startRow: 2, endRow: 2, startColumn: 0, endColumn: 4 }] as any,
    type: 'line',
    color: '#10b981',
  });
  workbook.definedNames['TaxRate'] = '0.15';

  const snapshot = workbook.snapshot();
  assert.equal(snapshot.schema, 'WorkbookSnapshotV1');
  assert.equal(snapshot.definedNames?.['TaxRate'], '0.15');

  const restored = WorkbookModel.fromSnapshot(snapshot);
  const restoredSheet = restored.getSheet('sheet-1');
  assert.equal(restoredSheet.cells.get(1, 2)?.formula, '=40+2');
  assert.equal(restoredSheet.cells.get(1, 2)?.style?.bold, true);
  assert.equal(restoredSheet.freeze.ySplit, 1);
  assert.equal(restoredSheet.rowHeights[0], 40);
  assert.equal(restoredSheet.charts.length, 1);
  assert.equal(restoredSheet.charts[0]?.title, 'Revenue');
  assert.equal(restoredSheet.shapes.length, 1);
  assert.equal(restoredSheet.sparklines.length, 1);
  assert.equal(restored.definedNames['TaxRate'], '0.15');
});
