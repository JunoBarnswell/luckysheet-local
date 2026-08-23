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

test('WorkbookModel manages multiple sheets with a stable primary sheet', () => {
  const workbook = new WorkbookModel('unit-sheets', 'MultiSheet');
  assert.equal(workbook.getSheets().length, 1);
  assert.equal(workbook.primarySheetId, 'sheet-1');

  const sheet2 = workbook.addSheet('sheet-2', 'Financials', 500, 50);
  assert.equal(sheet2.name, 'Financials');
  assert.equal(workbook.getSheets().length, 2);
  assert.equal(workbook.getSheetByName('financials')?.id, 'sheet-2');

  workbook.removeSheet('sheet-2');
  assert.equal(workbook.getSheets().length, 1);
  assert.equal(workbook.primarySheetId, 'sheet-1');

  // Removing the only remaining sheet must throw
  assert.throws(() => workbook.removeSheet('sheet-1'), /must keep at least one worksheet/);
});

test('WorkbookSnapshot round-trips complete model state including canonical drawings and metadata', () => {
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
  sheet.drawings.push({
    id: 'chart-1',
    sheetId: 'sheet-1',
    kind: 'chart',
    anchor: { kind: 'absolute' },
    transform: { x: 50, y: 50, width: 400, height: 250 },
    zIndex: 0,
    payloadId: 'chart-1',
  });
  sheet.drawingPayloads.set('chart-1', {
    kind: 'chart',
    chartId: 'chart-1',
    chartType: 'column',
    title: 'Revenue',
    sourceRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }],
  });
  sheet.drawings.push({
    id: 'shape-1',
    sheetId: 'sheet-1',
    kind: 'shape',
    visible: false,
    anchor: { kind: 'absolute' },
    transform: { x: 10, y: 10, width: 100, height: 50 },
    zIndex: 1,
    payloadId: 'shape-1',
  });
  sheet.drawingPayloads.set('shape-1', {
    kind: 'shape',
    type: 'rectangle',
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
  assert.equal(snapshot.schema, 'WorkbookSnapshot');
  assert.equal(snapshot.definedNames?.['TaxRate'], '0.15');
  assert.equal('charts' in snapshot.sheets[0]!, false);
  assert.equal('shapes' in snapshot.sheets[0]!, false);
  assert.equal('images' in snapshot.sheets[0]!, false);

  const restored = WorkbookModel.fromSnapshot(snapshot);
  const restoredSheet = restored.getSheet('sheet-1');
  assert.equal(restoredSheet.cells.get(1, 2)?.formula, '=40+2');
  assert.equal(restoredSheet.cells.get(1, 2)?.style?.bold, true);
  assert.equal(restoredSheet.freeze.ySplit, 1);
  assert.equal(restoredSheet.rowHeights[0], 40);
  assert.equal(restoredSheet.drawings.length, 2);
  assert.equal(restoredSheet.drawingPayloads.get('chart-1')?.kind, 'chart');
  assert.equal((restoredSheet.drawingPayloads.get('chart-1') as { title?: string }).title, 'Revenue');
  assert.equal(restoredSheet.drawingPayloads.get('shape-1')?.kind, 'shape');
  assert.equal(restoredSheet.drawings.find((drawing) => drawing.id === 'shape-1')?.visible, false);
  assert.equal(restoredSheet.sparklines.length, 1);
  assert.equal(restored.definedNames['TaxRate'], '0.15');
});

test('persists print documents and redacted query definitions in the workbook snapshot', () => {
  const workbook = new WorkbookModel('unit-persisted-features', 'Persisted Features');
  const sheetId = workbook.primarySheetId;
  workbook.setPrintDocument({
    schema: 'PrintDocument',
    unitId: workbook.unitId,
    sheetId,
    pageSetup: {
      paperSize: 'letter',
      orientation: 'landscape',
      margins: { top: 10, right: 11, bottom: 12, left: 13, header: 4, footer: 5 },
      scale: 90,
      printGridlines: true,
      printHeadings: false,
      centerHorizontally: true,
      centerVertically: false,
    },
    printAreas: [{ sheetId, range: { sheetId, startRow: 1, endRow: 10, startColumn: 2, endColumn: 6 } }],
    pageBreaks: [{ sheetId, row: 5 }],
  });
  workbook.setQueryDefinition({
    schema: 'QueryDefinition',
    id: 'query-1',
    name: 'Sales',
    connectorId: 'rest',
    connectorConfig: { url: 'https://example.test', apiKey: '[redacted]', nested: { token: '[redacted]' } },
    steps: [{ id: 'source-1', kind: 'source', name: 'Source', config: {}, enabled: true }],
    sourceRevision: 4,
  });

  const snapshot = workbook.snapshot();
  assert.deepEqual(snapshot.printDocuments?.[0]?.pageBreaks, [{ sheetId, row: 5 }]);
  assert.equal(snapshot.queryDefinitions?.[0]?.connectorConfig.apiKey, '[redacted]');
  assert.equal((snapshot.queryDefinitions?.[0]?.connectorConfig.nested as Record<string, unknown>).token, '[redacted]');
  const restored = WorkbookModel.fromSnapshot(snapshot);
  assert.deepEqual(restored.getPrintDocument(sheetId), workbook.getPrintDocument(sheetId));
  assert.deepEqual(restored.getQueryDefinition('query-1'), workbook.getQueryDefinition('query-1'));
  assert.throws(() => restored.setQueryDefinition({
    ...workbook.getQueryDefinition('query-1')!,
    connectorConfig: { apiKey: 'secret' },
  }), /redacted/);
});
