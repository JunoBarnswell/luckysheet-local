import test from 'node:test';
import assert from 'node:assert/strict';
import {
  CellMatrix,
  isCanonicalFontFamily,
  normalizeFontFamily,
  buildPivotTimelineTiles,
  normalizePivotNumberFormat,
  normalizePivotTimelinePeriod,
  pivotNumericValue,
  pivotTimelineInstant,
  WorkbookModel,
} from './index';
import { assertCanonicalWorkbookSnapshot, migrateStoredWorkbookSnapshot } from './snapshot';

test('v8 storage migration creates the single canonical v10 editing options contract', () => {
  const legacy = structuredClone(new WorkbookModel('unit-v8-editing', 'Legacy').snapshot()) as unknown as Record<string, unknown>;
  legacy.version = 8;
  delete legacy.editingOptions;
  const migrated = migrateStoredWorkbookSnapshot(legacy);
  assert.equal(migrated.version, 10);
  assert.deepEqual(migrated.editingOptions, { allowEditDirectly: true, moveAfterEnter: true, enterDirection: 'down', formulaAutoComplete: true, valueAutoComplete: true, fixedDecimalPlaces: null });
});

test('v9 storage migration folds defined names and legacy cell hyperlinks into canonical owners', () => {
  const legacy = structuredClone(new WorkbookModel('unit-v9-legacy', 'Legacy').snapshot()) as unknown as Record<string, any>;
  legacy.version = 9;
  delete legacy.definedNameModels;
  legacy.definedNames = { SalesTotal: '=Sheet1!A1' };
  legacy.sheets[0].cells = { '0': { '0': {
    value: 'Sales',
    hyperlinkDetail: { id: 'link-1', target: { kind: 'url', url: 'https://example.test/sales' } },
  } } };

  const migrated = migrateStoredWorkbookSnapshot(legacy);
  assert.equal(migrated.version, 10);
  assert.equal(Object.prototype.hasOwnProperty.call(migrated, 'definedNames'), false);
  assert.deepEqual(migrated.definedNameModels, [{ name: 'SalesTotal', formula: '=Sheet1!A1', scope: 'workbook' }]);
  assert.deepEqual(migrated.sheets[0]!.hyperlinks, [{ row: 0, column: 0, hyperlink: { id: 'link-1', target: { kind: 'url', url: 'https://example.test/sales' } } }]);
  assert.equal('hyperlinkDetail' in migrated.sheets[0]!.cells['0']!['0']!, false);
});

test('v9 storage migration fails closed when legacy and canonical representations conflict', () => {
  const legacy = structuredClone(new WorkbookModel('unit-v9-conflict', 'Legacy').snapshot()) as unknown as Record<string, any>;
  legacy.version = 9;
  legacy.definedNames = { SalesTotal: '=Sheet1!A1' };
  legacy.definedNameModels = [{ name: 'SalesTotal', formula: '=Sheet1!B1', scope: 'workbook' }];
  assert.throws(() => migrateStoredWorkbookSnapshot(legacy), /SNAPSHOT_MIGRATION_CONFLICT/);

  const hyperlinkConflict = structuredClone(new WorkbookModel('unit-v9-link-conflict', 'Legacy').snapshot()) as unknown as Record<string, any>;
  hyperlinkConflict.version = 9;
  hyperlinkConflict.sheets[0].cells = { '0': { '0': {
    value: 'Sales',
    hyperlink: 'https://example.test/a',
    hyperlinkDetail: { id: 'link-1', target: { kind: 'url', url: 'https://example.test/b' } },
  } } };
  assert.throws(() => migrateStoredWorkbookSnapshot(hyperlinkConflict), /SNAPSHOT_MIGRATION_CONFLICT/);
});

test('v10 runtime rejects removed snapshot fields', () => {
  const snapshot = new WorkbookModel('unit-v10-reject', 'Canonical').snapshot() as any;
  snapshot.definedNames = {};
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /removed definedNames/);
  delete snapshot.definedNames;
  snapshot.sheets[0].cells = { '0': { '0': { value: 'x', hyperlink: 'https://example.test' } } };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /legacy hyperlink/);
  assert.throws(() => WorkbookModel.fromSnapshot(snapshot), /legacy hyperlink/);
});

test('v9 data-region overlays migrate to canonical CellPatch carriers', () => {
  const legacy = structuredClone(new WorkbookModel('unit-v9-patch', 'Legacy').snapshot()) as unknown as Record<string, any>;
  legacy.version = 9;
  legacy.sheets[0].dataRegions = [{
    id: 'region-1', sourceId: 'source-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0, revision: 0,
  }];
  legacy.sheets[0].cells = { '1': { '1': { value: 999, style: { bold: true } } } };
  const migrated = migrateStoredWorkbookSnapshot(legacy);
  const cell = migrated.sheets[0]!.cells['1']!['1']! as any;
  assert.equal(cell.value, null);
  assert.deepEqual(cell.__cellPatch, { schema: 'CellPatch', value: { kind: 'inherit' }, style: { kind: 'set', value: { bold: true } } });
});

test('v10 canonical snapshots reject raw data-region overlays and inconsistent carriers', () => {
  const snapshot = new WorkbookModel('unit-v10-patch-reject', 'Canonical').snapshot() as any;
  snapshot.sheets[0].dataRegions = [{
    id: 'region-1', sourceId: 'source-1',
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    headerRow: 0, revision: 0,
  }];
  snapshot.sheets[0].cells = { '1': { '1': { value: 999 } } };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /non-canonical cell overlay/);
  snapshot.sheets[0].cells['1']['1'] = { value: null, __cellPatch: { schema: 'CellPatch', value: { kind: 'set', value: 999 } } };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /carrier value disagrees/);
});

test('canonical snapshots bound drawing source work', () => {
  const workbook = new WorkbookModel('unit-drawing-ranges', 'Drawing ranges');
  const snapshot = workbook.snapshot();
  const sheet = snapshot.sheets[0]!;
  sheet.rowCount = 1_000;
  sheet.columnCount = 1_000;
  sheet.drawingPayloads.camera = {
    kind: 'camera',
    sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 9, startColumn: 0, endColumn: 9 },
    refreshPolicy: 'live',
  };
  assert.doesNotThrow(() => assertCanonicalWorkbookSnapshot(snapshot));

  sheet.drawingPayloads.camera.sourceRange.endRow = 999;
  sheet.drawingPayloads.camera.sourceRange.endColumn = 999;
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /rendering limit/);
});

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

test('CellMatrix range iteration visits only persisted cells', () => {
  const matrix = new CellMatrix();
  matrix.set(2, 3, { value: 'inside' });
  matrix.set(20, 3, { value: 'outside' });
  const entries: string[] = [];
  matrix.forEachInRange(0, 10, 0, 10, (_cell, row, column) => entries.push(`${row}:${column}`));
  assert.deepEqual(entries, ['2:3']);
});

test('CellMatrix maintains sparse occupied bounds through overwrite, delete, and clear', () => {
  const matrix = new CellMatrix();
  matrix.set(100_000, 2, { value: 'tail-row' });
  matrix.set(3, 800, { value: 'wide-column' });
  matrix.set(100_000, 800, { value: 'corner' });
  assert.deepEqual(matrix.occupiedRange('sheet-1'), {
    sheetId: 'sheet-1', startRow: 3, endRow: 100_000, startColumn: 2, endColumn: 800,
  });
  matrix.set(100_000, 800, { value: 'overwritten' });
  assert.equal(matrix.count(), 3);
  matrix.delete(100_000, 2);
  matrix.delete(100_000, 800);
  assert.deepEqual(matrix.occupiedRange('sheet-1'), {
    sheetId: 'sheet-1', startRow: 3, endRow: 3, startColumn: 800, endColumn: 800,
  });
  matrix.clear();
  assert.deepEqual(matrix.occupiedRange('sheet-1'), {
    sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0,
  });
});

test('Worksheet used range combines the incremental cell and data-region indexes', () => {
  const workbook = new WorkbookModel('unit-used-range', 'Used range');
  const sheet = workbook.getSheet('sheet-1');
  sheet.addDataRegion({
    id: 'region-1', sourceId: 'source-1',
    range: { sheetId: sheet.id, startRow: 20, endRow: 40, startColumn: 10, endColumn: 30 },
    headerRow: 20, revision: 0,
  });
  sheet.cells.set(5, 2, { value: 'first' });
  sheet.cells.set(50, 1, { value: 'last' });
  assert.deepEqual(sheet.usedRange, {
    sheetId: sheet.id, startRow: 5, endRow: 50, startColumn: 1, endColumn: 30,
  });
  sheet.removeDataRegionAt(0);
  assert.deepEqual(sheet.usedRange, {
    sheetId: sheet.id, startRow: 5, endRow: 50, startColumn: 1, endColumn: 2,
  });
  assert.throws(() => sheet.addDataRegion({
    id: 'wrong-sheet', sourceId: 'source-1',
    range: { sheetId: 'other', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    headerRow: 0, revision: 0,
  }), /belongs to other/);
});

test('font families use one canonical trim/case contract while preserving unknown names', () => {
  assert.equal(normalizeFontFamily('  arial  '), 'Arial');
  assert.equal(normalizeFontFamily('  My Imported Font  '), 'My Imported Font');
  assert.equal(isCanonicalFontFamily('Arial'), true);
  assert.equal(isCanonicalFontFamily(' arial '), false);
  assert.throws(() => normalizeFontFamily('   '), /must not be empty/);
  assert.throws(() => normalizeFontFamily('A\u0000B'), /control characters/);

  const matrix = new CellMatrix();
  matrix.set(0, 0, { value: 'listed', style: { fontFamily: '  SEGOE UI ' } });
  matrix.set(0, 1, { value: 'imported', style: { fontFamily: '  My Imported Font  ' } });
  assert.equal(matrix.get(0, 0)?.style?.fontFamily, 'Segoe UI');
  assert.equal(matrix.get(0, 1)?.style?.fontFamily, 'My Imported Font');
});

test('Pivot numeric value resolution preserves canonical scalar types', () => {
  assert.equal(pivotNumericValue(10), 10);
  assert.equal(pivotNumericValue(0.5), 0.5);
  assert.equal(pivotNumericValue('10'), null);
  assert.equal(pivotNumericValue('$100'), null);
  assert.equal(pivotNumericValue('50%'), null);
  assert.equal(pivotNumericValue(true), null);
  assert.equal(pivotNumericValue(null), null);
  assert.equal(pivotNumericValue(Number.NaN as never), null);
  assert.equal(pivotNumericValue(Number.POSITIVE_INFINITY as never), null);
});

test('Pivot timeline dates use deterministic half-open civil-day bounds', () => {
  const bounds = normalizePivotTimelinePeriod({ start: '2026-08-25', end: '2026-08-25' });
  const dayStart = Date.UTC(2026, 7, 25);
  assert.deepEqual(bounds, { start: dayStart, endExclusive: dayStart + 86_400_000 });
  assert.equal(pivotTimelineInstant('2026-08-25T10:30:00'), Date.UTC(2026, 7, 25, 10, 30));
  assert.equal(pivotTimelineInstant('2026-08-25T10:30:00Z'), Date.UTC(2026, 7, 25, 10, 30));
  assert.equal(pivotTimelineInstant('2026-08-25T10:30:00+02:00'), Date.UTC(2026, 7, 25, 8, 30));
  assert.throws(() => normalizePivotTimelinePeriod({ start: '2026-08-26', end: '2026-08-25' }), /start must not be after end/);
  assert.throws(() => normalizePivotTimelinePeriod({ start: '2026-02-29' }), /Invalid Pivot timeline start date/);
  assert.equal(pivotTimelineInstant('not-a-date'), undefined);
});

test('Pivot timeline tiles provide canonical Years, Quarters, Months and Days levels', () => {
  const values = ['2024-01-15', '2024-04-02', '2025-01-01'];
  assert.deepEqual(buildPivotTimelineTiles(values, 'years').map((tile) => [tile.label, tile.hasData]), [['2024', true], ['2025', true]]);
  assert.deepEqual(buildPivotTimelineTiles(values, 'quarters').map((tile) => tile.label), ['2024 Q1', '2024 Q2', '2024 Q3', '2024 Q4', '2025 Q1']);
  assert.deepEqual(buildPivotTimelineTiles(values, 'months').map((tile) => tile.label), ['2024-01', '2024-02', '2024-03', '2024-04', '2024-05', '2024-06', '2024-07', '2024-08', '2024-09', '2024-10', '2024-11', '2024-12', '2025-01']);
  const days = buildPivotTimelineTiles(values, 'days');
  assert.equal(days[0]?.label, '2024-01-15');
  assert.equal(days.at(-1)?.label, '2025-01-01');
  assert.equal(days.some((tile) => tile.label === '2024-01-16' && !tile.hasData), true);
  assert.throws(() => buildPivotTimelineTiles(values, 'invalid' as never), /Invalid Pivot timeline level/);
});

test('Pivot value field number formats are canonical and fail closed', () => {
  assert.equal(normalizePivotNumberFormat('  #,##0.00  '), '#,##0.00');
  assert.equal(normalizePivotNumberFormat(undefined), undefined);
  assert.throws(() => normalizePivotNumberFormat(''), /must not be empty/);
  assert.throws(() => normalizePivotNumberFormat('[Red'), /unterminated/);
  assert.throws(() => normalizePivotNumberFormat('0.00\\'), /dangling escape/);
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

test('canonical snapshots reject drawings whose Pivot reference no longer exists', () => {
  const workbook = new WorkbookModel('pivot-reference-validation', 'Pivot Reference Validation');
  const sheet = workbook.getSheet('sheet-1');
  sheet.drawings.push({
    id: 'broken-pivot-chart',
    sheetId: sheet.id,
    kind: 'chart',
    payloadId: 'broken-pivot-chart-payload',
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 100, height: 80, rotation: 0 },
    zIndex: 0,
  });
  sheet.drawingPayloads.set('broken-pivot-chart-payload', {
    kind: 'chart',
    chartId: 'broken-pivot-chart',
    source: { kind: 'pivot', pivotId: 'missing-pivot' },
    chartType: 'column',
    subtype: 'clustered',
    elements: { hiddenData: 'show' },
  });
  assert.throws(() => assertCanonicalWorkbookSnapshot(workbook.snapshot()), /references missing Pivot/);
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
  sheet.pane = { kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1, state: 'frozen' };
  sheet.rowHeightsPx[0] = 40;
  sheet.columnWidthsPx[0] = 160;
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
    subtype: 'clustered',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }] },
    elements: { title: 'Revenue', hiddenData: 'show' },
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
  workbook.setDefinedName({ name: 'TaxRate', formula: '0.15', scope: 'workbook' });
  workbook.setDefinedName({ name: 'SharedName', formula: "='Sheet1'!A1", scope: 'workbook' });
  workbook.setDefinedName({ name: 'SharedName', formula: "='Sheet1'!B1", scope: 'sheet', sheetId: 'sheet-1' });
  workbook.setCellStyleTemplate({
    id: 'status-template',
    name: 'Status',
    style: { background: '#e2f0d9', indent: 2 },
    editor: { kind: 'combo-box', items: [{ value: 'Open' }, { value: 'Closed' }], editable: true },
  });

  const snapshot = workbook.snapshot();
  assert.equal(snapshot.schema, 'WorkbookSnapshot');
  assert.deepEqual(snapshot.definedNameModels.find((entry) => entry.name === 'TaxRate'), { name: 'TaxRate', formula: '0.15', scope: 'workbook' });
  assert.equal('charts' in snapshot.sheets[0]!, false);
  assert.equal('shapes' in snapshot.sheets[0]!, false);
  assert.equal('images' in snapshot.sheets[0]!, false);

  const restored = WorkbookModel.fromSnapshot(snapshot);
  const restoredSheet = restored.getSheet('sheet-1');
  assert.equal(restoredSheet.cells.get(1, 2)?.formula, '=40+2');
  assert.equal(restoredSheet.cells.get(1, 2)?.style?.bold, true);
  assert.equal(restoredSheet.pane.kind === 'frozen' ? restoredSheet.pane.ySplit : 0, 1);
  assert.equal(restoredSheet.rowHeightsPx[0], 40);
  assert.equal(restoredSheet.drawings.length, 2);
  assert.equal(restoredSheet.drawingPayloads.get('chart-1')?.kind, 'chart');
  assert.equal((restoredSheet.drawingPayloads.get('chart-1') as { elements?: { title?: string } }).elements?.title, 'Revenue');
  assert.equal(restoredSheet.drawingPayloads.get('shape-1')?.kind, 'shape');
  assert.equal(restoredSheet.drawings.find((drawing) => drawing.id === 'shape-1')?.visible, false);
  assert.equal(restoredSheet.sparklines.length, 1);
  assert.equal(restored.getDefinedNameExact('TaxRate', 'workbook')?.formula, '0.15');
  assert.equal(restored.getDefinedNameExact('SharedName', 'workbook')?.formula, "='Sheet1'!A1");
  assert.equal(restored.getDefinedNameExact('SharedName', 'sheet', 'sheet-1')?.formula, "='Sheet1'!B1");
  assert.equal(restored.getDefinedName('SharedName', 'sheet-1')?.scope, 'sheet');
  assert.equal(restored.listCellStyleTemplates()[0]?.style.indent, 2);
  const restoredEditor = restored.listCellStyleTemplates()[0]?.editor;
  assert.deepEqual(restoredEditor?.kind === 'combo-box' ? restoredEditor.items : undefined, [{ value: 'Open' }, { value: 'Closed' }]);
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
import './protection.test';
