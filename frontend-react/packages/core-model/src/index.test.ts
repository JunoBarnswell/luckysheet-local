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
  WorksheetModel,
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

test('v9 migration extracts legacy cell hyperlinks without hydrating deferred sheets', () => {
  const legacy = structuredClone(new WorkbookModel('unit-v9-hyperlinks', 'Legacy hyperlinks').snapshot()) as unknown as Record<string, any>;
  legacy.version = 9;
  const sheet = legacy.sheets[0] as Record<string, any>;
  sheet.cells = {
    '0': {
      '0': { value: 'legacy url', hyperlink: 'https://legacy.example' },
      '1': { value: 'legacy detail', hyperlinkDetail: { id: 'legacy-detail', target: { kind: 'email', address: 'link@example.com' } } },
    },
  };
  sheet.hyperlinks = [{ row: 0, column: 0, hyperlink: { id: 'canonical', target: { kind: 'url', url: 'https://canonical.example' } } }];

  const migrated = migrateStoredWorkbookSnapshot(legacy);
  const migratedSheet = migrated.sheets[0]!;
  assert.equal(migrated.version, 10);
  assert.deepEqual(migratedSheet.hyperlinks, [
    { row: 0, column: 1, hyperlink: { id: 'legacy-detail', target: { kind: 'email', address: 'link@example.com' } } },
    { row: 0, column: 0, hyperlink: { id: 'canonical', target: { kind: 'url', url: 'https://canonical.example' } } },
  ]);
  assert.equal('hyperlink' in migratedSheet.cells['0']!['0']!, false);
  assert.equal('hyperlinkDetail' in migratedSheet.cells['0']!['1']!, false);

  const restored = WorkbookModel.fromSnapshot(migrated);
  const restoredSheet = restored.getSheet(migratedSheet.id);
  assert.equal(restoredSheet.cells.isHydrated, false);
  assert.equal(restoredSheet.hyperlinks.get('0:0')?.id, 'canonical');
  assert.equal(restoredSheet.hyperlinks.get('0:1')?.id, 'legacy-detail');
});

test('canonical snapshots reject legacy cell hyperlinks and dangling sheet hyperlink targets', () => {
  const snapshot = new WorkbookModel('unit-hyperlink-contract', 'Hyperlink contract').snapshot();
  const sheet = snapshot.sheets[0]!;
  const cells = sheet.cells as unknown as Record<string, Record<string, Record<string, unknown>>>;
  cells['0'] = { '0': { hyperlink: 'https://legacy.example' } };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /legacy hyperlink metadata/);

  delete cells['0'];
  sheet.hyperlinks = [{
    row: 0,
    column: 0,
    hyperlink: { id: 'dangling', target: { kind: 'sheet', sheetId: 'missing-sheet', address: 'A1' } },
  }];
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /target worksheet not found/);
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

test('canonical snapshots enforce worksheet AutoFilter identity and column bounds', () => {
  const workbook = new WorkbookModel('unit-auto-filter-contract', 'AutoFilter contract');
  const snapshot = workbook.snapshot();
  const sheet = snapshot.sheets[0]!;
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };

  sheet.autoFilter = { sheetId: 'other-sheet', range, columns: {} };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /AutoFilter must target its worksheet/);

  sheet.autoFilter = {
    sheetId: sheet.id,
    range,
    columns: { 2: { column: 2, showButton: true, hiddenButton: false } },
  };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /AutoFilter column identity is invalid/);

  const aliasedColumns = {} as NonNullable<typeof sheet.autoFilter>['columns'];
  aliasedColumns[0] = { column: 0, showButton: true, hiddenButton: false };
  Object.defineProperty(aliasedColumns, '00', {
    value: { column: 0, showButton: true, hiddenButton: false },
    enumerable: true,
  });
  sheet.autoFilter = { sheetId: sheet.id, range, columns: aliasedColumns };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /AutoFilter column identity is invalid/);
});

test('canonical snapshots reject persisted font metadata that cannot be normalized', () => {
  const snapshot = new WorkbookModel('unit-invalid-font-snapshot', 'Invalid font snapshot').snapshot();
  snapshot.sheets[0]!.cells['0'] = { '0': { value: 'invalid', style: { fontFamily: '  ' } } };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /Font family must not be empty/);
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

test('CellMatrix range traversal stays ordered and refreshes row indexes after sparse mutations', () => {
  const matrix = new CellMatrix();
  matrix.set(100_000, 2, { value: 'far' });
  matrix.set(50, 2, { value: 'middle' });
  matrix.set(2, 2, { value: 'near' });

  const visited: number[] = [];
  matrix.forEachInRange(2, 50, 2, 2, (_cell, row) => visited.push(row));
  assert.deepEqual(visited, [2, 50]);

  matrix.set(5, 2, { value: 'inserted-row' });
  matrix.delete(50, 2);
  assert.deepEqual(matrix.getRegion(2, 5, 2, 2).map(({ row }) => row), [2, 5]);

  matrix.shiftRows(5, 2, 1);
  assert.deepEqual(matrix.getRegion(5, 7, 2, 2).map(({ row }) => row), [7]);
});

test('CellMatrix enumerates non-calculation formula owners without hydrating deferred cells', () => {
  const matrix = new CellMatrix();
  matrix.deferJSON({
    '2': {
      '3': {
        value: null,
        formula: '=A1',
        formulaMetadata: { kind: 'array', preservedOnly: true, sourceFormula: '=A1' },
      },
      '4': {
        value: null,
        presentation: {
          kind: 'barcode',
          symbology: 'qr',
          source: { kind: 'formula', formula: '=A1' },
          parameters: { symbology: 'qr' },
          options: { foreground: '#000000', background: '#ffffff', showText: true, labelPosition: 'below', quietZone: 2 },
        },
      },
    },
  });

  const owners: string[] = [];
  matrix.forEachFormulaOwner((_cell, row, column) => owners.push(`${row}:${column}`));
  assert.deepEqual(owners, ['2:3', '2:4']);
  assert.equal(matrix.isHydrated, false);
});

test('CellMatrix keeps deferred cells intact when normalization fails during hydration', () => {
  const matrix = new CellMatrix();
  const deferred = {
    '0': {
      '0': { value: 'valid' },
      '1': { value: 'invalid', style: { fontFamily: '  ' } },
    },
  };
  matrix.deferJSON(deferred);
  assert.throws(() => matrix.get(0, 0), /Font family must not be empty/);
  assert.equal(matrix.isHydrated, false);
  assert.deepEqual(matrix.toJSON(), deferred);
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

test('CellMatrix rejects invalid cell metadata before expanding worksheet extent', () => {
  const deferred = new CellMatrix();
  deferred.deferJSON({ '4': { '2': { value: 'existing' } } });
  assert.throws(() => deferred.set(1_200, 40, { value: 'invalid', style: { fontFamily: '  ' } }), /must not be empty/);
  assert.equal(deferred.isHydrated, false);
  assert.equal(deferred.count(), 1);

  const sheet = new WorksheetModel('atomic-cell-write', 'Atomic cell write');
  assert.throws(() => sheet.cells.set(1_200, 40, { value: 'invalid', style: { fontFamily: '  ' } }), /must not be empty/);
  assert.equal(sheet.cells.count(), 0);
  assert.equal(sheet.rowCount, 1_000);
  assert.equal(sheet.columnCount, 26);

  sheet.cells.set(1_200, 40, { value: 'valid', style: { fontFamily: ' Arial ' } });
  assert.equal(sheet.cells.get(1_200, 40)?.style?.fontFamily, 'Arial');
  assert.equal(sheet.rowCount, 1_201);
  assert.equal(sheet.columnCount, 41);
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

test('sheet deletion rejects external formula anchors, validation ranges, chart ranges, and shape formulas', () => {
  const assertSourceSheetDeletionRejected = (workbook: WorkbookModel, sourceSheetId: string): void => {
    const originalOrder = [...workbook.sheetOrder];
    assert.throws(() => workbook.removeSheet(sourceSheetId), /external references must be resolved first/);
    assert.equal(workbook.sheets.has(sourceSheetId), true);
    assert.deepEqual(workbook.sheetOrder, originalOrder);
  };

  const validationWorkbook = new WorkbookModel('delete-validation-source', 'Delete validation source');
  const validationOwner = validationWorkbook.getSheet('sheet-1');
  const validationSource = validationWorkbook.addSheet('validation-source', 'Validation Source');
  validationOwner.dataValidations.push({
    id: 'cross-sheet-list',
    sheetId: validationOwner.id,
    ranges: [{ sheetId: validationOwner.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    type: 'list',
    listSource: { kind: 'range', range: { sheetId: validationSource.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 } },
  });
  assertSourceSheetDeletionRejected(validationWorkbook, validationSource.id);

  const anchorWorkbook = new WorkbookModel('delete-anchor-source', 'Delete anchor source');
  const anchorSource = anchorWorkbook.addSheet('anchor-source', 'Anchor Source');
  anchorWorkbook.setDefinedName({
    name: 'RelativeName', formula: 'A1', scope: 'workbook',
    anchor: { sheetId: anchorSource.id, row: 0, column: 0 },
  });
  assertSourceSheetDeletionRejected(anchorWorkbook, anchorSource.id);

  const chartWorkbook = new WorkbookModel('delete-chart-source', 'Delete chart source');
  const chartOwner = chartWorkbook.getSheet('sheet-1');
  const chartSource = chartWorkbook.addSheet('chart-source', 'Chart Source');
  chartOwner.drawings.push({
    id: 'cross-sheet-chart', sheetId: chartOwner.id, kind: 'chart', payloadId: 'cross-sheet-chart',
    anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 200, height: 120 }, zIndex: 0,
  });
  chartOwner.drawingPayloads.set('cross-sheet-chart', {
    kind: 'chart', chartId: 'cross-sheet-chart', chartType: 'column', subtype: 'clustered',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: chartSource.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
    elements: { hiddenData: 'show' },
  });
  assertSourceSheetDeletionRejected(chartWorkbook, chartSource.id);

  const shapeWorkbook = new WorkbookModel('delete-shape-source', 'Delete shape source');
  const shapeOwner = shapeWorkbook.getSheet('sheet-1');
  const shapeSource = shapeWorkbook.addSheet('shape-source', 'Shape Source');
  shapeOwner.drawings.push({
    id: 'formula-shape', sheetId: shapeOwner.id, kind: 'shape', payloadId: 'formula-shape',
    anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 100, height: 60 }, zIndex: 0,
  });
  shapeOwner.drawingPayloads.set('formula-shape', {
    kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000',
    propertyFormula: "='Shape Source'!A1",
  });
  assertSourceSheetDeletionRejected(shapeWorkbook, shapeSource.id);

  const barcodeWorkbook = new WorkbookModel('delete-barcode-source', 'Delete barcode source');
  const barcodeOwner = barcodeWorkbook.getSheet('sheet-1');
  const barcodeSource = barcodeWorkbook.addSheet('barcode-source', 'Barcode Source');
  barcodeOwner.cells.set(0, 0, {
    value: null,
    presentation: {
      kind: 'barcode', symbology: 'code128',
      source: { kind: 'formula', formula: "='Barcode Source'!A1" },
      parameters: { symbology: 'code128' },
      options: { foreground: '#000000', background: '#ffffff', showText: true, labelPosition: 'below', quietZone: 2 },
    },
  });
  assertSourceSheetDeletionRejected(barcodeWorkbook, barcodeSource.id);

  const preservedFormulaWorkbook = new WorkbookModel('delete-preserved-formula-source', 'Delete preserved formula source');
  const preservedFormulaOwner = preservedFormulaWorkbook.getSheet('sheet-1');
  const preservedFormulaSource = preservedFormulaWorkbook.addSheet('preserved-formula-source', 'Preserved Formula Source');
  preservedFormulaOwner.cells.set(0, 0, {
    value: null,
    formulaMetadata: {
      kind: 'dataTable', preservedOnly: true, reason: 'Native formula retained from OOXML',
      sourceFormula: "='Preserved Formula Source'!A1",
    },
  });
  assertSourceSheetDeletionRejected(preservedFormulaWorkbook, preservedFormulaSource.id);

  const tableViewWorkbook = new WorkbookModel('delete-table-view-formula-source', 'Delete table view formula source');
  const tableViewOwner = tableViewWorkbook.getSheet('sheet-1');
  const tableViewSource = tableViewWorkbook.addSheet('table-view-source', 'Table View Source');
  tableViewOwner.kind = 'table-sheet';
  tableViewOwner.tableSheet = {
    viewId: 'table-view',
    columns: [{ fieldId: 'calculated', caption: 'Calculated', type: 'formula', formula: "='Table View Source'!A1" }],
    grouping: [],
  };
  assertSourceSheetDeletionRejected(tableViewWorkbook, tableViewSource.id);
});

test('sheet rename and duplication preserve every persisted formula owner identity', () => {
  const workbook = new WorkbookModel('sheet-formula-identity', 'Sheet formula identity');
  const source = workbook.addSheet('formula-source', 'Source');
  source.cells.set(0, 0, {
    value: null,
    formula: "='Source'!A1",
    formulaMetadata: { kind: 'normal', sourceFormula: "='Source'!A1" },
  });
  source.cells.set(0, 1, {
    value: null,
    presentation: {
      kind: 'barcode', symbology: 'code128',
      source: { kind: 'formula', formula: "='Source'!A1" },
      parameters: { symbology: 'code128' },
      options: { foreground: '#000000', background: '#ffffff', showText: true, labelPosition: 'below', quietZone: 2 },
    },
  });
  source.kind = 'table-sheet';
  source.tableSheet = {
    viewId: 'formula-view',
    columns: [{ fieldId: 'calculated', caption: 'Calculated', type: 'formula', formula: "='Source'!A1" }],
    grouping: [],
  };
  source.drawings.push({
    id: 'formula-shape', sheetId: source.id, kind: 'shape', payloadId: 'formula-shape',
    anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 100, height: 60 }, zIndex: 0,
  });
  source.drawingPayloads.set('formula-shape', {
    kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000',
    propertyFormula: "='Source'!A1",
  });
  workbook.setDefinedName({
    name: 'RelativeName', formula: "='Source'!A1", scope: 'sheet', sheetId: source.id,
    anchor: { sheetId: source.id, row: 0, column: 0 },
  });
  workbook.setCellStyleTemplate({
    id: 'formula-template', name: 'Formula template', style: {},
    dataValidation: { type: 'custom', formula1: "='Source'!A1" },
  });
  workbook.dataModel.views.set('formula-view', {
    id: 'formula-view', name: 'Formula view', tableId: 'source-table',
    fields: [{ fieldId: 'calculated', caption: 'Calculated', formula: "='Source'!A1" }],
  });
  const externalOwner = workbook.getSheet('sheet-1');
  externalOwner.cells.set(2, 2, {
    value: null,
    presentation: {
      kind: 'barcode', symbology: 'code128',
      source: { kind: 'formula', formula: "='Source'!A1" },
      parameters: { symbology: 'code128' },
      options: { foreground: '#000000', background: '#ffffff', showText: true, labelPosition: 'below', quietZone: 2 },
    },
  });
  externalOwner.drawings.push({
    id: 'external-formula-shape', sheetId: externalOwner.id, kind: 'shape', payloadId: 'external-formula-shape',
    anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 100, height: 60 }, zIndex: 0,
  });
  externalOwner.drawingPayloads.set('external-formula-shape', {
    kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000',
    propertyFormula: "='Source'!A1",
  });

  workbook.renameSheet(source.id, 'Renamed Sheet');
  const renamedBarcode = source.cells.get(0, 1)?.presentation;
  const renamedExternalBarcode = externalOwner.cells.get(2, 2)?.presentation;
  assert.equal(source.cells.get(0, 0)?.formula, "='Renamed Sheet'!A1");
  assert.equal(source.cells.get(0, 0)?.formulaMetadata?.sourceFormula, "='Renamed Sheet'!A1");
  assert.equal(renamedBarcode?.kind === 'barcode' && renamedBarcode.source.kind === 'formula' ? renamedBarcode.source.formula : undefined, "='Renamed Sheet'!A1");
  assert.equal(source.tableSheet?.columns[0]?.formula, "='Renamed Sheet'!A1");
  assert.equal((source.drawingPayloads.get('formula-shape') as { propertyFormula?: string }).propertyFormula, "='Renamed Sheet'!A1");
  assert.equal(workbook.dataModel.views.get('formula-view')?.fields[0]?.formula, "='Renamed Sheet'!A1");
  assert.equal(workbook.cellStyleTemplates.get('formula-template')?.dataValidation?.formula1, "='Renamed Sheet'!A1");
  assert.equal(renamedExternalBarcode?.kind === 'barcode' && renamedExternalBarcode.source.kind === 'formula' ? renamedExternalBarcode.source.formula : undefined, "='Renamed Sheet'!A1");
  assert.equal((externalOwner.drawingPayloads.get('external-formula-shape') as { propertyFormula?: string }).propertyFormula, "='Renamed Sheet'!A1");

  const duplicate = workbook.duplicateSheet(source.id, 'formula-copy', 'Copy Sheet');
  const duplicatedBarcode = duplicate.cells.get(0, 1)?.presentation;
  assert.equal(duplicate.cells.get(0, 0)?.formula, "='Copy Sheet'!A1");
  assert.equal(duplicate.cells.get(0, 0)?.formulaMetadata?.sourceFormula, "='Copy Sheet'!A1");
  assert.equal(duplicatedBarcode?.kind === 'barcode' && duplicatedBarcode.source.kind === 'formula' ? duplicatedBarcode.source.formula : undefined, "='Copy Sheet'!A1");
  assert.equal(duplicate.tableSheet?.columns[0]?.formula, "='Copy Sheet'!A1");
  assert.equal((duplicate.drawingPayloads.get('formula-shape::formula-copy') as { propertyFormula?: string } | undefined)?.propertyFormula, "='Copy Sheet'!A1");
  assert.equal(workbook.definedNameModels.find((entry) => entry.name === 'RelativeName' && entry.sheetId === duplicate.id)?.anchor?.sheetId, duplicate.id);
  assert.equal(workbook.dataModel.views.get('formula-view')?.fields[0]?.formula, "='Renamed Sheet'!A1");
  assert.equal(workbook.cellStyleTemplates.get('formula-template')?.dataValidation?.formula1, "='Renamed Sheet'!A1");
  const unchangedExternalBarcode = externalOwner.cells.get(2, 2)?.presentation;
  assert.equal(unchangedExternalBarcode?.kind === 'barcode' && unchangedExternalBarcode.source.kind === 'formula' ? unchangedExternalBarcode.source.formula : undefined, "='Renamed Sheet'!A1");
  assert.equal((externalOwner.drawingPayloads.get('external-formula-shape') as { propertyFormula?: string }).propertyFormula, "='Renamed Sheet'!A1");
});

test('sheet rename fails closed on a preserved-only formula reference without changing the workbook', () => {
  const workbook = new WorkbookModel('preserved-formula-rename', 'Preserved formula rename');
  const source = workbook.addSheet('preserved-source', 'Preserved Source');
  const formula = "='Preserved Source'!A1";
  source.cells.set(0, 0, {
    value: null,
    formulaMetadata: { kind: 'dataTable', preservedOnly: true, sourceFormula: formula, reason: 'Native OOXML formula' },
  });

  assert.throws(() => workbook.renameSheet(source.id, 'Renamed Source'), /cannot be rewritten safely/);
  assert.equal(source.name, 'Preserved Source');
  assert.equal(source.cells.get(0, 0)?.formulaMetadata?.sourceFormula, formula);
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
  assert.equal(snapshot.definedNames?.['TaxRate'], '0.15');
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
  assert.equal(restored.definedNames['TaxRate'], '0.15');
  assert.equal(restored.getDefinedNameExact('SharedName', 'workbook')?.formula, "='Sheet1'!A1");
  assert.equal(restored.getDefinedNameExact('SharedName', 'sheet', 'sheet-1')?.formula, "='Sheet1'!B1");
  assert.equal(restored.getDefinedName('SharedName', 'sheet-1')?.scope, 'sheet');
  assert.equal(restored.listCellStyleTemplates()[0]?.style.indent, 2);
  const restoredEditor = restored.listCellStyleTemplates()[0]?.editor;
  assert.deepEqual(restoredEditor?.kind === 'combo-box' ? restoredEditor.items : undefined, [{ value: 'Open' }, { value: 'Closed' }]);
});

test('defers sparse worksheet cell hydration until the sheet is read', () => {
  const source = new WorkbookModel('lazy-sheet-hydration', 'Lazy sheets');
  const second = source.addSheet('sheet-2', 'Second');
  source.getSheet('sheet-1').cells.set(2, 3, { value: 'first' });
  second.cells.set(100, 4, { value: 'second' });

  const restored = WorkbookModel.fromSnapshot(source.snapshot());
  const first = restored.getSheet('sheet-1');
  const deferred = restored.getSheet('sheet-2');

  assert.equal(first.cells.isHydrated, false);
  assert.equal(deferred.cells.isHydrated, false);
  assert.deepEqual(deferred.cells.occupiedRange(deferred.id), {
    sheetId: 'sheet-2', startRow: 100, endRow: 100, startColumn: 4, endColumn: 4,
  });
  assert.equal(deferred.cells.count(), 1);
  assert.equal(deferred.cells.revision, 1);
  assert.equal(deferred.cells.isHydrated, false);
  assert.equal(deferred.cells.get(100, 4)?.value, 'second');
  assert.equal(deferred.cells.isHydrated, true);
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
