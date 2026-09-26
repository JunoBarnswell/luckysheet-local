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
  planSheetTableRename,
  WorkbookModel,
  WorksheetModel,
} from './index';
import { assertCanonicalWorkbookSnapshot, migrateStoredWorkbookSnapshot, type WorkbookSnapshot } from './snapshot';

test('canonical workbook snapshots reject malformed or unowned pane fields', () => {
  const panes = [
    { kind: 'frozen', state: 'frozen', xSplit: 16_385, ySplit: 0, startRow: 0, startColumn: 0 },
    { kind: 'frozen', state: 'frozen', xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0, referenceHint: 'A1' },
    { kind: 'none', referenceHint: 'A1' },
  ];
  for (const pane of panes) {
    const candidate = structuredClone(new WorkbookModel('unit-pane-snapshot-validation', 'Pane snapshot').snapshot()) as unknown as Record<string, any>;
    candidate.sheets[0].pane = pane;
    assert.throws(() => assertCanonicalWorkbookSnapshot(candidate as unknown as WorkbookSnapshot), /Workbook snapshot pane/);
    assert.throws(() => WorkbookModel.fromSnapshot(candidate as unknown as WorkbookSnapshot), /Workbook snapshot pane/);
  }
});

test('worksheet identity collisions are rejected before loading, creating, renaming, or duplicating sheets', () => {
  const workbook = new WorkbookModel('unit-sheet-identities', 'Sheet identities');
  const duplicateId = structuredClone(workbook.snapshot());
  duplicateId.sheets.push({ ...structuredClone(duplicateId.sheets[0]!), name: 'Second' });
  assert.throws(() => assertCanonicalWorkbookSnapshot(duplicateId), /duplicate worksheet identity/);
  assert.throws(() => WorkbookModel.fromSnapshot(duplicateId), /duplicate worksheet identity/);

  const duplicateName = structuredClone(workbook.snapshot());
  duplicateName.sheets.push({ ...structuredClone(duplicateName.sheets[0]!), id: 'sheet-2', name: 'sHeEt1' });
  assert.throws(() => assertCanonicalWorkbookSnapshot(duplicateName), /duplicate worksheet name/);
  assert.throws(() => WorkbookModel.fromSnapshot(duplicateName), /duplicate worksheet name/);

  const other = workbook.addSheet('sheet-2', 'Budget');
  assert.throws(() => workbook.addSheet('sheet-3', 'budget'), /duplicate worksheet name/);
  assert.throws(() => workbook.renameSheet(workbook.primarySheetId, 'BUDGET'), /Sheet name already exists/);
  assert.throws(() => workbook.duplicateSheet(other.id, 'sheet-3', 'bUdGeT'), /Sheet name already exists/);
  for (const invalidName of ['Bad/Name', 'Bad\\Name', 'Bad?Name', 'Bad*Name', 'Bad:Name', 'Bad[Name]', "'Quoted", "Quoted'", 'History', 'a'.repeat(32)]) {
    assert.throws(() => workbook.addSheet(`invalid-${invalidName}`, invalidName), /Excel naming rules/);
    assert.throws(() => workbook.renameSheet(other.id, invalidName), /Excel naming rules/);
    assert.throws(() => workbook.duplicateSheet(other.id, `invalid-copy-${invalidName}`, invalidName), /Excel naming rules/);
  }
  const maxLengthWorkbook = new WorkbookModel('unit-sheet-name-limit', 'Sheet name limit');
  assert.equal(maxLengthWorkbook.addSheet('sheet-2', 'a'.repeat(31)).name, 'a'.repeat(31));
  assert.equal(maxLengthWorkbook.addSheet('sheet-3', "O'Brien").name, "O'Brien");
  assert.equal(workbook.getSheet(workbook.primarySheetId).name, 'Sheet1');
  assert.equal(workbook.sheetOrder.length, 2);
});

test('snapshot hydration rejects duplicate map-backed workbook owner identities', () => {
  const workbook = new WorkbookModel('unit-owner-identities', 'Owner identities');
  const duplicateOwners: Array<[string, (snapshot: Record<string, any>) => void]> = [
    ['workbook table', (snapshot) => { snapshot.dataModel.tables = [{ id: 'table-1' }, { id: 'table-1' }]; }],
    ['data source', (snapshot) => { snapshot.dataModel.sources = [{ id: 'source-1' }, { id: 'source-1' }]; }],
    ['data relationship', (snapshot) => { snapshot.dataModel.relationships = [{ id: 'relationship-1' }, { id: 'relationship-1' }]; }],
    ['data view', (snapshot) => { snapshot.dataModel.views = [{ id: 'view-1' }, { id: 'view-1' }]; }],
    ['query definition', (snapshot) => { snapshot.queryDefinitions = [{ id: 'query-1' }, { id: 'query-1' }]; }],
    ['cell style template', (snapshot) => { snapshot.cellStyleTemplates = [{ id: 'template-1' }, { id: 'template-1' }]; }],
    ['print document owner', (snapshot) => { snapshot.printDocuments = [{ sheetId: 'sheet-1' }, { sheetId: 'sheet-1' }]; }],
  ];

  for (const [owner, duplicate] of duplicateOwners) {
    const candidate = structuredClone(workbook.snapshot()) as unknown as Record<string, any>;
    duplicate(candidate);
    assert.throws(() => assertCanonicalWorkbookSnapshot(candidate as unknown as WorkbookSnapshot), new RegExp(owner));
    assert.throws(() => WorkbookModel.fromSnapshot(candidate as unknown as WorkbookSnapshot), new RegExp(owner));
  }
});

test('canonical snapshots and model replacement reject duplicate defined-name owners', () => {
  const workbook = new WorkbookModel('unit-defined-name-identity', 'Defined names');
  const candidate = structuredClone(workbook.snapshot());
  const duplicateOwners = [
    [
      { name: 'TaxRate', formula: '0.1', scope: 'workbook' as const },
      { name: 'taxrate', formula: '0.2', scope: 'workbook' as const },
    ],
    [
      { name: 'LocalRate', formula: '0.1', scope: 'sheet' as const, sheetId: 'sheet-1' },
      { name: 'localrate', formula: '0.2', scope: 'sheet' as const, sheetId: 'sheet-1' },
    ],
  ];

  workbook.setDefinedName({ name: 'Preserved', formula: '0.3', scope: 'workbook' });
  for (const definitions of duplicateOwners) {
    candidate.definedNameModels = definitions;
    candidate.definedNames = {};
    assert.throws(() => assertCanonicalWorkbookSnapshot(candidate), /duplicate defined-name identity/);
    assert.throws(() => WorkbookModel.fromSnapshot(candidate), /duplicate defined-name identity/);
    assert.throws(() => workbook.replaceDefinedNames(definitions), /Defined-name owner identity is duplicated/);
    assert.equal(workbook.getDefinedNameExact('Preserved', 'workbook')?.formula, '0.3');
  }

  const legacyProjection = structuredClone(workbook.snapshot()) as unknown as Record<string, any>;
  delete legacyProjection.definedNameModels;
  legacyProjection.definedNames = { TaxRate: '0.1', taxrate: '0.2' };
  assert.throws(() => assertCanonicalWorkbookSnapshot(legacyProjection as unknown as WorkbookSnapshot), /definedNames projection contains duplicate identity/);
  assert.throws(() => WorkbookModel.fromSnapshot(legacyProjection as unknown as WorkbookSnapshot), /definedNames projection contains duplicate identity/);

  const staleProjection = structuredClone(workbook.snapshot());
  staleProjection.definedNameModels = [];
  staleProjection.definedNames = { Stale: '0.1' };
  assert.throws(() => assertCanonicalWorkbookSnapshot(staleProjection), /does not match canonical definedNameModels/);

  const whitespaceFormula = structuredClone(workbook.snapshot());
  whitespaceFormula.definedNameModels = [{ name: 'TaxRate', formula: '\u00a00.1', scope: 'workbook' }];
  whitespaceFormula.definedNames = {};
  assert.throws(() => assertCanonicalWorkbookSnapshot(whitespaceFormula), /identity or formula is invalid/);

  workbook.replaceDefinedNames([
    { name: 'Shared', formula: '0.1', scope: 'workbook' },
    { name: 'Shared', formula: '0.2', scope: 'sheet', sheetId: 'sheet-1' },
  ]);
  assert.equal(workbook.definedNameModels.length, 2);
});

test('defined-name projection preserves the legal __proto__ owner key', () => {
  const workbook = new WorkbookModel('unit-defined-name-prototype-key', 'Defined names');
  workbook.setDefinedName({ name: '__proto__', formula: '0.1', scope: 'workbook' });

  const snapshot = workbook.snapshot();

  assert.equal(Object.prototype.hasOwnProperty.call(snapshot.definedNames, '__proto__'), true);
  assert.equal(snapshot.definedNames?.['__proto__'], '0.1');
  assertCanonicalWorkbookSnapshot(snapshot);
});

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

test('canonical snapshots reject malformed chart linked formulas', () => {
  const snapshot = new WorkbookModel('unit-invalid-chart-formula', 'Invalid chart formula').snapshot();
  const sheet = snapshot.sheets[0]!;
  sheet.drawingPayloads['chart-1'] = {
    kind: 'chart',
    chartId: 'chart-1',
    chartType: 'line',
    subtype: 'line',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
    elements: { hiddenData: 'show', titleText: { linkedFormula: 7 as unknown as string } },
  };
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /is not a non-empty formula/);
});

test('canonical chart text formulas must resolve to one cell owner', () => {
  const snapshot = new WorkbookModel('unit-chart-formula-reference', 'Chart formula reference').snapshot();
  const sheet = snapshot.sheets[0]!;
  sheet.drawingPayloads['chart-1'] = {
    kind: 'chart',
    chartId: 'chart-1',
    chartType: 'line',
    subtype: 'line',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
    elements: { hiddenData: 'show', titleText: { linkedFormula: '=A1&B1' } },
  };
  const before = structuredClone(snapshot);
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), {
    name: 'Error',
    message: 'UNSUPPORTED_FEATURE: chart text formula titleText.linkedFormula on chart-1 is not a resolvable single-cell reference: formula root binary-expression is not a single cell reference',
  });
  assert.deepEqual(snapshot, before, 'rejecting a chart text expression cannot change the authored payload');
});

test('canonical chart text formulas preserve a structural #REF! result', () => {
  const snapshot = new WorkbookModel('unit-chart-invalid-reference', 'Chart invalid reference').snapshot();
  const sheet = snapshot.sheets[0]!;
  sheet.drawingPayloads['chart-1'] = {
    kind: 'chart',
    chartId: 'chart-1',
    chartType: 'line',
    subtype: 'line',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
    elements: { hiddenData: 'show', titleText: { linkedFormula: '=#REF!', text: 'stale cached title' } },
  };
  assert.doesNotThrow(() => assertCanonicalWorkbookSnapshot(snapshot));
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

test('canonical snapshots require Sheet Table columns to match the range width', () => {
  const snapshot = new WorkbookModel('unit-sheet-table-contract', 'Sheet Table contract').snapshot();
  const sheet = snapshot.sheets[0]!;
  sheet.sheetTables = [{
    id: 'table-1', sheetId: sheet.id, name: 'Table1',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 1, endColumn: 2 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false,
    showFirstColumn: false, showLastColumn: false, showFilterButton: false, autoExpand: 'none',
    columns: [{ id: 'column-1', name: 'A' }, { id: 'column-2', name: 'B' }],
  }];
  assert.doesNotThrow(() => assertCanonicalWorkbookSnapshot(snapshot));

  sheet.sheetTables[0]!.columns.pop();
  assert.throws(() => assertCanonicalWorkbookSnapshot(snapshot), /Sheet Table columns must match its range width/);
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

  const moved = matrix.get(5, 2)!;
  matrix.delete(5, 2);
  matrix.set(7, 2, moved);
  assert.deepEqual(matrix.getRegion(5, 7, 2, 2).map(({ row }) => row), [7]);
});

test('CellMatrix visits all persisted sparse cells without hydration and propagates reader failures', () => {
  const input = {
    '0': { '0': { value: 10 }, '1': { value: null, formula: '=A1' } },
    '600000': { '16383': { value: 'tail' } },
  };
  const matrix = new CellMatrix();
  matrix.deferJSON(input);
  const before = structuredClone(input);
  const revision = matrix.revision;
  const read: Array<{ row: number; column: number; value: unknown; formula?: string }> = [];
  matrix.forEachWithoutHydration((cell, row, column) => read.push({ row, column, value: cell.value,
    ...(cell.formula === undefined ? {} : { formula: cell.formula }) }));
  assert.deepEqual(read, [
    { row: 0, column: 0, value: 10 },
    { row: 0, column: 1, value: null, formula: '=A1' },
    { row: 600_000, column: 16_383, value: 'tail' },
  ]);
  assert.equal(matrix.isHydrated, false);
  assert.equal(matrix.revision, revision);
  const readerError = new Error('reader rejected input');
  assert.throws(() => matrix.forEachWithoutHydration(() => { throw readerError; }), (error) => error === readerError);
  assert.equal(matrix.isHydrated, false);
  assert.deepEqual(input, before);

  const hydrated = CellMatrix.fromJSON(input);
  const visited: string[] = [];
  hydrated.forEachWithoutHydration((_cell, row, column) => visited.push(`${row}:${column}`));
  assert.deepEqual(visited, ['0:0', '0:1', '600000:16383']);

  const addresses: string[] = [];
  matrix.forEachWithoutHydration((_cell, row, column) => {
    addresses.push(`${row}:${column}`);
    // A caller can explicitly request materialization during enumeration.
    if (row === 0 && column === 0) matrix.get(0, 0);
  });
  assert.deepEqual(addresses, visited);
  assert.equal(matrix.isHydrated, true);
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

test('CellMatrix rewrites deferred formula owners without hydrating or mutating source JSON', () => {
  const deferred = {
    '2': {
      '3': { value: null, formula: '=Sales[Amount]' },
      '4': { value: 'unchanged' },
    },
  };
  const matrix = new CellMatrix();
  matrix.deferJSON(deferred);
  const current = matrix.getFormulaOwnerWithoutHydration(2, 3)!;

  assert.equal(matrix.replaceFormulaOwnerWithoutHydration(2, 3, { ...current, formula: '=Orders[Amount]' }), true);

  assert.equal(matrix.isHydrated, false);
  assert.equal(matrix.revision, 3);
  assert.equal(matrix.toJSON()['2']?.['3']?.formula, '=Orders[Amount]');
  assert.equal(deferred['2']?.['3']?.formula, '=Sales[Amount]');
  assert.equal(matrix.replaceFormulaOwnerWithoutHydration(2, 4, { value: 'changed' }), false);
});

test('CellMatrix applies sparse additions, replacements and deletions without loading an inactive sheet', () => {
  const input = {
    '2': { '3': { value: 10 }, '4': { value: 'untouched' } },
    '900000': { '16383': { value: 'tail' } },
  };
  const original = structuredClone(input);
  Object.freeze(input['2']);
  Object.freeze(input['900000']);
  Object.freeze(input);
  const matrix = new CellMatrix();
  matrix.deferJSON(input);
  let revision = matrix.revision;
  assert.equal(matrix.count(), 3);
  matrix.set(2, 3, { value: 20, style: { fontFamily: ' arial ' } });
  assert.equal(matrix.revision, ++revision);
  matrix.set(2, 5, { value: 'new column' });
  assert.equal(matrix.revision, ++revision);
  matrix.set(1, 0, { value: 'new row' });
  assert.equal(matrix.revision, ++revision);
  matrix.delete(900_000, 16_383);
  assert.equal(matrix.revision, ++revision);
  matrix.delete(2, 3);
  assert.equal(matrix.revision, ++revision);
  matrix.delete(2, 3);
  assert.equal(matrix.revision, revision, 'deleting an absent cell is not a content change');
  assert.equal(matrix.count(), 3);
  assert.equal(matrix.isHydrated, false);
  assert.deepEqual(matrix.occupiedRange('inactive'), {
    sheetId: 'inactive', startRow: 1, endRow: 2, startColumn: 0, endColumn: 5,
  });
  assert.equal(matrix.getWithoutHydration(2, 4), input['2']['4']);
  assert.deepEqual(input, original, 'the persisted input must remain unchanged');
  const serialized = matrix.toJSON();
  assert.equal(matrix.get(2, 5)?.value, 'new column');
  assert.equal(matrix.isHydrated, true);
  assert.equal(matrix.revision, revision, 'materialization cannot invalidate an unchanged content token');
  assert.deepEqual(matrix.toJSON(), serialized);
});

test('CellMatrix keeps empty-row recreation and deferred revision/count/bounds coherent', () => {
  const matrix = new CellMatrix();
  matrix.deferJSON({});
  assert.equal(matrix.revision, 0);
  assert.equal(matrix.count(), 0);
  for (let iteration = 0; iteration < 8; iteration += 1) {
    matrix.set(8, 4, { value: iteration });
    assert.equal(matrix.revision, iteration * 4 + 1);
    matrix.set(8, 7, { value: iteration });
    assert.equal(matrix.revision, iteration * 4 + 2);
    assert.equal(matrix.count(), 2);
    matrix.delete(8, 4);
    assert.equal(matrix.revision, iteration * 4 + 3);
    matrix.delete(8, 7);
    assert.equal(matrix.revision, iteration * 4 + 4);
    assert.equal(matrix.count(), 0);
    assert.equal(matrix.isHydrated, false);
    assert.deepEqual(matrix.toJSON(), {});
  }
  assert.deepEqual(matrix.occupiedRange('inactive'), {
    sheetId: 'inactive', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0,
  });
  matrix.get(8, 4);
  assert.equal(matrix.revision, 32);
});

test('CellMatrix deferred writes reject invalid metadata before modifying payload, extent or revision', () => {
  const sheet = new WorksheetModel('inactive-write', 'Inactive', 10, 10);
  const input = { '2': { '3': { value: 'original' } } };
  sheet.cells.deferJSON(input);
  const revision = sheet.cells.revision;
  assert.throws(() => sheet.cells.set(200, 40, { value: 'invalid', style: { fontFamily: ' ' } }), /must not be empty/);
  assert.throws(() => sheet.cells.replaceCellWithoutHydration(2, 3, { value: 'invalid', style: { fontFamily: ' ' } }), /must not be empty/);
  assert.equal(sheet.cells.revision, revision);
  assert.deepEqual(sheet.cells.toJSON(), input);
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(sheet.rowCount, 10);
  assert.equal(sheet.columnCount, 10);
  sheet.cells.set(200, 40, { value: 'valid' });
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(sheet.rowCount, 201);
  assert.equal(sheet.columnCount, 41);
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
  const replacement = {
    id: 'region-2', sourceId: 'source-1',
    range: { sheetId: sheet.id, startRow: 60, endRow: 70, startColumn: 10, endColumn: 30 },
    headerRow: 60, revision: 0,
  };
  sheet.replaceDataRegions([replacement]);
  assert.equal(sheet.usedRange.endRow, 70);
  assert.throws(() => sheet.replaceDataRegions([replacement, { ...replacement, range: { ...replacement.range, startRow: 80, endRow: 90 } }]), /already exists/);
  assert.deepEqual(sheet.dataRegions, [replacement]);
  assert.equal(sheet.usedRange.endRow, 70);
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

test('sheet duplication allocates workbook-unique table identity and rewrites copied structured references', () => {
  const workbook = new WorkbookModel('sheet-table-identity', 'Sheet Table identity');
  const source = workbook.getSheet('sheet-1');
  source.sheetTables.push({
    id: 'sales-table', sheetId: source.id, name: 'Sales',
    range: { sheetId: source.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
    showFirstColumn: false, showLastColumn: false, showFilterButton: false, autoExpand: 'none',
    columns: [{ id: 'sales-item', name: 'Item' }, { id: 'sales-amount', name: 'Amount' }],
  });
  source.cells.set(0, 2, { value: null, formula: '=SUM(Sales[Amount])' });

  const duplicate = workbook.duplicateSheet(source.id, 'sheet-copy', 'Copy Sheet');

  assert.equal(duplicate.sheetTables[0]?.name, 'Sales_2');
  assert.equal(duplicate.cells.get(0, 2)?.formula, '=SUM(Sales_2[Amount])');
  assert.equal(source.cells.get(0, 2)?.formula, '=SUM(Sales[Amount])');
});

test('Sheet Table rename plans complete cell and preserved-source formula owner deltas', () => {
  const workbook = new WorkbookModel('sheet-table-rename-owners', 'Sheet Table rename owners');
  const sheet = workbook.getSheet('sheet-1');
  sheet.sheetTables.push({
    id: 'sales-table', sheetId: sheet.id, name: 'Sales',
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
    showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'none',
    columns: [{ id: 'sales-amount', name: 'Amount' }],
  });
  const deferredCells = {
    '0': { '2': { value: null, formula: '=SUM(Sales[Amount])' } },
    '1': { '2': {
      value: null,
      formulaMetadata: { kind: 'dataTable' as const, range: 'C2:D4', preservedOnly: true, sourceFormula: '=Sales[Amount]', reason: 'Native OOXML formula' },
    } },
  };
  sheet.cells.deferJSON(deferredCells);

  const plan = planSheetTableRename(workbook, 'sales-table', 'Orders');
  const effect = plan.apply();

  const cells = sheet.cells.toJSON();
  assert.equal(cells['0']?.['2']?.formula, '=SUM(Orders[Amount])');
  assert.equal(cells['1']?.['2']?.formulaMetadata?.sourceFormula, '=Orders[Amount]');
  assert.equal(cells['1']?.['2']?.formulaMetadata?.range, 'C2:D4');
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(deferredCells['0']?.['2']?.formula, '=SUM(Sales[Amount])');
  assert.equal(deferredCells['1']?.['2']?.formulaMetadata?.sourceFormula, '=Sales[Amount]');
  assert.equal(effect.formulaOwnerDeltas?.length, 2);
  assert.equal(effect.formulaOwnerDeltas?.[0]?.kind, 'formula-cell');
  assert.equal(effect.formulaOwnerDeltas?.[1]?.kind, 'formula-cell');
});

test('Sheet Table rename still rejects shared formula groups without changing their owner', () => {
  const workbook = new WorkbookModel('sheet-table-rename-shared-formula', 'Sheet Table rename shared formula');
  const sheet = workbook.getSheet('sheet-1');
  sheet.sheetTables.push({
    id: 'sales-table', sheetId: sheet.id, name: 'Sales',
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
    showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'none',
    columns: [{ id: 'sales-amount', name: 'Amount' }],
  });
  const owner = {
    value: null,
    formula: '=Sales[Amount]',
    formulaMetadata: { kind: 'shared' as const, sharedIndex: 7, sharedMaster: true, range: 'C1:C2', sourceFormula: '=Sales[Amount]' },
  };
  sheet.cells.set(0, 2, owner);

  assert.throws(() => planSheetTableRename(workbook, 'sales-table', 'Orders'), /formula group/);
  assert.equal(sheet.sheetTables[0]?.name, 'Sales');
  assert.deepEqual(sheet.cells.get(0, 2), owner);
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

test('sheet lifecycle restores only its own deferred cells while preserving other scoped names', () => {
  const workbook = new WorkbookModel('unit-sheet-lifecycle-lazy', 'Lifecycle');
  const second = workbook.addSheet('sheet-2', 'Second');
  second.cells.set(100, 4, { value: 'sparse' });
  workbook.setDefinedName({ name: 'FirstRate', formula: '0.1', scope: 'sheet', sheetId: 'sheet-1' });
  workbook.setDefinedName({ name: 'SecondRate', formula: '0.2', scope: 'sheet', sheetId: 'sheet-2' });
  const saved = workbook.getSheetSnapshot(second.id);
  workbook.removeSheet(second.id);
  assert.equal(workbook.getDefinedNameExact('SecondRate', 'sheet', second.id), undefined);

  workbook.restoreSheetSnapshot(saved);
  assert.equal(workbook.getDefinedNameExact('FirstRate', 'sheet', 'sheet-1')?.formula, '0.1');
  assert.equal(workbook.getDefinedNameExact('SecondRate', 'sheet', second.id)?.formula, '0.2');
  assert.equal(workbook.getSheet(second.id).cells.isHydrated, false);
  assert.equal(workbook.getSheet(second.id).cells.count(), 1);
  assert.equal(workbook.getSheet(second.id).cells.isHydrated, false);
  assert.equal(workbook.getSheet(second.id).cells.get(100, 4)?.value, 'sparse');
});

test('sheet lifecycle rejects invalid owners and indexes without partial restoration', () => {
  const workbook = new WorkbookModel('unit-sheet-lifecycle-atomic', 'Lifecycle');
  const second = workbook.addSheet('sheet-2', 'Second');
  workbook.setDefinedName({ name: 'Preserved', formula: '0.1', scope: 'sheet', sheetId: 'sheet-1' });
  const saved = workbook.getSheetSnapshot(second.id);
  workbook.removeSheet(second.id);
  const before = workbook.snapshot();
  const assertUnchanged = () => assert.deepEqual(workbook.snapshot(), before);

  assert.throws(() => workbook.restoreSheetSnapshot(saved, 0.5), /Invalid sheet restore index/);
  assertUnchanged();
  const wrongOwner = structuredClone(saved);
  wrongOwner.lifecycleDefinedNames = [{ name: 'Foreign', formula: '0.2', scope: 'sheet', sheetId: 'sheet-1' }];
  assert.throws(() => workbook.restoreSheetSnapshot(wrongOwner), /defined-name owner does not match worksheet/);
  assertUnchanged();
  const duplicateOwner = structuredClone(saved);
  duplicateOwner.lifecycleDefinedNames = [
    { name: 'Rate', formula: '0.2', scope: 'sheet', sheetId: second.id },
    { name: 'rate', formula: '0.3', scope: 'sheet', sheetId: second.id },
  ];
  assert.throws(() => workbook.restoreSheetSnapshot(duplicateOwner), /duplicate defined-name identity/);
  assertUnchanged();
  const wrongDocument = structuredClone(saved);
  wrongDocument.lifecyclePrintDocument = {
    schema: 'PrintDocument', unitId: workbook.unitId, sheetId: 'sheet-1',
    pageSetup: { paperSize: 'a4', orientation: 'portrait', scale: 100,
      margins: { top: 1, right: 1, bottom: 1, left: 1, header: 1, footer: 1 },
      printGridlines: false, printHeadings: false, centerHorizontally: false, centerVertically: false },
    printAreas: [], pageBreaks: [],
  };
  assert.throws(() => workbook.restoreSheetSnapshot(wrongDocument), /print document owner does not match worksheet/);
  assertUnchanged();
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
