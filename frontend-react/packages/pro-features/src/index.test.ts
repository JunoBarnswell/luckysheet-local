import test from 'node:test';
import assert from 'node:assert/strict';
import {
  paginateRange,
  serializeSnapshot,
  deserializeSnapshot,
  exportSnapshotToXlsxXml,
  parseXlsxXmlToSnapshot,
  computePivotTable,
  diffSnapshots,
} from './index';
import { ChangesetStateMachine } from './collaboration';
import { WorkbookModel } from '@react-sheets/core-model';

test('print pagination creates deterministic inclusive ranges', () => {
  const pages = paginateRange({ sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 }, 2, 3);
  assert.equal(pages.length, 6);
  assert.deepEqual(pages[0]?.range, { sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 });
});

test('changeset state machine enforces revision order', () => {
  const machine = new ChangesetStateMachine();
  machine.submit({
    schema: 'CollaborationChangeSetV1',
    operationId: 'op-1',
    unitId: 'unit-1',
    actorId: 'actor-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [],
    createdAt: new Date(0).toISOString(),
  });
  assert.equal(machine.acknowledge('op-1', 1).status, 'acknowledged');
  assert.throws(() =>
    machine.submit({
      schema: 'CollaborationChangeSetV1',
      operationId: 'op-2',
      unitId: 'unit-1',
      actorId: 'actor-1',
      clientSequence: 2,
      baseRevision: 0,
      mutations: [],
      createdAt: new Date(0).toISOString(),
    }),
  );
});

test('snapshot serializer rejects non-v1 payloads', () => {
  const snapshot = {
    schema: 'WorkbookSnapshotV1' as const,
    unitId: 'unit-1',
    name: 'Test',
    activeSheetId: 'sheet-1',
    sheets: [],
  };
  assert.deepEqual(deserializeSnapshot(serializeSnapshot(snapshot)), snapshot);
  assert.throws(() => deserializeSnapshot('{"schema":"legacy"}'));
});

test('xlsx export produces valid OpenXML structure and imports back', () => {
  const workbook = new WorkbookModel('unit-xlsx', 'XLSX Test');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Product' });
  sheet.cells.set(0, 1, { value: 'Sales' });
  sheet.cells.set(1, 0, { value: 'Widget' });
  sheet.cells.set(1, 1, { value: 150 });
  sheet.cells.set(2, 0, { value: 'Total' });
  sheet.cells.set(2, 1, { value: null, formula: '=SUM(B2:B2)' });

  const snapshot = workbook.snapshot();
  const xmlFiles = exportSnapshotToXlsxXml(snapshot);

  assert.ok(xmlFiles['[Content_Types].xml']);
  assert.ok(xmlFiles['xl/workbook.xml']);
  assert.ok(xmlFiles['xl/worksheets/sheet1.xml']);
  assert.ok(xmlFiles['xl/sharedStrings.xml']);

  const imported = parseXlsxXmlToSnapshot(xmlFiles);
  assert.equal(imported.sheets.length, 1);
  assert.equal(imported.sheets[0]?.name, 'Sheet1');
  assert.equal(imported.sheets[0]?.cells['1']?.['1']?.value, 150);
});

test('pivot engine aggregates source data by dimensions and value fields', () => {
  const workbook = new WorkbookModel('unit-pivot', 'Pivot Test');
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 'Region' });
  sheet.cells.set(0, 1, { value: 'Sales' });

  sheet.cells.set(1, 0, { value: 'North' });
  sheet.cells.set(1, 1, { value: 100 });

  sheet.cells.set(2, 0, { value: 'South' });
  sheet.cells.set(2, 1, { value: 200 });

  sheet.cells.set(3, 0, { value: 'North' });
  sheet.cells.set(3, 1, { value: 300 });

  const result = computePivotTable(workbook, {
    id: 'pivot-1',
    sheetId: 'sheet-1',
    sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    layout: {
      rows: [{ field: 'Region' }],
      columns: [],
      filters: [],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      showSubtotals: true,
      showGrandTotals: true,
      compact: false,
      repeatLabels: false,
    },
  });

  assert.deepEqual(result.headers, ['Region', 'SUM of Sales']);
  assert.equal(result.rows.length, 2);
  const northRow = result.rows.find((r) => r.keys[0] === 'North');
  assert.deepEqual(northRow?.values, [400]);
});

test('diffSnapshots identifies exact cell changes between revisions', () => {
  const wb1 = new WorkbookModel('unit-diff', 'Diff1');
  wb1.getSheet('sheet-1').cells.set(0, 0, { value: 'Original' });
  wb1.getSheet('sheet-1').cells.set(0, 1, { value: 100 });

  const wb2 = new WorkbookModel('unit-diff', 'Diff2');
  wb2.getSheet('sheet-1').cells.set(0, 0, { value: 'Modified' });
  wb2.getSheet('sheet-1').cells.set(0, 1, { value: 100 }); // unchanged

  const diff = diffSnapshots(wb1.snapshot(), wb2.snapshot());
  assert.equal(diff.length, 1);
  assert.equal(diff[0]?.row, 0);
  assert.equal(diff[0]?.column, 0);
  assert.equal(diff[0]?.oldValue, 'Original');
  assert.equal(diff[0]?.newValue, 'Modified');
});

test('changeset state machine handles reject and offline transitions', () => {
  const machine = new ChangesetStateMachine();
  machine.submit({
    schema: 'CollaborationChangeSetV1',
    operationId: 'op-fail',
    unitId: 'unit-1',
    actorId: 'actor-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [],
    createdAt: new Date(0).toISOString(),
  });
  const rejected = machine.reject('op-fail', 'Constraint violation');
  assert.equal(rejected.status, 'rejected');
  assert.equal(rejected.error, 'Constraint violation');

  const offline = machine.markOffline();
  assert.equal(offline.status, 'offline');
});
