import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel, canonicalRecordFieldFormula, assertRecordRelationship, type WorkbookTableModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerRecordCommands } from './commands';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { synchronizeRecordCalculations, recordSheetTables } from './record-calculation';
import { createRemoteReadySessionFixture } from '../../session-test-fixtures';
import { hydrateRuntime, startCollaborationSession } from '../../runtime';
import { CollabSocketClient } from '@react-sheets/protocol';
import { refreshExternalLinks } from './external-link-host';

function fixture() {
  const workbook = new WorkbookModel('records', 'Records');
  const orders = workbook.getSheet(workbook.primarySheetId); orders.name = 'Orders';
  const customers = workbook.addSheet('customers-sheet', 'Customers');
  [['ID', 'Customer', 'Quantity', 'Price', 'Amount'], ['o1', 'c1', 2, 10, null], ['o2', 'c1', 3, 20, null], ['o3', 'c2', 4, 5, null]].forEach((row, r) => row.forEach((value, c) => orders.cells.set(r, c, { value })));
  [['ID', 'Name', 'Amounts', 'Total'], ['c1', 'Alpha', null, null], ['c2', 'Beta', null, null]].forEach((row, r) => row.forEach((value, c) => customers.cells.set(r, c, { value })));
  const table = (id: string, sheetId: string, names: string[], count: number): WorkbookTableModel => ({ id, name: id, sourceSheetId: sheetId, sourceRange: { sheetId, startRow: 0, endRow: count, startColumn: 0, endColumn: names.length - 1 }, rowCount: count, recordIdFieldId: `${id}-id`, fields: names.map((name, ordinal) => ({ id: `${id}-${name.toLowerCase()}`, name, ordinal, type: ordinal < 2 ? 'text' : 'number' })), blockSize: 1024, blocks: [], revision: 0 });
  const orderTable = table('Orders', orders.id, ['ID', 'Customer', 'Quantity', 'Price', 'Amount'], 3);
  orderTable.fields[4]!.calculation = { kind: 'formula', formula: canonicalRecordFieldFormula(orderTable, '=[Quantity]*[Price]') };
  workbook.addTable(orderTable); workbook.addTable(table('Customers', customers.id, ['ID', 'Name', 'Amounts', 'Total'], 2));
  const relation = { id: 'orders-customers', fromTableId: 'Orders', fromFieldId: 'Orders-customer', toTableId: 'Customers', toFieldId: 'Customers-id', cardinality: 'many-to-one' as const };
  assertRecordRelationship(workbook, relation); workbook.dataModel.relationships.set(relation.id, relation);
  workbook.getTable('Customers').fields[2]!.calculation = { kind: 'lookup', relationshipId: relation.id, targetFieldId: 'Orders-amount', direction: 'reverse' };
  workbook.getTable('Customers').fields[3]!.calculation = { kind: 'rollup', relationshipId: relation.id, targetFieldId: 'Orders-amount', direction: 'reverse', aggregate: 'SUM' };
  const engine = new FormulaEngine({ defaultSheetId: orders.id, sheetOrder: workbook.getSheets().map(({ id, name }) => ({ id, name })) });
  engine.setSheetTables(recordSheetTables(workbook));
  for (const sheet of workbook.getSheets()) sheet.cells.forEach((cell, row, column) => engine.setValue({ sheetId: sheet.id, row, column }, cell.value ?? null));
  synchronizeRecordCalculations(engine, workbook);
  return { workbook, orders, customers, engine };
}

test('Record formulas, reverse Lookup and Rollup share one dependency graph and native persistence', async () => {
  const { workbook, orders, engine } = fixture();
  await engine.recalculateAsync();
  assert.equal(engine.getRecordFieldResult('Orders', 'o1', 'Orders-amount')?.value, 20);
  assert.deepEqual(engine.getRecordFieldResult('Customers', 'c1', 'Customers-amounts')?.value, [[20], [60]]);
  assert.equal(engine.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 80);
  engine.setValue({ sheetId: orders.id, row: 1, column: 2 }, 5);
  assert.equal(engine.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 110);
  assert.deepEqual(WorkbookModel.fromSnapshot(workbook.snapshot()).getDataModel(), workbook.getDataModel());
  assert.equal(orders.cells.getWithoutHydration(1, 4)?.formula, undefined);
});
test('stable Record/Field writes reject missing relationship targets without modifying the model', () => {
  const { workbook, orders } = fixture();
  const commands = new CommandRuntime(workbook); registerRecordCommands(commands);
  const before = workbook.snapshot();
  assert.throws(() => commands.execute('record.set', { tableId: 'Orders', recordId: 'o1', fieldId: 'Orders-customer', value: 'missing' }), /missing/i);
  assert.deepEqual(workbook.snapshot(), before);
  assert.throws(() => commands.execute('record.set', { tableId: 'Orders', recordId: 'o1', fieldId: 'Orders-amount', value: 123 }), /read.only/i);
  commands.execute('record.set', { tableId: 'Orders', recordId: 'o2', fieldId: 'Orders-quantity', value: 7 });
  assert.equal(orders.cells.getWithoutHydration(2, 2)?.value, 7);
});

test('3D endpoint deletion has reversible formula facts across canonical command undo and redo', () => {
  const workbook = new WorkbookModel('lifecycle', 'Lifecycle');
  const first = workbook.getSheet(workbook.primarySheetId); first.name = 'Jan';
  workbook.addSheet('feb', 'Feb'); workbook.addSheet('mar', 'Mar');
  const summary = workbook.addSheet('summary', 'Summary');
  summary.cells.set(0, 0, { value: null, formula: '=SUM(Jan:Mar!A1)' });
  summary.cells.set(1, 0, { value: null, formula: '=Jan!A1' });
  const commands = new CommandRuntime(workbook); registerSheetCommands(commands);
  commands.execute('sheet.remove', { id: first.id });
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Feb:Mar!A1)');
  assert.equal(summary.cells.get(1, 0)?.formula, '=#REF!');
  commands.undo();
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Jan:Mar!A1)');
  assert.equal(summary.cells.get(1, 0)?.formula, '=Jan!A1');
  commands.redo();
  assert.equal(summary.cells.get(0, 0)?.formula, '=SUM(Feb:Mar!A1)');
});

test('a committed deletion ACK recalculates surviving owners without addressing the removed sheet', async () => {
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    const first = runtime.model.getSheet(runtime.model.primarySheetId); first.name = 'Jan';
    first.cells.set(0, 0, { value: 10 });
    runtime.model.addSheet('feb', 'Feb').cells.set(0, 0, { value: 20 });
    runtime.model.addSheet('mar', 'Mar').cells.set(0, 0, { value: 30 });
    runtime.model.addSheet('summary', 'Summary').cells.set(0, 0, { value: null, formula: '=SUM(Jan:Mar!A1)' });
    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    let facts: import('@react-sheets/command-runtime').MutationInfo | undefined;
    const detach = runtime.commands.onMutation((mutation, source) => { if (source === 'command') facts = mutation; });
    const result = runtime.commands.execute('sheet.remove', { id: first.id }); detach();
    assert.ok(facts);
    runtime.commands.applyCommittedStructuralPatches(result.operationId, [{ ...facts, structuralDefinedNameOwnerDeltas: [], structuralRangeOwnerDeltas: [] }], 1);
    await app.waitForFormulaCalculation();
    assert.equal(runtime.formula.getCellResult({ sheetId: 'summary', row: 0, column: 0 })?.value, 50);
  } finally { app.dispose(); }
});

 test('runtime loading preserves calculated field owners over empty physical source cells', async () => {
  const { workbook } = fixture();
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    hydrateRuntime(runtime, { snapshot: workbook.snapshot(), revision: 0 });
    await app.waitForFormulaCalculation();
    assert.equal(runtime.formula.getRecordFieldResult('Orders', 'o2', 'Orders-amount')?.value, 60);
    assert.equal(runtime.formula.getRecordFieldResult('Customers', 'c1', 'Customers-total')?.value, 80);
  } finally { app.dispose(); }
});

test('TableSheet projections follow calculated source owners after engine replacement', async () => {
  const { workbook } = fixture();
  const view = workbook.addSheet('orders-view', 'Orders View');
  view.kind = 'table-sheet';
  view.tableSheet = { viewId: 'Orders', columns: [{ fieldId: 'Orders-amount', caption: 'Amount', type: 'number' }], grouping: [] };
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    hydrateRuntime(runtime, { snapshot: workbook.snapshot(), revision: 0 });
    await app.waitForFormulaCalculation();
    const projection = app['projection'];
    const owner = runtime.model.getSheet('orders-view');
    const before = projection.getCanvasProjection(owner);
    assert.equal(before.getCell(1, 0)?.displayValue, '20');
    const table = structuredClone(runtime.model.getTable('Orders'));
    table.fields[4]!.calculation = { kind: 'formula', formula: canonicalRecordFieldFormula(table, '=[Quantity]*[Price]*2') };
    runtime.commands.execute('table.configure', { table });
    await app.waitForFormulaCalculation();
    const after = projection.getCanvasProjection(owner);
    assert.notEqual(after, before);
    assert.equal(after.getCell(1, 0)?.displayValue, '40');
    const invalid = structuredClone(table); invalid.fields[4]!.id = 'replacement';
    assert.throws(() => runtime.commands.execute('table.configure', { table: invalid }), /IDENTITY_IMMUTABLE/);
  } finally { app.dispose(); }
});

test('Record undo restores sparse cell absence and existing cell metadata exactly', () => {
  const { workbook, orders } = fixture();
  orders.cells.delete(2, 2);
  const commands = new CommandRuntime(workbook); registerRecordCommands(commands);
  const before = workbook.snapshot();
  commands.execute('record.set', { tableId: 'Orders', recordId: 'o2', fieldId: 'Orders-quantity', value: 7 });
  commands.undo();
  assert.deepEqual(workbook.snapshot(), before);
  commands.redo();
  assert.equal(orders.cells.getWithoutHydration(2, 2)?.value, 7);
});

test('Record tables reject overlapping source owners before changing the workbook', () => {
  const { workbook } = fixture();
  const before = workbook.snapshot();
  const overlapping = structuredClone(workbook.getTable('Orders'));
  overlapping.id = 'overlapping-orders';
  assert.throws(() => workbook.addTable(overlapping), /one table owner/);
  assert.deepEqual(workbook.snapshot(), before);
  assert.deepEqual(WorkbookModel.fromSnapshot(before).getDataModel(), workbook.getDataModel());
});

test('external refresh coalesces within one model and permits a fresh engine to refresh immediately', async () => {
  const app = createRemoteReadySessionFixture();
  try {
    const runtime = app['runtime'];
    const binding = { id: 'source-link', token: 'Source.xlsx', sourceUnitId: 'source', sheets: [{ token: 'Sheet1', sheetId: 'sheet-1' }] };
    runtime.model.dataModel.externalLinks.set(binding.id, binding);
    const source = new WorkbookModel('source', 'Source'); source.getSheet('sheet-1').cells.set(0, 0, { value: 25 });
    const result = { schema: 'ExternalCalculationGraph' as const, rootUnitId: runtime.model.unitId, subject: 'reader', nodes: [
      { unitId: runtime.model.unitId, state: 'connected' as const, snapshot: runtime.model.snapshot(), revision: 0, accessRevision: 0, blockedRanges: [] },
      { unitId: source.unitId, state: 'connected' as const, snapshot: source.snapshot(), revision: 1, accessRevision: 0, blockedRanges: [] },
    ] };
    let resolveOld!: (value: typeof result) => void;
    let calls = 0;
    runtime.api.getExternalCalculationGraph = async () => {
      calls += 1;
      if (calls === 1) return new Promise(resolve => { resolveOld = resolve; });
      return result;
    };
    const old = refreshExternalLinks(runtime);
    assert.equal(refreshExternalLinks(runtime), old);
    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    const current = refreshExternalLinks(runtime);
    assert.notEqual(current, old);
    await current;
    assert.equal(calls, 2);
    assert.equal(runtime.formula.getExternalCalculationLinks()[0]?.state, 'connected');
    resolveOld({ ...result, nodes: result.nodes.map(node => ({ ...node, revision: 0 })) }); await old;
    assert.equal(runtime.formula.getExternalCalculationLinks()[0]?.sourceRevision, 1);
  } finally { app.dispose(); }
});

test('initial collaboration preserves a valid access projection without a persisted cache, and purges a changed projection', async t => {
  const priorWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  Object.defineProperty(globalThis, 'window', { configurable: true, value: { location: { protocol: 'http:', host: 'localhost' }, setInterval, clearInterval } });
  t.after(() => { if (priorWindow) Object.defineProperty(globalThis, 'window', priorWindow); else Reflect.deleteProperty(globalThis, 'window'); });
  let onStatus!: (status: 'connecting' | 'open' | 'closed') => void;
  t.mock.method(CollabSocketClient.prototype, 'onStatus', (listener: typeof onStatus) => { onStatus = listener; return () => {}; });
  t.mock.method(CollabSocketClient.prototype, 'open', () => { onStatus('open'); });
  t.mock.method(CollabSocketClient.prototype, 'send', () => true);
  t.mock.method(CollabSocketClient.prototype, 'markSynchronized', () => {});
  t.mock.method(CollabSocketClient.prototype, 'close', () => {});
  for (const accessRevision of [0, 1]) {
    const app = createRemoteReadySessionFixture();
    let detach: (() => void) | undefined;
    try {
      const runtime = app['runtime'];
      runtime.model.addSheet('selected', 'Selected');
      const snapshot = runtime.model.snapshot();
      runtime.workspaceRecord = null;
      runtime.accessProjection = { unitId: snapshot.unitId, role: 'owner', accessRevision: 0, regions: [] };
      runtime.api.getSnapshot = async () => ({ unitId: snapshot.unitId, snapshot, revision: 0 });
      runtime.api.getAccess = async () => ({ unitId: snapshot.unitId, role: 'owner', accessRevision, regions: [] });
      runtime.api.listRevisions = async () => [];
      const resetSheets: string[] = [];
      const notices: string[] = [];
      runtime.handlers.onActiveSheetChange = id => resetSheets.push(id);
      runtime.handlers.onNotice = message => notices.push(message);
      const synchronized = new Promise<void>(resolve => { runtime.handlers.onSaveState = state => { if (state === 'saved') resolve(); }; });
      detach = startCollaborationSession(runtime, () => 'selected:0:0');
      await synchronized;
      assert.equal(runtime.remoteConnected, true, notices.join('; '));
      assert.deepEqual(resetSheets, accessRevision === 0 ? [] : [runtime.model.primarySheetId]);
      assert.equal(runtime.accessProjection.accessRevision, accessRevision);
    } finally { detach?.(); app.dispose(); }
  }
});

test('one authorized multi-workbook graph evaluates shared dependencies and propagates leaf faults through COUNT and IFERROR', async () => {
  const { calculateExternalGraph } = await import('./external-link-host');
  const { validateExternalCalculationGraph } = await import('@react-sheets/protocol');
  const a = new WorkbookModel('dag-a', 'A'), b = new WorkbookModel('dag-b', 'B'), c = new WorkbookModel('dag-c', 'C');
  const bind = (source: WorkbookModel, token: string) => ({ id: token, token, sourceUnitId: source.unitId, sheets: [{ token: source.getSheet(source.primarySheetId).name, sheetId: source.primarySheetId }] });
  a.getSheet(a.primarySheetId).cells.set(0, 0, { value: 10 });
  const ab = bind(a, 'A.xlsx'), bc = bind(b, 'B.xlsx');
  b.dataModel.externalLinks.set(ab.id, ab); c.dataModel.externalLinks.set(bc.id, bc);
  b.getSheet(b.primarySheetId).cells.set(0, 0, { value: null, formula: "='[A.xlsx]Sheet1'!A1*2" });
  const node = (book: WorkbookModel) => ({ unitId: book.unitId, state: 'connected' as const, snapshot: book.snapshot(), revision: 1, accessRevision: 0, blockedRanges: [] });
  const graph = { schema: 'ExternalCalculationGraph' as const, rootUnitId: c.unitId, subject: 'reader', nodes: [node(c), node(b), node(a)] };
  const engine = new FormulaEngine({ defaultSheetId: c.primarySheetId });
  for (const [column, formula] of ["=SUM('[B.xlsx]Sheet1'!A1)", "=COUNT('[B.xlsx]Sheet1'!A1)", "=IFERROR('[B.xlsx]Sheet1'!A1,99)"].entries()) engine.setFormula({ sheetId: c.primarySheetId, row: 0, column }, formula);
  const read = () => [0, 1, 2].map(column => engine.getCellResult({ sheetId: c.primarySheetId, row: 0, column })!.value);
  try {
    engine.applyExternalCalculationLinks(await calculateExternalGraph(graph, [bc])); await engine.recalculateAsync();
    assert.deepEqual(read(), [20, 1, 20]);
    const faultGraph = validateExternalCalculationGraph({ ...graph, nodes: [node(c), node(b), { unitId: a.unitId, state: 'denied', error: { code: 'FORBIDDEN', message: 'Leaf revoked' } }] }, c.unitId);
    engine.applyExternalCalculationLinks(await calculateExternalGraph(faultGraph, [bc])); await engine.recalculateAsync();
    for (const value of read()) assert.equal((value as { code: string }).code, '#BLOCKED!');
    a.getSheet(a.primarySheetId).cells.set(0, 0, { value: 40 });
    engine.applyExternalCalculationLinks(await calculateExternalGraph({ ...graph, nodes: [node(c), node(b), node(a)] }, [bc])); await engine.recalculateAsync();
    assert.deepEqual(read(), [80, 1, 80]);
    const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
    try { await worker.recalculateAsync(undefined, undefined, true); assert.deepEqual([0, 1, 2].map(column => worker.getCellResult({ sheetId: c.primarySheetId, row: 0, column })!.value), [80, 1, 80]); }
    finally { worker.disposeCalculationTasks(); }
  } finally { engine.disposeCalculationTasks(); }
});

test('external calculation graph rejects missing duplicate cyclic and malformed nodes instead of partial input', async () => {
  const { validateExternalCalculationGraph, decodeClientOperationMessage, encodeClientOperationMessage } = await import('@react-sheets/protocol');
  const a = new WorkbookModel('invalid-a', 'A'), b = new WorkbookModel('invalid-b', 'B');
  const binding = { id: 'B.xlsx', token: 'B.xlsx', sourceUnitId: b.unitId, sheets: [{ token: 'Sheet1', sheetId: b.primarySheetId }] };
  a.dataModel.externalLinks.set(binding.id, binding);
  const node = (book: WorkbookModel) => ({ unitId: book.unitId, state: 'connected', snapshot: book.snapshot(), revision: 1, accessRevision: 0, blockedRanges: [] });
  const graph = { schema: 'ExternalCalculationGraph', rootUnitId: a.unitId, subject: 'reader', nodes: [node(a), node(b)] };
  assert.equal(validateExternalCalculationGraph(graph, a.unitId).nodes.length, 2);
  assert.throws(() => validateExternalCalculationGraph({ ...graph, schema: 'LegacyGraph' }, a.unitId));
  assert.throws(() => validateExternalCalculationGraph({ ...graph, subject: '' }, a.unitId));
  for (const nodes of [[node(a)], [node(a), node(b), node(b)], [node(a), { ...node(b), revision: -1 }], [node(a), { unitId: b.unitId, state: 'denied', snapshot: b.snapshot(), error: { code: 'FORBIDDEN', message: 'Denied' } }]]) {
    assert.throws(() => validateExternalCalculationGraph({ ...graph, nodes }, a.unitId));
  }
  assert.throws(() => validateExternalCalculationGraph({ ...graph, nodes: [node(a), { ...node(b), blockedRanges: [{ sheetId: 'missing', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] }] }, a.unitId), /RANGE_INVALID/);
  const oversized = b.snapshot(); oversized.sheets[0]!.rowCount = 100_001;
  oversized.sheets[0]!.cells = Object.fromEntries(Array.from({ length: 100_001 }, (_, row) => [String(row), { '0': { value: row } }]));
  assert.throws(() => validateExternalCalculationGraph({ ...graph, nodes: [node(a), { ...node(b), snapshot: oversized }] }, a.unitId), /INPUT_BUDGET/);
  b.dataModel.externalLinks.set('A.xlsx', { id: 'A.xlsx', token: 'A.xlsx', sourceUnitId: a.unitId, sheets: [{ token: 'Sheet1', sheetId: a.primarySheetId }] });
  assert.throws(() => validateExternalCalculationGraph({ ...graph, nodes: [node(a), node(b)] }, a.unitId), /CIRCULAR_DEPENDENCY/);
  assert.deepEqual(decodeClientOperationMessage(encodeClientOperationMessage({ type: 'calculation.subscribe', unitId: a.unitId })), { type: 'calculation.subscribe', unitId: a.unitId });
  assert.throws(() => decodeClientOperationMessage(JSON.stringify({ type: 'calculation.changed', unitId: a.unitId, sourceUnitId: b.unitId })), /Server-only/);
});
