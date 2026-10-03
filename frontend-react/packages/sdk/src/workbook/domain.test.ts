import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import assert from 'node:assert/strict';
import test from 'node:test';
import { getWorkbookObjectPort, WorkbookSession } from '@react-sheets/spreadsheet-app';
import { Workbook } from './workbook';
import { SdkError } from '../error';
import { CollabSocketClient, type OperationMessage } from '@react-sheets/protocol';
import { consumeBrowserCalculationTaskWithEngine, type CalculationBrowserWorker, type FormulaEngine } from '@react-sheets/formula-engine';
import { startCollaborationSession } from '../../../spreadsheet-app/src/runtime';
import { DataDomain } from '../data/domain';
function fixture(id: string, scope: object = {}) {
  const session = new WorkbookSession({ unitId: id });
  const data = new DataDomain(session);
  const workbook = new Workbook(getWorkbookObjectPort(session), scope, () => session.dispose(), current => data.actionsFor(current));
  return { session, workbook, sheet: session['runtime'].model.getSheet(session.getActiveSheetId()) };
}
const invalid = (cause: unknown) => cause instanceof SdkError && cause.code === 'INVALID_ARGUMENT' && Boolean(cause.object?.workbookId) && Boolean(cause.recovery);
test('public Workbook data actions share the canonical owner and retire captured actions on close', async () => {
  const { workbook } = fixture('public-data');
  const sheet = workbook.worksheets.at(0), actions = workbook.data;
  await sheet.ranges.get('A1:B3').setValues([['Group', 'Amount'], ['East', 2], ['East', 4]]);
  assert.deepEqual(await actions.subtotal({ range: { sheetId: sheet.id, address: 'A1:B3' }, functionName: 'PRODUCT' }), { status: 'applied' });
  assert.equal((await sheet.cells.get('B6').read()).calculatedValue, 8);
  await workbook.undo(); assert.equal((await sheet.cells.get('B6').read()).value, null);
  await workbook.redo(); assert.equal((await sheet.cells.get('B6').read()).calculatedValue, 8);
  workbook.close();
  const result = await actions.subtotal({ range: { sheetId: sheet.id, address: 'A1:B3' } });
  assert.equal(result.status, 'rejected');
  if (result.status === 'rejected') { assert.equal(result.error.code, 'RUNTIME_DISPOSED'); assert.equal(result.error.operation, 'data.subtotal'); assert.equal(result.error.object?.workbookId, 'public-data'); }
});
test('object cells own explicit addresses, preserve literal values/styles and share the canonical formula engine', async () => {
  const { session, workbook, sheet } = fixture('object-1');
  try {
    const worksheet = workbook.worksheets.at(0);
    assert.equal(workbook.worksheets.byId(worksheet.id), worksheet);
    assert.equal(workbook.worksheets.byName(worksheet.name.toLowerCase()), worksheet);
    const cell = worksheet.cells.get('$D$8');
    assert.equal(cell, worksheet.cells.get('d8'));
    sheet.cells.set(7, 3, { value: 9, style: { bold: true } });
    session.selectAddress('A1');
    await cell.setValue('=1+2');
    assert.equal((await cell.read()).value, '=1+2');
    assert.equal((await cell.read()).formula, undefined);
    assert.equal(sheet.cells.get(7, 3)?.style?.bold, true);
    await cell.setFormula('=SUM(1,2)');
    assert.equal((await cell.read()).calculatedValue, 3);
    assert.equal(sheet.cells.get(7, 3)?.formula, '=SUM(1,2)');
    assert.deepEqual([...worksheet.ranges.get('D8:E9').cells()].map(cell => cell.address), ['D8', 'E8', 'D9', 'E9']);
    for (const forbidden of ['session', 'runtime', 'model', 'dispatch', 'api', 'token', 'getAccessToken']) assert.equal(forbidden in workbook, false);
    assert.equal(Object.isFrozen(await cell.read()), true);
  } finally { workbook.close(); }
});
test('invalid object addresses and values reject without changes', async () => {
  const { session, workbook } = fixture('object-invalid');
  try {
    const sheet = workbook.worksheets.at(0), before = session['runtime'].model.snapshot();
    for (const address of ['A0', 'A1048577', 'XFE1', 'A$$1', 'Sheet1!A1', 'A1:B2']) assert.throws(() => sheet.cells.get(address), invalid);
    assert.throws(() => sheet.ranges.get('B2:A1'), invalid);
    assert.throws(() => workbook.worksheets.at(-1), invalid);
    await assert.rejects(sheet.cells.get('D8').setValue(Infinity), invalid);
    await assert.rejects(sheet.cells.get('D8').setFormula('SUM(1,2)'), invalid);
    assert.deepEqual(session['runtime'].model.snapshot(), before);
  } finally { workbook.close(); }
});
test('object permission checks reject viewer writes and hidden-range reads without mutating canonical cells', async () => {
  const { session, workbook, sheet } = fixture('object-permission');
  try {
    const cell = workbook.worksheets.at(0).cells.get('D8');
    const before = session['runtime'].model.snapshot();
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
    await assert.rejects(cell.setValue(4), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 2,
      regions: [{ range: { sheetId: sheet.id, startRow: 7, endRow: 7, startColumn: 3, endColumn: 3 }, access: 'hidden' }] });
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
  } finally { workbook.close(); }
});
test('closing one workbook retires its cell and range handles while another stays usable', async () => {
  const scope = {}, first = fixture('first', scope), second = fixture('second', scope);
  try {
    const cell = first.workbook.worksheets.at(0).cells.get('D8');
    const range = first.workbook.worksheets.at(0).ranges.get('A1:B2');
    first.workbook.close(); first.workbook.close();
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    assert.throws(() => [...range.cells()], (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    await second.workbook.worksheets.at(0).cells.get('D8').setValue(42);
    assert.equal((await second.workbook.worksheets.at(0).cells.get('D8').read()).value, 42);
  } finally { second.workbook.close(); }
});

test('canonical Session disposal and workbook lifecycle retire public objects without affecting another workbook', async () => {
  const first = fixture('session-retirement'), other = fixture('session-survivor');
  try {
    const cell = first.workbook.worksheets.at(0).cells.get('A1'), range = first.workbook.worksheets.at(0).ranges.get('A1:B2');
    first.session['runtime'].handlers.onWorkbookLifecycle?.('active');
    await cell.setValue(40);
    first.session['runtime'].handlers.onWorkbookLifecycle?.('trashed');
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    assert.throws(() => [...range.cells()], (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    await other.workbook.worksheets.at(0).cells.get('A1').setValue(7);
    assert.equal((await other.workbook.worksheets.at(0).cells.get('A1').read()).value, 7);
    other.session.dispose();
    assert.throws(() => other.workbook.name, (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
  } finally { first.workbook.close(); other.workbook.close(); }
});

test('source lifecycle retirement precedes initial synchronization and rejects late snapshot publication', async t => {
  const prior = Object.getOwnPropertyDescriptor(globalThis, 'window');
  Object.defineProperty(globalThis, 'window', { configurable: true, value: { location: { protocol: 'http:', host: 'localhost' }, setInterval, clearInterval } });
  t.after(() => { if (prior) Object.defineProperty(globalThis, 'window', prior); else Reflect.deleteProperty(globalThis, 'window'); });
  let status!: (status: 'connecting' | 'open' | 'closed') => void, message!: (message: OperationMessage) => void;
  t.mock.method(CollabSocketClient.prototype, 'onStatus', (listener: typeof status) => { status = listener; return () => {}; });
  t.mock.method(CollabSocketClient.prototype, 'onMessage', (listener: typeof message) => { message = listener; return () => {}; });
  t.mock.method(CollabSocketClient.prototype, 'open', () => status('open'));
  t.mock.method(CollabSocketClient.prototype, 'send', () => true);
  t.mock.method(CollabSocketClient.prototype, 'close', () => {});
  const { session, workbook } = fixture('life-during-sync');
  const cell = workbook.worksheets.at(0).cells.get('A1'), runtime = session['runtime'], snapshot = runtime.model.snapshot();
  runtime.localOnly = false;
  let release!: (snapshot: import('@react-sheets/protocol').SnapshotResponse) => void;
  let snapshotStarted!: () => void;
  const started = new Promise<void>(resolve => { snapshotStarted = resolve; });
  runtime.api.getSnapshot = async () => new Promise(resolve => { release = resolve; snapshotStarted(); });
  runtime.api.getAccess = async () => ({ unitId: workbook.id, role: 'owner', accessRevision: 0, regions: [] });
  runtime.api.listRevisions = async () => [];
  const detach = startCollaborationSession(runtime, () => 'A1');
  try {
    await started;
    message({ type: 'workbook.lifecycle.changed', unitId: 'another-book', lifecycle: 'purged' });
    assert.equal(runtime.disposed, false);
    message({ type: 'workbook.lifecycle.changed', unitId: workbook.id, lifecycle: 'purged' });
    assert.equal(runtime.disposed, true);
    await assert.rejects(cell.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    const oldModel = runtime.model;
    release({ unitId: workbook.id, snapshot, revision: 999 });
    await Promise.resolve(); await Promise.resolve();
    assert.equal(runtime.model, oldModel); assert.equal(runtime.remoteConnected, false); assert.notEqual(runtime.remoteRevision, 999);
  } finally { detach(); workbook.close(); }
});
test('cross-workbook formulas refresh authorized revisions and clear revoked inputs', async () => {
  const scope = {}, source = fixture('object-source', scope), target = fixture('object-target', scope);
  try {
    const runtime = target.session['runtime'];
    let revision = 1, denied = false;
    runtime.api.validateExternalLinkBinding = async () => {};
    runtime.api.getExternalCalculationGraph = async () => ({ schema: 'ExternalCalculationGraph', rootUnitId: target.workbook.id, subject: 'object-owner', nodes: [
      { unitId: target.workbook.id, state: 'connected', snapshot: runtime.model.snapshot(), revision: 0, accessRevision: 0, blockedRanges: [] },
      denied ? { unitId: source.workbook.id, state: 'denied', error: { code: 'FORBIDDEN', message: 'Source access revoked' } }
      : { unitId: source.workbook.id, state: 'connected', snapshot: source.session['runtime'].model.snapshot(), revision, accessRevision: 0, blockedRanges: [] },
    ] });
    const sourceSheet = source.workbook.worksheets.at(0);
    await sourceSheet.cells.get('D8').setValue(10);
    await sourceSheet.cells.get('D9').setValue(20);
    await target.workbook.externalLinks.bind(source.workbook, 'Source.xlsx');
    const cell = target.workbook.worksheets.at(0).cells.get('D8');
    await cell.setFormula(`=SUM('[Source.xlsx]${sourceSheet.name}'!D8:D9)`);
    await target.workbook.externalLinks.refresh();
    assert.equal((await cell.read()).calculatedValue, 30);
    await sourceSheet.cells.get('D9').setValue(40); revision++;
    assert.equal((await target.workbook.externalLinks.refresh())[0]?.sourceRevision, 2);
    assert.equal((await cell.read()).calculatedValue, 50);
    denied = true;
    assert.equal((await target.workbook.externalLinks.refresh())[0]?.state, 'denied');
    assert.equal(runtime.formula.getExternalCalculationLinks()[0]?.cells.length, 0);
    assert.equal(((await cell.read()).calculatedValue as { code: string }).code, '#BLOCKED!');
    assert.equal((await cell.read()).formula, `=SUM('[Source.xlsx]${sourceSheet.name}'!D8:D9)`);
  } finally { source.workbook.close(); target.workbook.close(); }
});
test('cross-scope, self and invalid-token binding refuse before touching target definitions', async () => {
  const first = fixture('bind-first'), second = fixture('bind-second');
  try {
    const before = first.session['runtime'].model.snapshot();
    for (const [source, token] of [[second.workbook, 'Source.xlsx'], [first.workbook, 'Self.xlsx'], [first.workbook, '../bad']] as const) {
      await assert.rejects(first.workbook.externalLinks.bind(source, token), invalid);
      assert.deepEqual(first.session['runtime'].model.snapshot(), before);
    }
  } finally { first.workbook.close(); second.workbook.close(); }
});

test('one explicit matrix preserves presentation and owns one undo/redo entry independent of selection', async () => {
  const { session, workbook, sheet } = fixture('matrix-history');
  try {
    const range = workbook.worksheets.at(0).ranges.get('D8:E9');
    sheet.cells.set(7, 3, { value: 99, style: { bold: true, borders: { bottom: { style: 'thin', color: '#112233' } } }, numberFormat: '0.00' });
    session.selectAddress('A1');
    const history = session['runtime'].commands.getUndoEntries().length;
    await range.setValues([[1, '=literal'], [true, null]]);
    assert.deepEqual(await range.readValues(), [[1, '=literal'], [true, null]]);
    assert.equal(session['runtime'].commands.getUndoEntries().length, history + 1);
    assert.equal((await range.read())[0]![0]!.style?.bold, true);
    assert.equal((await range.read())[0]![0]!.numberFormat, '0.00');
    assert.equal(await workbook.undo(), true);
    assert.deepEqual(await range.readValues(), [[99, null], [null, null]]);
    assert.equal(await workbook.redo(), true);
    assert.deepEqual(await range.readValues(), [[1, '=literal'], [true, null]]);
    assert.equal(sheet.cells.get(0, 0), undefined);
    const read = await range.read();
    assert.equal(Object.isFrozen(read), true); assert.equal(Object.isFrozen(read[0]), true);
    assert.equal(Object.isFrozen(read[0]![0]!.style?.borders), true);
    assert.throws(() => { (read[0]![0]!.style as { bold: boolean }).bold = false; }, TypeError);
    assert.equal(sheet.cells.get(7, 3)?.style?.bold, true);
  } finally { workbook.close(); }
});

test('typed values and formula matrices share the canonical calculation owner', async () => {
  const { workbook } = fixture('matrix-formulas');
  try {
    const sheet = workbook.worksheets.at(0), range = sheet.ranges.get('B3:C4');
    await range.setInputs([[{ kind: 'value', value: 10 }, { kind: 'formula', formula: '=B3*2' }], [{ kind: 'formula', formula: '=SUM(B3:C3)' }, { kind: 'value', value: '=literal' }]]);
    assert.deepEqual(await range.readValues(), [[10, 20], [30, '=literal']]);
    await sheet.ranges.get('E1:E2').setFormulas([['=SUM(B3:C3)'], ['=B4*2']]);
    assert.deepEqual(await sheet.ranges.get('E1:E2').readValues(), [[30], [60]]);
  } finally { workbook.close(); }
});

test('matrix shape, scalar, extent, budget and kernel target order reject before model or history changes', async () => {
  const { session, workbook, sheet } = fixture('matrix-invalid');
  try {
    const range = workbook.worksheets.at(0).ranges.get('A1:B2');
    const before = session['runtime'].model.snapshot(), history = session['runtime'].commands.getUndoEntries().length;
    await assert.rejects(range.setValues([[1, 2]]), invalid);
    await assert.rejects(range.setValues([[1, 2], [3, Infinity]]), invalid);
    await assert.rejects(range.setFormulas([['=1', '=2'], ['=3', '4']]), invalid);
    await assert.rejects(workbook.worksheets.at(0).ranges.get('A1:XFD2').read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE');
    await assert.rejects(workbook.worksheets.at(0).ranges.get(`A${sheet.rowCount + 1}`).setValues([[1]]), invalid);
    assert.throws(() => session['runtime'].commands.execute('sheet.cells.commitMatrix', { sheetId: sheet.id,
      range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      entries: [{ row: 0, column: 0, input: { kind: 'value', value: 1 } }, { row: 0, column: 0, input: { kind: 'value', value: 2 } }] }), /row-major/);
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    assert.equal(session['runtime'].commands.getUndoEntries().length, history);
  } finally { workbook.close(); }
});

test('matrix DV, hidden source and viewer/history permissions reject the whole operation', async () => {
  const { session, workbook, sheet } = fixture('matrix-permissions');
  try {
    session['runtime'].commands.execute('sheet.dv.add', { sheetId: sheet.id, rule: { id: 'matrix-minimum', sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }], type: 'whole', operator: 'greaterThan', formula1: '10', alertStyle: 'stop' } });
    const range = workbook.worksheets.at(0).ranges.get('A1:B1');
    const before = session['runtime'].model.snapshot(), history = session['runtime'].commands.getUndoEntries().length;
    await assert.rejects(range.setValues([[2, 5]]), (cause: unknown) => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    assert.equal(session['runtime'].commands.getUndoEntries().length, history);
    await range.setValues([[2, 11]]);
    const written = session['runtime'].model.snapshot();
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
    await assert.rejects(range.setValues([[3, 12]]), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    await assert.rejects(workbook.undo(), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), written);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'owner', accessRevision: 2,
      regions: [{ range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }, access: 'hidden' }] });
    await assert.rejects(range.read(), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    await assert.rejects(range.setValues([[3, 12]]), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), written);
  } finally { workbook.close(); }
});

test('range clear families preserve content or presentation and border operations stay canonical', async () => {
  const { workbook } = fixture('range-presentation');
  try {
    const range = workbook.worksheets.at(0).ranges.get('D8:E9');
    await range.setValues([[1, 2], [3, 4]]); await range.setStyle({ bold: true }, { numberFormat: '0.00' });
    await range.setBorders('outside', { style: 'thin', color: '#123456' });
    assert.equal((await range.read())[0]![0]!.style?.borders?.top?.color, '#123456');
    await range.clear('contents');
    assert.deepEqual(await range.readValues(), [[null, null], [null, null]]);
    assert.equal((await range.read())[0]![0]!.style?.bold, true);
    await workbook.undo(); assert.deepEqual(await range.readValues(), [[1, 2], [3, 4]]);
    await range.clear('formats');
    assert.deepEqual(await range.readValues(), [[1, 2], [3, 4]]);
    assert.equal((await range.read())[0]![0]!.style, undefined);
    assert.equal((await range.read())[0]![0]!.numberFormat, undefined);
  } finally { workbook.close(); }
});

test('clear undo and redo synchronize dependent calculations and copied scalar values', async () => {
  for (const kind of ['contents', 'all'] as const) {
    const { workbook, session } = fixture(`clear-calculation-${kind}`);
    try {
      const sheet = workbook.worksheets.at(0), range = sheet.ranges.get('D8:E9'), formulas = sheet.ranges.get('G1:G2');
      await range.setValues([[1, '=literal'], [true, null]]);
      await formulas.setFormulas([['=SUM(D8:E8)'], ['=COUNT(D8:E9)']]);
      assert.deepEqual(await formulas.readValues(), [[1], [1]]);
      await range.clear(kind);
      assert.deepEqual(await formulas.readValues(), [[0], [0]]);
      assert.equal(await workbook.undo(), true);
      assert.deepEqual(await range.readValues(), [[1, '=literal'], [true, null]]);
      assert.deepEqual(await formulas.readValues(), [[1], [1]]);
      assert.equal(await workbook.redo(), true);
      assert.deepEqual(await formulas.readValues(), [[0], [0]]);
      await workbook.undo();
      await formulas.copyValuesTo(sheet.ranges.get('H1:H2'));
      assert.deepEqual(await sheet.ranges.get('H1:H2').readValues(), [[1], [1]]);
      session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
      const before = session['runtime'].model.snapshot();
      await assert.rejects(range.clear(kind), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
      await assert.rejects(workbook.undo(), (cause: unknown) => cause instanceof SdkError && cause.code === 'FORBIDDEN');
      assert.deepEqual(session['runtime'].model.snapshot(), before);
      assert.deepEqual(await formulas.readValues(), [[1], [1]]);
    } finally { workbook.close(); }
  }
});

test('manual clear and inverse publish the canonical input journal to the Worker executor', async () => {
  const { workbook, session } = fixture('clear-worker-journal');
  let workerEngine: FormulaEngine | null = null, calculationPosts = 0;
  const listeners = new Set<(event: { data?: unknown }) => void>();
  const worker: CalculationBrowserWorker = {
    postMessage(message) {
      if ((message as { kind?: string }).kind !== 'recalculate') return;
      calculationPosts += 1;
      const consumed = consumeBrowserCalculationTaskWithEngine(message, workerEngine);
      workerEngine = consumed.engine;
      queueMicrotask(() => { for (const listener of listeners) listener({ data: consumed.result }); });
    },
    terminate() { listeners.clear(); },
    addEventListener(type, listener) { if (type === 'message') listeners.add(listener); },
    removeEventListener(type, listener) { if (type === 'message') listeners.delete(listener); },
  };
  let port: ReturnType<FormulaEngine['createCalculationTaskPort']> | undefined;
  try {
    const sheet = workbook.worksheets.at(0), range = sheet.ranges.get('D8:E9'), formulas = sheet.ranges.get('G1:G2');
    await range.setValues([[1, '=literal'], [true, null]]);
    await formulas.setFormulas([['=SUM(D8:E8)'], ['=COUNT(D8:E9)']]);
    assert.deepEqual(await formulas.readValues(), [[1], [1]]);
    session.setRecalculationMode('manual');
    const engine = session['runtime'].formula;
    port = engine.createCalculationTaskPort({ workerFactory: () => worker });
    await range.clear('contents');
    assert.deepEqual(await formulas.readValues(), [[1], [1]]);
    await engine.recalculateAsync(undefined, port);
    assert.deepEqual(await formulas.readValues(), [[0], [0]]);
    await workbook.undo();
    assert.deepEqual(await formulas.readValues(), [[0], [0]]);
    await engine.recalculateAsync(undefined, port);
    assert.deepEqual(await formulas.readValues(), [[1], [1]]);
    await workbook.redo();
    assert.deepEqual(await formulas.readValues(), [[1], [1]]);
    await engine.recalculateAsync(undefined, port);
    assert.deepEqual(await formulas.readValues(), [[0], [0]]);
    assert.equal(calculationPosts, 3);
  } finally { port?.dispose?.(); workbook.close(); }
});

test('worksheet dimensions and visibility share canonical ownership and offline structural changes reject', async () => {
  const { workbook, session } = fixture('worksheet-objects');
  try {
    const sheet = workbook.worksheets.at(0);
    await sheet.rows.setPixels([2], 30); await sheet.columns.setPixels([3], 90);
    await sheet.rows.setHidden([2], true); await sheet.columns.setHidden([3], true);
    await sheet.cells.get('D3').setValue(42); assert.equal((await sheet.cells.get('D3').read()).value, 42);
    assert.equal(sheet.snapshot().rowHeightsPx[2], 30); assert.equal(sheet.snapshot().columnWidthsPx[3], 90);
    assert.deepEqual(sheet.snapshot().hiddenRows, [2]); assert.deepEqual(sheet.snapshot().hiddenColumns, [3]);
    await sheet.setPane({ kind: 'frozen', xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1, state: 'frozen' });
    assert.equal(sheet.snapshot().pane.kind, 'frozen');
    const before = session['runtime'].model.snapshot();
    for (const action of [() => sheet.rename('Changed'), () => sheet.rows.insert(0), () => sheet.duplicate('Copy'), () => workbook.worksheets.add({ name: 'Second' })]) {
      await assert.rejects(action(), (cause: unknown) => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED' && /STRUCTURAL_PLANNER_OFFLINE/.test(cause.message));
    }
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    await assert.rejects(sheet.columns.setHidden([-1], true), invalid);
  } finally { workbook.close(); }
});

test('fill requires the real structural planner and explicit merge confirmation retains undo semantics', async () => {
  const { workbook } = fixture('range-fill-merge');
  try {
    const sheet = workbook.worksheets.at(0), seed = sheet.ranges.get('A1:A2'), target = sheet.ranges.get('A1:A4');
    await seed.setValues([[1], [2]]);
    await assert.rejects(target.fillFrom(seed, 'down', 'series'), (cause: unknown) => cause instanceof SdkError && /STRUCTURAL_PLANNER_OFFLINE/.test(cause.message));
    assert.deepEqual(await target.readValues(), [[1], [2], [null], [null]]);
    const merge = sheet.ranges.get('D8:E8'); await merge.setValues([[10, 20]]);
    await assert.rejects(merge.merge(), (cause: unknown) => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED');
    assert.deepEqual(await merge.readValues(), [[10, 20]]);
    await merge.merge({ confirmDataLoss: true }); assert.deepEqual(await merge.readValues(), [[10, null]]);
    assert.equal(sheet.snapshot().merges.length, 1);
    await workbook.undo(); assert.deepEqual(await merge.readValues(), [[10, 20]]); assert.equal(sheet.snapshot().merges.length, 0);
  } finally { workbook.close(); }
});

test('two-workbook value copy preserves source and rejects errors or another SDK scope', async () => {
  const scope = {}, source = fixture('values-source', scope), target = fixture('values-target', scope), other = fixture('values-other');
  try {
    const from = source.workbook.worksheets.at(0).ranges.get('A1:B1'), to = target.workbook.worksheets.at(0).ranges.get('D8:E8');
    await from.setInputs([[{ kind: 'formula', formula: '=2+3' }, { kind: 'value', value: 10 }]]);
    await from.copyValuesTo(to); assert.deepEqual(await to.readValues(), [[5, 10]]); assert.deepEqual(await from.readValues(), [[5, 10]]);
    assert.equal((await to.read())[0]![0]!.formula, undefined);
    await assert.rejects(from.copyValuesTo(other.workbook.worksheets.at(0).ranges.get('A1:B1')), invalid);
    await source.workbook.worksheets.at(0).cells.get('A1').setFormula('=1/0');
    const before = target.session['runtime'].model.snapshot();
    await assert.rejects(from.copyValuesTo(to), (cause: unknown) => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE');
    assert.deepEqual(target.session['runtime'].model.snapshot(), before);
  } finally { source.workbook.close(); target.workbook.close(); other.workbook.close(); }
});

test('matrix intent is captured before async authorization and calculation waits', async () => {
  const { workbook } = fixture('matrix-intent');
  try {
    const range = workbook.worksheets.at(0).ranges.get('A1');
    const inputs: { kind: 'value'; value: number }[][] = [[{ kind: 'value', value: 1 }]];
    const pending = range.setInputs(inputs);
    inputs[0]![0]!.value = 999;
    await pending;
    assert.deepEqual(await range.readValues(), [[1]]);
  } finally { workbook.close(); }
});

test('canonical worksheet rename publishes cell, rule and name facts for committed undo/redo', () => {
  const model = new WorkbookModel('rename-facts', 'Rename facts'), owner = model.getSheet(model.primarySheetId), source = model.addSheet('source', 'Source');
  owner.cells.set(0, 0, { value: null, formula: '=Source!A1' });
  model.setDefinedName({ name: 'Rate', formula: '=Source!A1', scope: 'workbook' });
  const runtime = new CommandRuntime(model); registerSheetCommands(runtime);
  const range = { sheetId: owner.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  runtime.execute('sheet.cf.add', { sheetId: owner.id, rule: { id: 'cf', sheetId: owner.id, ranges: [range], type: 'highlight', operator: 'formula', value1: '=Source!A1>0' } });
  runtime.execute('sheet.dv.add', { sheetId: owner.id, rule: { id: 'dv', sheetId: owner.id, ranges: [range], type: 'list', listSource: { kind: 'formula', formula: '=Source!A1:A2' } } });
  const result = runtime.execute('sheet.rename', { sheetId: source.id, name: 'Renamed' });
  const entry = runtime.getUndoEntries().at(-1)!, forward = entry.forwardMutations[0]!;
  assert.equal(forward.structuralFormulaOwnerDeltas?.length, 3);
  assert.equal(forward.structuralDefinedNameOwnerDeltas?.length, 1);
  runtime.applyCommittedStructuralPatches(result.operationId, [{ ...structuredClone(forward), structuralRangeOwnerDeltas: [] }], 1);
  assert.equal(runtime.undo(), true);
  assert.equal(source.name, 'Source'); assert.equal(owner.cells.get(0, 0)?.formula, '=Source!A1');
  assert.equal(model.getDefinedNameExact('Rate', 'workbook')?.formula, '=Source!A1');
  assert.equal(runtime.redo(), true);
  assert.equal(source.name, 'Renamed'); assert.match(owner.cells.get(0, 0)!.formula!, /Renamed/);
});

test('committed worksheet rename rejects forged owner semantics without repairing model or history', () => {
  const model = new WorkbookModel('rename-rejection', 'Rename rejection'), owner = model.getSheet(model.primarySheetId), source = model.addSheet('source', 'Source');
  owner.cells.set(0, 0, { value: null, formula: '=Source!A1' });
  const runtime = new CommandRuntime(model); registerSheetCommands(runtime);
  const result = runtime.execute('sheet.rename', { sheetId: source.id, name: 'Renamed' });
  const before = model.snapshot(), history = structuredClone(runtime.getUndoEntries());
  const forged = structuredClone(history.at(-1)!.forwardMutations[0]!);
  forged.structuralDefinedNameOwnerDeltas = []; forged.structuralRangeOwnerDeltas = [];
  const delta = forged.structuralFormulaOwnerDeltas![0]!;
  if (delta.kind !== 'formula-cell') throw new Error('Expected canonical cell owner');
  forged.structuralFormulaOwnerDeltas = [{ ...delta, after: { ...delta.after, formula: '=Other!B99' } }];
  assert.throws(() => runtime.applyCommittedStructuralPatches(result.operationId, [forged], 1), /STRUCTURAL_PATCH_MISMATCH/);
  assert.deepEqual(model.snapshot(), before); assert.deepEqual(runtime.getUndoEntries(), history);
});


test('object defined names preserve canonical scope, anchors, immutable identity and calculation history', async () => {
  const { workbook, session } = fixture('named-objects');
  try {
    const sheet = workbook.worksheets.at(0);
    await sheet.ranges.get('A1:B1').setValues([[2, 3]]);
    await workbook.names.define({ name: 'Rate', formula: '=2', scope: 'workbook', hidden: true, comment: 'Global' });
    const local = await workbook.names.define({ name: 'Rate', formula: '=3', scope: 'sheet', sheetId: sheet.id });
    const relative = await workbook.names.define({ name: 'Relative', formula: '=A1', scope: 'workbook', anchor: { sheetId: sheet.id, row: 3, column: 3 }, comment: 'Relative owner' });
    assert.equal(workbook.names.byName('rate', 'sheet', sheet.id), local);
    assert.equal(Object.isFrozen(relative.snapshot().anchor), true);
    await sheet.cells.get('D1').setFormula('=Rate*10');
    await sheet.cells.get('D4').setFormula('=Relative');
    await sheet.cells.get('E4').setFormula('=Relative');
    assert.equal((await sheet.cells.get('D1').read()).calculatedValue, 30);
    assert.equal((await sheet.cells.get('D4').read()).calculatedValue, 2);
    assert.equal((await sheet.cells.get('E4').read()).calculatedValue, 3);
    await local.setFormula('=4'); assert.equal((await sheet.cells.get('D1').read()).calculatedValue, 40);
    const namePreimage = session['runtime'].model.snapshot();
    await local.remove(); assert.equal((await sheet.cells.get('D1').read()).calculatedValue, 20);
    assert.throws(() => local.snapshot(), invalid);
    await workbook.undo(); assert.equal(workbook.names.byName('Rate', 'sheet', sheet.id), local); assert.equal((await sheet.cells.get('D1').read()).calculatedValue, 40);
    assert.deepEqual(session['runtime'].model.snapshot(), namePreimage);
    await workbook.redo(); assert.equal((await sheet.cells.get('D1').read()).calculatedValue, 20);
    const before = session['runtime'].model.snapshot(), depth = session['runtime'].commands.getHistoryDepth();
    for (const model of [{ name: 'NoScope', formula: '=1' }, { name: 'Legacy', value: '=1', scope: 'workbook' },
      { name: 'Bad', formula: '=1', scope: 'sheet', sheetId: 'missing' }, { name: 'Bad', formula: '=A1', scope: 'workbook', anchor: { sheetId: 'missing', row: 0, column: 0 } }]) {
      await assert.rejects(workbook.names.define(model as never), cause => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED');
    }
    assert.deepEqual(session['runtime'].model.snapshot(), before); assert.deepEqual(session['runtime'].commands.getHistoryDepth(), depth);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
    await assert.rejects(relative.setFormula('=2'), cause => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    workbook.close(); assert.throws(() => relative.snapshot(), cause => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
  } finally { workbook.close(); }
});

test('rich-text object commits preserve literal text and style in one canonical history transaction', async () => {
  const { workbook, session } = fixture('richtext-objects');
  try {
    const sheet = workbook.worksheets.at(0), range = sheet.ranges.get('D8:E9');
    await range.setValues([[1, 2], [3, 4]]); await range.setStyle({ bold: true }, { numberFormat: '0.00' });
    const depth = session['runtime'].commands.getHistoryDepth().undo;
    const runs = [{ text: '=lit', style: { bold: true, fontFamily: ' arial ' } }, { text: 'eral', style: { italic: true, textColor: '#123456' } }];
    await range.setRichText('=literal', runs);
    assert.equal(session['runtime'].commands.getHistoryDepth().undo, depth + 1);
    assert.deepEqual(await range.readValues(), [['=literal', '=literal'], ['=literal', '=literal']]);
    for (const row of await range.read()) for (const cell of row) { assert.equal(cell.formula, undefined); assert.equal(cell.numberFormat, '0.00'); assert.equal(cell.richText?.[0]?.style?.fontFamily, 'Arial'); assert.equal(Object.isFrozen(cell.richText?.[1]?.style), true); }
    await workbook.undo(); assert.deepEqual(await range.readValues(), [[1, 2], [3, 4]]);
    await workbook.redo(); assert.equal((await range.read())[1]![1]!.richText?.[1]?.style?.italic, true);
    await sheet.cells.get('D8').setStyle({ italic: true }); await sheet.cells.get('D8').setNumberFormat('0');
    assert.equal((await sheet.cells.get('D8').read()).style?.italic, true);
    const before = session['runtime'].model.snapshot();
    await assert.rejects(range.setRichText('mismatch', runs), invalid);
    for (const style of [{ fontSizePx: NaN }, { bold: 'yes' }, { fontFamily: ' ' }, { arbitrary: true }]) await assert.rejects(range.setRichText('x', [{ text: 'x', style } as never]), cause => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED');
    await assert.rejects(sheet.ranges.get('A1:XFD1048576').setRichText('x', [{ text: 'x' }]), cause => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE' && Boolean(cause.object?.workbookId) && Boolean(cause.recovery));
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    session.runCommand('sheet.dv.add', { sheetId: sheet.id, rule: { id: 'richtext-dv', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 7, endRow: 7, startColumn: 4, endColumn: 4 }], type: 'whole', operator: 'greaterThan', formula1: '10' } });
    const withRule = session['runtime'].model.snapshot();
    await assert.rejects(range.setRichText('invalid', [{ text: 'invalid' }]), cause => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED');
    assert.deepEqual(session['runtime'].model.snapshot(), withRule);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'viewer', accessRevision: 1, regions: [] });
    await assert.rejects(sheet.cells.get('A1').setRichText('x', [{ text: 'x' }]), cause => cause instanceof SdkError && cause.code === 'FORBIDDEN');
  } finally { workbook.close(); }
});

test('rich-text captures caller intent before async source authorization', async () => {
  const { workbook, session } = fixture('richtext-intent');
  try {
    const port = getWorkbookObjectPort(session), read = port.readCells.bind(port);
    let resume!: () => void; const blocked = new Promise<void>(resolve => { resume = resolve; });
    port.readCells = async range => { await blocked; return read(range); };
    const runs = [{ text: 'captured', style: { bold: true } }];
    const commit = workbook.worksheets.at(0).cells.get('D8').setRichText('captured', runs);
    runs[0]!.text = 'changed'; runs[0]!.style.bold = false; resume(); await commit;
    assert.equal((await workbook.worksheets.at(0).cells.get('D8').read()).value, 'captured');
    assert.equal((await workbook.worksheets.at(0).cells.get('D8').read()).richText?.[0]?.style?.bold, true);
  } finally { workbook.close(); }
});

test('worksheet protection objects enforce canonical ownership, owner ACL and reversible allow flags', async () => {
  const { workbook, session } = fixture('protection-objects');
  try {
    const sheet = workbook.worksheets.at(0), protection = sheet.protection;
    await sheet.cells.get('D8').setValue(1); await sheet.cells.get('E8').setStyle({ locked: false });
    await protection.set({ id: 'lock', scope: 'sheet', sheetId: sheet.id, locked: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: true } });
    assert.equal(Object.isFrozen(protection.list()[0]?.allow), true);
    await assert.rejects(sheet.cells.get('D8').setValue(2), cause => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    await sheet.cells.get('E8').setValue(3);
    await protection.remove('lock'); await workbook.undo(); assert.equal(protection.list().length, 1);
    await workbook.redo(); assert.deepEqual(protection.list(), []);
    const before = session['runtime'].model.snapshot();
    for (const rule of [{ id: 'bad', scope: 'sheet', locked: true, allow: { sort: 'yes' } }, { id: 'bad', scope: 'sheet', locked: true, allow: {}, arbitrary: true },
      { id: 'bad', scope: 'range', locked: true, allow: {}, range: { sheetId: 'other', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }]) await assert.rejects(protection.set(rule as never), cause => cause instanceof SdkError);
    await assert.rejects(protection.set({ id: 'unsupported', scope: 'workbook', locked: true, allow: {} }), cause => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE');
    await assert.rejects(protection.remove('missing'), invalid);
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    session['permission'].applyServerAccess({ unitId: workbook.id, role: 'editor', accessRevision: 1, regions: [] });
    await assert.rejects(protection.set({ id: 'editor-lock', scope: 'sheet', locked: true, allow: {} }), cause => cause instanceof SdkError && cause.code === 'FORBIDDEN');
    assert.deepEqual(session['runtime'].model.snapshot(), before);
    workbook.close(); assert.throws(() => protection.list(), cause => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
  } finally { workbook.close(); }
});


test('canonical name restore rejects invalid positions and duplicate owners without changing snapshots', () => {
  const workbook = new WorkbookModel('ordered-names', 'Ordered names');
  for (const name of ['First', 'Last']) workbook.setDefinedName({ name, scope: 'workbook', formula: '=1' });
  const before = workbook.snapshot(), middle = { name: 'Middle', scope: 'workbook' as const, formula: '=2' };
  for (const position of [-1, 3, 0.5, NaN]) { assert.throws(() => workbook.restoreDefinedName(middle, position), /position is invalid/); assert.deepEqual(workbook.snapshot(), before); }
  assert.throws(() => workbook.restoreDefinedName({ ...middle, name: 'first' }, 1), /identity already exists/);
  assert.deepEqual(workbook.snapshot(), before);
  workbook.restoreDefinedName(middle, 1); assert.deepEqual(workbook.definedNameModels.map(name => name.name), ['First', 'Middle', 'Last']);
});


test('Range.moveTo uses the canonical planner and rejects unequal, cross-workbook, oversized and retired destinations', async () => {
  const first = fixture('sdk-cut'), second = fixture('sdk-cut-other');
  try {
    const sheet = first.workbook.worksheets.at(0), source = sheet.ranges.get('A1:B1');
    await source.setValues([[7, 11]]); const before = first.session['runtime'].model.snapshot();
    await assert.rejects(source.moveTo(sheet.ranges.get('C3')), invalid);
    await assert.rejects(source.moveTo(second.workbook.worksheets.at(0).ranges.get('C3:D3')), (cause: unknown) => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE');
    await assert.rejects(sheet.ranges.get('A1:A10001').moveTo(sheet.ranges.get('B1:B10001')), (cause: unknown) => cause instanceof SdkError && cause.code === 'UNSUPPORTED_FEATURE');
    await assert.rejects(source.moveTo(sheet.ranges.get('C3:D3')), (cause: unknown) => cause instanceof SdkError && cause.code === 'REQUEST_REJECTED' && /STRUCTURAL_PLANNER_OFFLINE/.test(cause.message));
    assert.deepEqual(first.session['runtime'].model.snapshot(), before);
    const retired = second.workbook.worksheets.at(0).ranges.get('C3:D3'); second.workbook.close();
    await assert.rejects(source.moveTo(retired), (cause: unknown) => cause instanceof SdkError && cause.code === 'RUNTIME_DISPOSED');
    assert.deepEqual(first.session['runtime'].model.snapshot(), before);
  } finally { first.workbook.close(); second.workbook.close(); }
});
