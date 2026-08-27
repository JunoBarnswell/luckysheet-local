import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { WorkbookSession } from './workbook-session';
import { hydrateRuntime } from './runtime';

describe('WorkbookSession drawing integration', () => {
  it('places a text box through the placement session and commits one text-frame mutation', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.insertTextBox();
    assert.equal(app.getUiSnapshot().textBoxPlacement, true);
    assert.equal(app['runtime'].model.getSheet(sheetId).drawings.filter((drawing) => drawing.kind === 'textbox').length, 0);
    app.placeTextBox({ x: 40, y: 50, width: 180, height: 60 });
    let snapshot = app.getUiSnapshot();
    const drawing = app['runtime'].model.getSheet(sheetId).drawings.find((entry) => entry.kind === 'textbox');
    assert.ok(drawing);
    assert.equal(snapshot.textBoxPlacement, false);
    assert.deepEqual(snapshot.textBoxEdit, { sheetId, drawingId: drawing.id, draftText: '' });
    app.setTextBoxDraft('Canonical text');
    app.commitTextBoxEdit();
    snapshot = app.getUiSnapshot();
    const payload = app['runtime'].model.getSheet(sheetId).drawingPayloads.get(drawing.payloadId);
    assert.equal(payload?.kind === 'textbox' ? payload.text : '', 'Canonical text');
    assert.equal(snapshot.textBoxEdit, null);
    assert.equal(snapshot.historyEntries.at(-1)?.redo[0]?.id, 'drawing.payload.update');
    app.undo();
    assert.equal((app['runtime'].model.getSheet(sheetId).drawingPayloads.get(drawing.payloadId) as { text?: string } | undefined)?.text, '');
    app.redo();
    assert.equal((app['runtime'].model.getSheet(sheetId).drawingPayloads.get(drawing.payloadId) as { text?: string } | undefined)?.text, 'Canonical text');
  });
  it('addShape routes through the canonical drawing aggregate', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addShape({
      id: 'draw-shape-test-1',
      sheetId,
      kind: 'shape',
      payloadId: 'shape-test-1',
      anchor: { kind: 'absolute' },
      transform: { x: 40, y: 40, width: 120, height: 48, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#dbeafe',
      stroke: '#2563eb',
      strokeWidth: 2,
      text: 'Box',
      textColor: '#1e3a8a',
      fontSize: 13,
    });

    const sheet = app['runtime'].model.getSheet(sheetId);
    const drawing = sheet.drawings.find((entry) => entry.payloadId === 'shape-test-1');
    assert.equal(drawing?.kind, 'shape');
    assert.equal((sheet.drawingPayloads.get('shape-test-1') as { text?: string }).text, 'Box');
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, 'draw-shape-test-1');

    app.updateShapeBounds('shape-test-1', { x: 80, y: 90, width: 120, height: 48 });
    snapshot = app.getUiSnapshot();
    assert.equal(sheet.drawings.find((entry) => entry.payloadId === 'shape-test-1')?.transform.x, 80);
    assert.equal(sheet.drawings.find((entry) => entry.payloadId === 'shape-test-1')?.transform.y, 90);

    app.bringSelectedDrawingForward();
    app.removeSelectedDrawing();
    snapshot = app.getUiSnapshot();
    assert.equal(sheet.drawings.some((entry) => entry.payloadId === 'shape-test-1'), false);
    assert.equal(snapshot.selectedFloatingId, null);
  });

  it('enters and leaves Shape Format context from canonical drawing selection', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addShape({ id: 'shape-context', sheetId, kind: 'shape', payloadId: 'shape-context-payload', anchor: { kind: 'absolute' }, transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 }, zIndex: 0 }, {
      kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#111827', strokeWidth: 1,
    });
    assert.deepEqual(app.getUiSnapshot().activeContext, { kind: 'drawing', sheetId, drawingId: 'shape-context' });
    assert.equal(app.getUiSnapshot().ribbon.activeTab, 'shapeFormat');
    assert.equal(app.getUiSnapshot().panels.active, 'shape');
    app.setDrawingSelection([]);
    assert.deepEqual(app.getUiSnapshot().activeContext, { kind: 'none' });
    assert.equal(app.getUiSnapshot().ribbon.activeTab, 'home');
  });

  it('insertShape creates a drawable rectangle', () => {
    const app = new WorkbookSession();
    app.insertShape('rectangle');
    const snapshot = app.getUiSnapshot();
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const drawing = sheet.drawings.find((entry) => entry.kind === 'shape');
    assert.equal((sheet.drawingPayloads.get(drawing?.payloadId ?? '') as { type?: string }).type, 'rectangle');
  });

  it('inserts a bound connector only for two selected shapes and exposes Shape Format context', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const add = (id: string, payloadId: string, x: number) => app.addShape({ id, sheetId, kind: 'shape', payloadId, anchor: { kind: 'absolute' }, transform: { x, y: 20, width: 80, height: 40, rotation: 0 }, zIndex: 0 }, {
      kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000',
    });
    add('connector-source', 'connector-source-payload', 20);
    add('connector-target', 'connector-target-payload', 180);
    app.setDrawingSelection(['connector-source', 'connector-target']);
    app.insertConnector('elbow');
    const sheet = app['runtime'].model.getSheet(sheetId);
    const connector = sheet.drawings.find((drawing) => drawing.kind === 'connector');
    assert.ok(connector);
    const payload = sheet.drawingPayloads.get(connector.payloadId);
    assert.equal(payload?.kind, 'connector');
    assert.equal(payload?.kind === 'connector' ? payload.connectorType : undefined, 'elbow');
    assert.equal(app.getUiSnapshot().ribbon.activeTab, 'shapeFormat');
    assert.throws(() => { app.setDrawingSelection(['connector-source']); app.insertConnector(); }, /exactly two/);
    app.undo();
    assert.equal(sheet.drawings.some((drawing) => drawing.kind === 'connector'), false);
  });

  it('clears drawing selection and context on undo, with deterministic grid focus on redo', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addShape({
      id: 'draw-undo-reconcile',
      sheetId,
      kind: 'shape',
      payloadId: 'shape-undo-reconcile',
      anchor: { kind: 'absolute' },
      transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#fff',
      stroke: '#000',
    });
    app.setDrawingSelectionMode(true);
    app.undo();
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, null);
    assert.deepEqual(snapshot.selectedDrawingIds, []);
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });
    assert.equal(snapshot.drawingSelectionMode, false);

    app.redo();
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, null);
    assert.deepEqual(snapshot.selectedDrawingIds, []);
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });
    assert.equal(snapshot.selectedSheet.drawings.some((drawing) => drawing.id === 'draw-undo-reconcile'), true);
  });

  it('keeps valid multi-selection while removing the deleted object from active context', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const add = (id: string, payloadId: string, x: number) => app.addShape({
      id,
      sheetId,
      kind: 'shape',
      payloadId,
      anchor: { kind: 'absolute' },
      transform: { x, y: 20, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#fff',
      stroke: '#000',
    });
    add('draw-multi-a', 'shape-multi-a', 20);
    add('draw-multi-b', 'shape-multi-b', 140);
    app.setDrawingSelection(['draw-multi-a', 'draw-multi-b']);
    app.runCommand('drawing.remove', { sheetId, drawingId: 'draw-multi-a' });

    const snapshot = app.getUiSnapshot();
    assert.deepEqual(snapshot.selectedDrawingIds, ['draw-multi-b']);
    assert.deepEqual(snapshot.activeContext, { kind: 'drawing', sheetId, drawingId: 'draw-multi-b' });
  });

  it('reconciles remote removal and deleted-sheet selection through the same boundary', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addShape({
      id: 'draw-remote-reconcile',
      sheetId,
      kind: 'shape',
      payloadId: 'shape-remote-reconcile',
      anchor: { kind: 'absolute' },
      transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#fff',
      stroke: '#000',
    });
    const runtime = app['runtime'];
    runtime.commands.applyRemoteMutations([{
      id: 'drawing.remove',
      unitId: runtime.model.unitId,
      sheetId,
      params: { sheetId, drawingId: 'draw-remote-reconcile' },
      affectedRanges: [],
    }]);
    runtime.handlers.onMutationsApplied?.();
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, null);
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });

    runtime.commands.execute('sheet.add', { id: 'sheet-with-drawing', name: 'Drawing Sheet' });
    app.selectSheet('sheet-with-drawing');
    app.addShape({
      id: 'draw-deleted-sheet',
      sheetId: 'sheet-with-drawing',
      kind: 'shape',
      payloadId: 'shape-deleted-sheet',
      anchor: { kind: 'absolute' },
      transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#fff',
      stroke: '#000',
    });
    app.deleteSheet('sheet-with-drawing');
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.some((sheet) => sheet.id === 'sheet-with-drawing'), false);
    assert.equal(snapshot.selectedFloatingId, null);
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });
  });

  it('clears stale drawing state when a snapshot replacement omits the selected object', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addShape({
      id: 'draw-hydration-reconcile',
      sheetId,
      kind: 'shape',
      payloadId: 'shape-hydration-reconcile',
      anchor: { kind: 'absolute' },
      transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#fff',
      stroke: '#000',
    });
    const runtime = app['runtime'];
    const replacement = new WorkbookModel(runtime.model.unitId, 'Replacement').snapshot();
    hydrateRuntime(runtime, { snapshot: replacement, revision: 7 });
    runtime.handlers.onMutationsApplied?.();

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, null);
    assert.deepEqual(snapshot.selectedDrawingIds, []);
    assert.deepEqual(snapshot.activeContext, { kind: 'none' });
    assert.equal(snapshot.selectedSheet.drawings.length, 0);
  });

  it('exposes Page Layout arrange actions through the existing drawing command path', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const add = (id: string, payloadId: string, x: number, y: number) => app.addShape({
      id,
      sheetId,
      kind: 'shape',
      payloadId,
      anchor: { kind: 'absolute' },
      transform: { x, y, width: 80, height: 40, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'shape',
      type: 'rectangle',
      fill: '#ffffff',
      stroke: '#111827',
      strokeWidth: 1,
      text: id,
    });
    add('arrange-a', 'arrange-a-payload', 20, 20);
    add('arrange-b', 'arrange-b-payload', 180, 80);
    add('arrange-c', 'arrange-c-payload', 340, 140);
    app.runCommand('drawing.select', { sheetId, drawingIds: ['arrange-a', 'arrange-b', 'arrange-c'] });
    app.alignSelectedDrawings('left');
    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.drawings.find((entry) => entry.id === 'arrange-a')?.transform.x, 20);
    assert.equal(sheet.drawings.find((entry) => entry.id === 'arrange-b')?.transform.x, 20);
    assert.equal(sheet.drawings.find((entry) => entry.id === 'arrange-c')?.transform.x, 20);
    app.distributeSelectedDrawings('vertical');
    app.bringSelectedDrawingToFront();
    app.sendSelectedDrawingToBack();
    assert.ok(sheet.drawings.every((entry) => Number.isFinite(entry.zIndex)));
  });
});
