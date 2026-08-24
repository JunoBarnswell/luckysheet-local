import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession drawing integration', () => {
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

  it('insertShape creates a drawable rectangle', () => {
    const app = new WorkbookSession();
    app.insertShape('rectangle');
    const snapshot = app.getUiSnapshot();
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const drawing = sheet.drawings.find((entry) => entry.kind === 'shape');
    assert.equal((sheet.drawingPayloads.get(drawing?.payloadId ?? '') as { type?: string }).type, 'rectangle');
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
