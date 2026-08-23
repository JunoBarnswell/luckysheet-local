import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

describe('SpreadsheetApplication drawing integration', () => {
  it('addShape routes through the canonical drawing aggregate', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.addShape({
      id: 'shape-test-1',
      sheetId,
      type: 'rectangle',
      fill: '#dbeafe',
      stroke: '#2563eb',
      strokeWidth: 2,
      text: 'Box',
      textColor: '#1e3a8a',
      fontSize: 13,
      bounds: { x: 40, y: 40, width: 120, height: 48 },
    });

    const sheet = app.getWorkbook().getSheet(sheetId);
    const drawing = sheet.drawings.find((entry) => entry.payloadId === 'shape-test-1');
    assert.equal(drawing?.kind, 'shape');
    assert.equal((sheet.drawingPayloads.get('shape-test-1') as { text?: string }).text, 'Box');
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, 'shape-test-1');

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

  it('insertQuickShape creates a drawable rectangle', () => {
    const app = new SpreadsheetApplication();
    app.insertQuickShape('rectangle');
    const snapshot = app.getUiSnapshot();
    const sheet = app.getWorkbook().getSheet(app.getActiveSheetId());
    const drawing = sheet.drawings.find((entry) => entry.kind === 'shape');
    assert.equal((sheet.drawingPayloads.get(drawing?.payloadId ?? '') as { type?: string }).type, 'rectangle');
  });
});
