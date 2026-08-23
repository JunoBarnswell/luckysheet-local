import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

describe('SpreadsheetApplication drawing integration', () => {
  it('addShape routes through drawing.add and syncs legacy shapes', () => {
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

    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.shapes.length, 1);
    assert.equal(snapshot.selectedSheet.shapes[0]?.text, 'Box');
    assert.equal(snapshot.selectedFloatingId, 'shape-test-1');

    app.updateShapeBounds('shape-test-1', { x: 80, y: 90, width: 120, height: 48 });
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.shapes[0]?.bounds.x, 80);
    assert.equal(snapshot.selectedSheet.shapes[0]?.bounds.y, 90);

    app.bringSelectedDrawingForward();
    app.removeSelectedDrawing();
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.shapes.length, 0);
    assert.equal(snapshot.selectedFloatingId, null);
  });

  it('insertQuickShape creates a drawable rectangle', () => {
    const app = new SpreadsheetApplication();
    app.insertQuickShape('rectangle');
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.shapes.length, 1);
    assert.equal(snapshot.selectedSheet.shapes[0]?.type, 'rectangle');
  });
});
