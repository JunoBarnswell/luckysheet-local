import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

describe('SpreadsheetApplication chart integration', () => {
  it('addChart routes through chart.insert and syncs drawings + legacy charts', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.addChart({
      id: 'chart-test-1',
      sheetId,
      type: 'column',
      title: 'Revenue',
      sourceRanges: [{ sheetId, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }],
      bounds: { x: 60, y: 60, width: 320, height: 220 },
      legendPosition: 'right',
      showDataLabels: true,
    });

    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.charts.length, 1);
    assert.equal(snapshot.selectedSheet.charts[0]?.title, 'Revenue');
    assert.equal(snapshot.selectedSheet.charts[0]?.legendPosition, 'right');
    assert.equal(snapshot.selectedFloatingId, 'chart-test-1');

    app.updateChartBounds('chart-test-1', { x: 120, y: 140, width: 320, height: 220 });
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.charts[0]?.bounds.x, 120);

    app.updateChartType('chart-test-1', 'line');
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.charts[0]?.type, 'line');

    app.removeChart('chart-test-1');
    snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.charts.length, 0);
    assert.equal(snapshot.selectedFloatingId, null);
  });

  it('insertQuickChart uses the current selection as source range', () => {
    const app = new SpreadsheetApplication();
    app.execute('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 1, endRow: 4, startColumn: 0, endColumn: 2 }],
      primaryRangeIndex: 0,
      primaryRowIndex: 1,
      primaryColumnIndex: 0,
    });
    app.insertQuickChart('bar');
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.charts.length, 1);
    assert.equal(snapshot.selectedSheet.charts[0]?.type, 'bar');
    assert.deepEqual(snapshot.selectedSheet.charts[0]?.sourceRanges[0], {
      sheetId: app.getActiveSheetId(),
      startRow: 1,
      endRow: 4,
      startColumn: 0,
      endColumn: 2,
    });
  });
});
