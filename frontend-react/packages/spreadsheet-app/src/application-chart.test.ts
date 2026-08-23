import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession chart integration', () => {
  it('addChart routes through the canonical drawing aggregate', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addChart({
      id: 'draw-chart-test-1',
      sheetId,
      kind: 'chart',
      payloadId: 'chart-test-1',
      anchor: { kind: 'absolute' },
      transform: { x: 60, y: 60, width: 320, height: 220, rotation: 0 },
      zIndex: 0,
    }, {
      kind: 'chart',
      chartId: 'chart-test-1',
      chartType: 'column',
      title: 'Revenue',
      sourceRanges: [{ sheetId, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }],
      legendPosition: 'right',
      showDataLabels: true,
    });

    const sheet = app['runtime'].model.getSheet(sheetId);
    const drawing = sheet.drawings.find((entry) => entry.payloadId === 'chart-test-1');
    assert.equal(drawing?.kind, 'chart');
    assert.equal(sheet.drawingPayloads.get('chart-test-1')?.kind, 'chart');
    assert.equal((sheet.drawingPayloads.get('chart-test-1') as { title?: string; legendPosition?: string }).title, 'Revenue');
    assert.equal((sheet.drawingPayloads.get('chart-test-1') as { legendPosition?: string }).legendPosition, 'right');
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, 'chart-test-1');

    app.updateChartBounds('chart-test-1', { x: 120, y: 140, width: 320, height: 220 });
    snapshot = app.getUiSnapshot();
    assert.equal(sheet.drawings.find((entry) => entry.payloadId === 'chart-test-1')?.transform.x, 120);

    app.updateChartType('chart-test-1', 'line');
    snapshot = app.getUiSnapshot();
    assert.equal((sheet.drawingPayloads.get('chart-test-1') as { chartType?: string }).chartType, 'line');

    app.removeChart('chart-test-1');
    snapshot = app.getUiSnapshot();
    assert.equal(sheet.drawings.some((entry) => entry.payloadId === 'chart-test-1'), false);
    assert.equal(snapshot.selectedFloatingId, null);
  });

  it('insertQuickChart uses the current selection as source range', () => {
    const app = new WorkbookSession();
    app.runCommand('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 1, endRow: 4, startColumn: 0, endColumn: 2 }],
      primaryRangeIndex: 0,
      primaryRowIndex: 1,
      primaryColumnIndex: 0,
    });
    app.insertQuickChart('bar');
    const snapshot = app.getUiSnapshot();
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const drawing = sheet.drawings.find((entry) => entry.kind === 'chart');
    const payload = sheet.drawingPayloads.get(drawing?.payloadId ?? '') as { chartType?: string; sourceRanges?: unknown[] } | undefined;
    assert.equal(payload?.chartType, 'bar');
    assert.deepEqual(payload?.sourceRanges?.[0], {
      sheetId: app.getActiveSheetId(),
      startRow: 1,
      endRow: 4,
      startColumn: 0,
      endColumn: 2,
    });
  });
});
