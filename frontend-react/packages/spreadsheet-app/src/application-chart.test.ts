import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import type { ChartDrawingPayload } from '@react-sheets/core-model';

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
      subtype: 'clustered',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }] },
      elements: { title: 'Revenue', legend: { visible: true, position: 'right' }, dataLabels: { visible: true }, hiddenData: 'show' },
    });

    const sheet = app['runtime'].model.getSheet(sheetId);
    const drawing = sheet.drawings.find((entry) => entry.payloadId === 'chart-test-1');
    assert.equal(drawing?.kind, 'chart');
    assert.equal(sheet.drawingPayloads.get('chart-test-1')?.kind, 'chart');
    assert.equal((sheet.drawingPayloads.get('chart-test-1') as ChartDrawingPayload).elements.title, 'Revenue');
    assert.equal((sheet.drawingPayloads.get('chart-test-1') as ChartDrawingPayload).elements.legend?.position, 'right');
    let snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedFloatingId, 'draw-chart-test-1');

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

  it('insertChart uses the current selection as source range', () => {
    const app = new WorkbookSession();
    app.runCommand('selection.set', {
      sheetId: app.getActiveSheetId(),
      ranges: [{ sheetId: app.getActiveSheetId(), startRow: 1, endRow: 4, startColumn: 0, endColumn: 2 }],
      primaryRangeIndex: 0,
      activeCell: { row: 1, column: 0 },
      anchorCell: { row: 1, column: 0 },
    });
    app.insertChart('bar');
    const snapshot = app.getUiSnapshot();
    const sheet = app['runtime'].model.getSheet(app.getActiveSheetId());
    const drawing = sheet.drawings.find((entry) => entry.kind === 'chart');
    const payload = sheet.drawingPayloads.get(drawing?.payloadId ?? '') as import('@react-sheets/core-model').ChartDrawingPayload | undefined;
    assert.equal(payload?.chartType, 'bar');
    assert.deepEqual(payload?.source.kind === 'worksheet-ranges' ? payload.source.ranges[0] : undefined, {
      sheetId: app.getActiveSheetId(),
      startRow: 1,
      endRow: 4,
      startColumn: 0,
      endColumn: 2,
    });
  });
});
