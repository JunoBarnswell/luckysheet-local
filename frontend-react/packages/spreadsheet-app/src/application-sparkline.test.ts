import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession sparkline integration', () => {
  it('addSparkline routes through sparkline.insert', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 1,
      startColumn: 1,
      values: [[{ value: 1 }, { value: 3 }, { value: 2 }, { value: 5 }]],
    });
    app.addSparkline({
      id: 'spark-test-1',
      sheetId,
      anchor: { row: 1, column: 5 },
      sourceRange: { sheetId, startRow: 1, endRow: 1, startColumn: 1, endColumn: 4 },
      type: 'line',
      color: '#2563eb',
      negativeColor: '#ef4444',
      highlightMax: true,
      highlightMin: true,
    });

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.sparklines.length, 1);
    assert.equal(snapshot.selectedSheet.sparklines[0]?.type, 'line');
    assert.equal(snapshot.selectedSheet.sparklines[0]?.anchor.column, 5);
  });

  it('insertSparkline uses the current selection and places sparkline to the right', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 2,
      startColumn: 1,
      values: [[{ value: 4 }, { value: 6 }, { value: 8 }]],
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 2, endRow: 2, startColumn: 1, endColumn: 3 }],
      primaryRangeIndex: 0,
      activeCell: { row: 2, column: 1 },
      anchorCell: { row: 2, column: 1 },
    });
    const sparklineId = app.insertSparkline('column');
    assert.ok(sparklineId);
    const snapshot = app.getUiSnapshot();
    const sparkline = snapshot.selectedSheet.sparklines.find((entry) => entry.id === sparklineId);
    assert.ok(sparkline);
    assert.equal(sparkline.type, 'column');
    assert.equal(sparkline.anchor.row, 2);
    assert.equal(sparkline.anchor.column, 4);
  });

  it('quick insertion is one history entry and restores complete style on Undo/Redo', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 2,
      startColumn: 1,
      values: [[{ value: 4 }, { value: 6 }, { value: 8 }]],
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 2, endRow: 2, startColumn: 1, endColumn: 3 }],
      primaryRangeIndex: 0,
      activeCell: { row: 2, column: 1 },
      anchorCell: { row: 2, column: 1 },
    });
    const beforeHistory = app.getUiSnapshot().historyEntries.length;
    const sparklineId = app.insertSparkline('win-loss');
    assert.ok(sparklineId);
    const inserted = app.getUiSnapshot().selectedSheet.sparklines.find((entry) => entry.id === sparklineId);
    assert.equal(inserted?.highlightMax, true);
    assert.equal(inserted?.highlightMin, true);
    assert.equal(app.getUiSnapshot().historyEntries.length, beforeHistory + 1);

    app.undo();
    assert.equal(app.getUiSnapshot().selectedSheet.sparklines.some((entry) => entry.id === sparklineId), false);
    app.redo();
    const restored = app.getUiSnapshot().selectedSheet.sparklines.find((entry) => entry.id === sparklineId);
    assert.equal(restored?.type, 'win-loss');
    assert.equal(restored?.highlightMax, true);
    assert.equal(restored?.highlightMin, true);
  });

  it('createSparklineGroup and updateSparklineGroup sync group settings', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 1,
      startColumn: 1,
      values: [[{ value: 2 }, { value: 4 }, { value: 6 }]],
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 1, endRow: 1, startColumn: 1, endColumn: 3 }],
      primaryRangeIndex: 0,
      activeCell: { row: 1, column: 1 },
      anchorCell: { row: 1, column: 1 },
    });
    const sparklineId = app.insertSparkline('line');
    assert.ok(sparklineId);
    const groupId = app.createSparklineGroup([sparklineId!], { showAxis: true, showMarkers: true });
    app.updateSparklineGroup(groupId, { showMarkers: false });
    const snapshot = app.getUiSnapshot();
    const sparkline = snapshot.selectedSheet.sparklines.find((entry) => entry.id === sparklineId);
    assert.equal(sparkline?.groupId, groupId);
    assert.equal(sparkline?.showAxis, true);
    assert.equal(sparkline?.showMarkers, false);
  });

  it('removeSparkline deletes through sparkline.remove command contract', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [[{ value: 1 }, { value: 2 }, { value: 3 }]],
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const sparklineId = app.insertSparkline('line');
    assert.ok(sparklineId);
    app.removeSparkline(sparklineId!);
    assert.equal(app.getUiSnapshot().selectedSheet.sparklines.length, 0);
  });
});
