import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession pivot integration', () => {
  it('addPivot computes a local pivot result tree', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Region' }, { value: 'Amount' }],
        [{ value: 'East' }, { value: 10 }],
        [{ value: 'West' }, { value: 20 }],
      ],
    });
    app.addPivot({
      id: 'pivot-test-1',
      sheetId,
      sourceRange: { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      layout: {
        rows: [{ field: 'Region' }],
        columns: [],
        filters: [],
        values: [{ field: 'Amount', summarizeBy: 'sum' }],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
        calculatedFields: [],
        calculatedItems: [],
      },
    });

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivots.length, 1);
    assert.ok(snapshot.selectedSheet.pivotResults['pivot-test-1']);
    assert.ok(snapshot.selectedSheet.pivotResults['pivot-test-1']!.rows.length > 0);
  });

  it('insertQuickPivot builds a pivot from the current selection', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Category' }, { value: 'Qty' }],
        [{ value: 'A' }, { value: 3 }],
        [{ value: 'B' }, { value: 5 }],
      ],
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const pivotId = app.insertQuickPivot();
    assert.ok(pivotId);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivots.length, 1);
    assert.ok(snapshot.selectedSheet.pivotResults[pivotId!]);
  });

  it('drillDownPivot creates a detail sheet through pivot.drillDown', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addPivot({
      id: 'pivot-drill',
      sheetId,
      sourceRange: { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      layout: {
        rows: [{ field: 'Region' }],
        columns: [],
        filters: [],
        values: [{ field: 'Amount', summarizeBy: 'sum' }],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
        calculatedFields: [],
        calculatedItems: [],
      },
    });
    const beforeCount = app.getUiSnapshot().sheets.length;
    app.drillDownPivot('pivot-drill', 'East', [{ sheetId, row: 1 }]);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.length, beforeCount + 1);
    assert.notEqual(snapshot.activeSheetId, sheetId);
  });

  it('setPivotSlicer and refreshPivot keep results in sync', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addPivot({
      id: 'pivot-slicer',
      sheetId,
      sourceRange: { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      layout: {
        rows: [{ field: 'Region' }],
        columns: [],
        filters: [],
        values: [{ field: 'Amount', summarizeBy: 'sum' }],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
        calculatedFields: [],
        calculatedItems: [],
      },
    });
    app.setPivotSlicer('pivot-slicer', { id: 'slicer-Region', field: 'Region', selected: ['East'], connectedPivotIds: ['pivot-slicer'] });
    app.refreshPivot('pivot-slicer');
    const snapshot = app.getUiSnapshot();
    assert.ok(snapshot.selectedSheet.pivotResults['pivot-slicer']);
  });
});
