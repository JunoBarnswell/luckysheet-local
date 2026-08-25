import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { PivotModel } from '@react-sheets/core-model';
import { WorkbookSession } from './workbook-session';

function seed(app: WorkbookSession): { sheetId: string; pivot: PivotModel } {
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
  const range = { sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  const fields = app.getPivotFieldCatalog(range);
  const region = fields.find((field) => field.name === 'Region')!;
  const amount = fields.find((field) => field.name === 'Amount')!;
  return {
    sheetId,
    pivot: {
      schema: 'PivotDefinition',
      id: 'pivot-test',
      source: { kind: 'worksheet-range', range },
      target: { sheetId, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [{ fieldId: region.fieldId }],
        columns: [],
        filters: [],
        values: [{ fieldId: amount.fieldId, summarizeBy: 'sum' }],
        showSubtotals: true,
        showGrandTotals: true,
        compact: true,
        repeatLabels: false,
        calculatedFields: [],
        calculatedItems: [],
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    },
  };
}

describe('WorkbookSession PivotTable integration', () => {
  it('creates a new worksheet and PivotTable as one history entry', () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const beforeHistory = app.getUiSnapshot().historyEntries.length;
    const pivotId = app.createPivotTable({ destination: { kind: 'new-sheet' } });
    assert.ok(pivotId);
    const createdSheetId = app.getActiveSheetId();
    assert.notEqual(createdSheetId, sheetId);
    assert.equal(app.getUiSnapshot().historyEntries.length, beforeHistory + 1);
    assert.equal(app.getUiSnapshot().selectedSheet.pivots[0]?.id, pivotId);

    app.undo();
    assert.equal(app.getUiSnapshot().sheets.some((sheet) => sheet.id === createdSheetId), false);
    app.redo();
    const restoredSheet = app.getUiSnapshot().sheets.find((sheet) => sheet.id === createdSheetId);
    assert.ok(restoredSheet);
    assert.equal(restoredSheet.pivots[0]?.id, pivotId);
  });

  it('leaves workbook and history unchanged when create preflight rejects duplicate headers', () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'Region' } });
    const before = app.getUiSnapshot();
    const pivotId = app.createPivotTable({ destination: { kind: 'new-sheet' } });
    assert.equal(pivotId, undefined);
    const after = app.getUiSnapshot();
    assert.equal(after.sheets.length, before.sheets.length);
    assert.equal(after.historyEntries.length, before.historyEntries.length);
    assert.equal(after.notice, 'Pivot source header is duplicated: Region');
  });

  it('rejects PivotTable creation for a viewer before any worksheet mutation', () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    const before = app.getUiSnapshot();
    assert.equal(app.createPivotTable({ destination: { kind: 'new-sheet' } }), undefined);
    const after = app.getUiSnapshot();
    assert.equal(after.sheets.length, before.sheets.length);
    assert.equal(after.historyEntries.length, before.historyEntries.length);
    assert.match(after.notice, /cannot perform|Permission denied/);
    assert.equal(sheetId, app.getActiveSheetId());
  });

  it('addPivot computes a local result tree', () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    app.addPivot(pivot);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivots.length, 1);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]!.rows.length > 0);
  });

  it('insertPivotFromSelection builds a PivotTable from the current selection', () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const pivotId = app.insertPivotFromSelection();
    assert.ok(pivotId);
    assert.ok(app.getUiSnapshot().selectedSheet.pivotResults[pivotId!]);
  });

  it('drillDownPivot creates a detail worksheet through the canonical command', () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-drill';
    app.addPivot(pivot);
    const beforeCount = app.getUiSnapshot().sheets.length;
    app.drillDownPivot(pivot.id, 'East', [{ sheetId, row: 1 }]);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.length, beforeCount + 1);
    assert.notEqual(snapshot.activeSheetId, sheetId);
  });

  it('creates a slicer drawing and refreshes a derived result without persisted refresh state', () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    pivot.id = 'pivot-slicer';
    app.addPivot(pivot);
    app.createPivotSlicerControl(pivot.id, pivot.fieldCatalog.fields[0]!.fieldId);
    app.refreshPivot(pivot.id);
    assert.equal(app.listPivotControls(pivot.id).length, 1);
    assert.ok(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]);
  });

  it('recomputes the Pivot projection when pivot.update enters through the public dispatch path', () => {
    const app = new WorkbookSession();
    const { sheetId, pivot } = seed(app);
    pivot.id = 'pivot-dispatch';
    app.addPivot(pivot);
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    const nextLayout = structuredClone(pivot.layout);
    nextLayout.values = [{ fieldId: amount.fieldId, summarizeBy: 'count' }];
    assert.equal(app.dispatch({ commandId: 'pivot.update', params: { sheetId, pivotId: pivot.id, layout: nextLayout } }), true);
    assert.equal(app.getUiSnapshot().selectedSheet.pivotResults[pivot.id]?.grandTotal?.values[0], 2);
  });
});
