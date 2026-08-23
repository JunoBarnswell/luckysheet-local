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
  it('addPivot computes a local result tree', () => {
    const app = new WorkbookSession();
    const { pivot } = seed(app);
    app.addPivot(pivot);
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.selectedSheet.pivots.length, 1);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]);
    assert.ok(snapshot.selectedSheet.pivotResults[pivot.id]!.rows.length > 0);
  });

  it('insertQuickPivot builds a PivotTable from the current selection', () => {
    const app = new WorkbookSession();
    const { sheetId } = seed(app);
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    const pivotId = app.insertQuickPivot();
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
});
