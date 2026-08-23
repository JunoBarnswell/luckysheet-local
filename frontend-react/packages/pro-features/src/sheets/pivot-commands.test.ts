import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerPivotFeature } from '../../../spreadsheet-app/src/features/pivot';

describe('pivot commands', () => {
  it('adds and updates a semantic pivot without writing ordinary worksheet cells', () => {
    const workbook = new WorkbookModel('pivot-command-test', 'Pivot Commands');
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    const pivot = {
      id: 'pivot-1',
      sheetId: 'sheet-1',
      sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      layout: { rows: [{ field: 'Region' }], columns: [], filters: [], values: [{ field: 'Amount', summarizeBy: 'sum' as const }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
    };
    runtime.execute('pivot.add', pivot);
    runtime.execute('pivot.update', { sheetId: 'sheet-1', pivotId: pivot.id, layout: { ...pivot.layout, rows: [{ field: 'Region' }, { field: 'Owner' }] } });
    assert.equal(workbook.getSheet('sheet-1').pivots[0]?.layout.rows.length, 2);
    assert.equal(workbook.getSheet('sheet-1').cells.count(), 0);
    runtime.undo();
    assert.equal(workbook.getSheet('sheet-1').pivots[0]?.layout.rows.length, 1);
  });

  it('drill-down is an add/remove sheet transaction with a real inverse', () => {
    const workbook = new WorkbookModel('pivot-drilldown-test', 'Pivot Drilldown');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Amount'], ['East', 10], ['West', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const runtime = new CommandRuntime(workbook);
    registerPivotFeature(runtime);
    runtime.execute('pivot.add', {
      id: 'pivot-1',
      sheetId: 'sheet-1',
      sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      layout: { rows: [{ field: 'Region' }], columns: [], filters: [], values: [{ field: 'Amount', summarizeBy: 'sum' as const }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
    });
    runtime.execute('pivot.drillDown', {
      sheetId: 'sheet-1',
      pivotId: 'pivot-1',
      label: 'East',
      sourceRowPaths: [{ sheetId: 'sheet-1', row: 1 }],
      targetSheetId: 'drill-1',
      targetAnchor: { row: 0, column: 0 },
    });
    assert.equal(workbook.getSheet('drill-1').cells.get(0, 0)?.value, 'Region');
    assert.equal(workbook.getSheet('drill-1').cells.get(1, 1)?.value, 10);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.sheets.has('drill-1'), false);
    assert.equal(runtime.redo(), true);
    assert.equal(workbook.getSheet('drill-1').cells.get(1, 0)?.value, 'East');
  });
});
