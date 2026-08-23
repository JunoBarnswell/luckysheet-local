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
});
