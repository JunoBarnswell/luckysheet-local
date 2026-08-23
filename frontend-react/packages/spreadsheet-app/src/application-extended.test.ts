import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession extended integration', () => {
  it('runs goal seek through extended.whatIf.goalSeek command path', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=B1*2' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { value: 10 },
    });
    app.recalculateFormulas();

    const result = app.runGoalSeek({
      setCell: { row: 0, column: 0 },
      toValue: 100,
      byChangingCell: { row: 0, column: 1 },
    });

    assert.equal(result.status, 'converged');
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.value, 50);
    assert.equal(app.getUiSnapshot().activePanel, 'extended');
  });

  it('blocks goal seek for viewers', () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    const result = app.runGoalSeek({
      setCell: { row: 0, column: 0 },
      toValue: 1,
      byChangingCell: { row: 0, column: 1 },
    });
    assert.equal(result.status, 'not-converged');
    assert.match(result.message ?? app.getUiSnapshot().notice, /permission|viewer|Goal Seek/i);
  });

  it('runs scenario analysis through extended.whatIf.scenario command path', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=B1*2' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 10 } });
    app.recalculateFormulas();

    const result = app.runScenarioAnalysis({
      id: 'growth',
      name: 'Growth',
      changingCells: [{ row: 0, column: 1, value: 20 }],
      resultCells: [{ row: 0, column: 0 }],
    });

    assert.equal(result.status, 'completed');
    assert.equal(result.outputs[0]?.value, 40);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.value, 20);
  });

  it('runs data table through extended.whatIf.dataTable command path', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { formula: '=B1*2' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 5 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 2, value: { value: 10 } });
    app.recalculateFormulas();

    const result = app.runDataTableAnalysis({
      columnInputCell: { row: 0, column: 1 },
      tableRange: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
    });

    assert.equal(result.status, 'completed');
    assert.equal(result.filledCells, 2);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(1, 1)?.value, 10);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(1, 2)?.value, 20);
  });
});
