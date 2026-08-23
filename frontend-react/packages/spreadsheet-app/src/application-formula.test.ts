import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

function cellValue(app: SpreadsheetApplication, row: number, column: number): string {
  return app.getUiSnapshot().selectedSheet.getCell(row, column)?.value ?? '';
}

describe('SpreadsheetApplication formula integration', () => {
  it('recalculates dependent formulas automatically when source values change', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 2 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { formula: '=A1*3' },
    });
    assert.equal(cellValue(app, 0, 1), '6');

    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 5 },
    });
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('manual recalculation mode defers updates until recalculateFormulas()', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.setRecalculationMode('manual');
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 2 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { formula: '=A1*3' },
    });
    assert.equal(cellValue(app, 0, 1), '6');

    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 5 },
    });
    assert.equal(cellValue(app, 0, 1), '6');

    app.recalculateFormulas();
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('resolves defined names through workbook.name.set', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('workbook.name.set', { name: 'TaxRate', value: '0.1' });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 2,
      column: 2,
      value: { value: 100 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 3,
      column: 3,
      value: { formula: '=C3*TaxRate' },
    });
    assert.equal(cellValue(app, 3, 3), '10');

    app.runCommand('workbook.name.set', { name: 'TaxRate', value: '0.2' });
    assert.equal(cellValue(app, 3, 3), '20');
  });

  it('tracks dynamic-array spill ranges and child values', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=SEQUENCE(2,2,1,1)' },
    });

    const sheet = app.getWorkbook().getSheet(sheetId);
    assert.equal(sheet.spillRanges.length, 1);
    assert.equal(sheet.spillRanges[0]?.state, 'ok');
    assert.equal(cellValue(app, 0, 0), '1');
    assert.equal(cellValue(app, 0, 1), '2');
    assert.equal(cellValue(app, 1, 0), '3');
    assert.equal(cellValue(app, 1, 1), '4');
  });
});
