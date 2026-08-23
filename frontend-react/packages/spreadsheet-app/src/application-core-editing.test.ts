import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { copyRangeToClipboardData } from '@react-sheets/sheet-features';
import { SpreadsheetApplication } from './application';

function selectCell(app: SpreadsheetApplication, row: number, column: number): void {
  const sheetId = app.getActiveSheetId();
  app.execute('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }],
    primaryRangeIndex: 0,
    primaryRowIndex: row,
    primaryColumnIndex: column,
  });
}

describe('SpreadsheetApplication core editing integration', () => {
  it('selectAddress jumps to the requested cell', () => {
    const app = new SpreadsheetApplication();
    assert.equal(app.selectAddress('C3'), true);
    assert.equal(app.getUiSnapshot().activeCell, 'C3');
  });

  it('paste special values copies values without formulas', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 42, formula: '=21*2' },
    });
    selectCell(app, 0, 0);
    const range = app.getPrimaryRange();
    app.setClipboard({ ...copyRangeToClipboardData(app['runtime'].model, range), isCut: false });
    selectCell(app, 0, 1);
    app.pasteSpecial('values');

    const target = app['runtime'].model.getSheet(sheetId).cells.get(0, 1);
    assert.equal(target?.value, 42);
    assert.equal(target?.formula, undefined);
  });

  it('formatCells applies number format through sheet.format.set', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 1,
      value: { value: 1234.5 },
    });
    selectCell(app, 1, 1);
    app.formatCells({ numberFormat: '#,##0.00' });

    const cell = app['runtime'].model.getSheet(sheetId).cells.get(1, 1);
    assert.equal(cell?.style?.numberFormat, '#,##0.00');
  });

  it('freezeAtPrimary updates worksheet freeze panes', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 2, 1);
    app.freezeAtPrimary();

    const freeze = app['runtime'].model.getSheet(sheetId).freeze;
    assert.equal(freeze?.ySplit, 2);
    assert.equal(freeze?.xSplit, 1);
  });

  it('duplicateSheet creates a copy and switches to it', () => {
    const app = new SpreadsheetApplication();
    const sourceId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId: sourceId,
      row: 0,
      column: 0,
      value: { value: 'duplicate-me' },
    });
    const beforeCount = app.getUiSnapshot().sheets.length;
    app.duplicateSheet(sourceId);

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.length, beforeCount + 1);
    assert.notEqual(app.getActiveSheetId(), sourceId);
    assert.equal(
      app['runtime'].model.getSheet(app.getActiveSheetId()).cells.get(0, 0)?.value,
      'duplicate-me',
    );
  });

  it('selectSheet updates only transient workbook session state', () => {
    const app = new SpreadsheetApplication();
    const sourceId = app.getActiveSheetId();
    app.runCommand('sheet.add', { id: 'sheet-2', name: 'Second' });

    app.selectSheet('sheet-2');
    assert.equal(app.getActiveSheetId(), 'sheet-2');
    assert.equal(app['runtime'].model.primarySheetId, sourceId);
    assert.equal(app['runtime'].commands.getHistoryDepth().undo, 1);
  });

  it('shiftCells down moves cell contents within the selected range', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'top' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 0,
      value: { value: 'bottom' },
    });
    app.execute('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }],
      primaryRangeIndex: 0,
      primaryRowIndex: 0,
      primaryColumnIndex: 0,
    });
    app.shiftCells('down');

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(1, 0)?.value, 'top');
    assert.equal(sheet.cells.get(0, 0)?.value, undefined);
  });

  it('undo reverts the last core editing mutation', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'before-undo' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'after-edit' },
    });
    app.undo();
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0)?.value, 'before-undo');
  });
});
