import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

function selectRange(
  app: SpreadsheetApplication,
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): void {
  const sheetId = app.getActiveSheetId();
  app.execute('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow, endRow, startColumn, endColumn }],
    primaryRangeIndex: 0,
    primaryRowIndex: startRow,
    primaryColumnIndex: startColumn,
  });
}

describe('SpreadsheetApplication data tools integration', () => {
  it('textToColumnsFromSelection splits delimited text into columns', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'a,b,c' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 0,
      value: { value: '1,2,3' },
    });
    selectRange(app, 0, 0, 1, 0);
    app.textToColumnsFromSelection(',');

    const sheet = app.getWorkbook().getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'a');
    assert.equal(sheet.cells.get(0, 1)?.value, 'b');
    assert.equal(sheet.cells.get(0, 2)?.value, 'c');
    assert.equal(sheet.cells.get(1, 0)?.value, '1');
    assert.equal(sheet.cells.get(1, 2)?.value, '3');
  });

  it('removeDuplicatesFromSelection keeps unique rows', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Key' }, { value: 'Value' }],
        [{ value: 'A' }, { value: 1 }],
        [{ value: 'A' }, { value: 1 }],
        [{ value: 'B' }, { value: 2 }],
      ],
    });
    selectRange(app, 0, 0, 3, 1);
    app.removeDuplicatesFromSelection();

    const sheet = app.getWorkbook().getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'Key');
    assert.equal(sheet.cells.get(1, 0)?.value, 'A');
    assert.equal(sheet.cells.get(2, 0)?.value, 'B');
    assert.equal(sheet.cells.get(3, 0)?.value, undefined);
  });

  it('groupRowsFromSelection adds an outline group to the sheet model', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    selectRange(app, 1, 0, 4, 3);
    app.groupRowsFromSelection();

    const groups = app.getWorkbook().getSheet(sheetId).outline?.groups ?? [];
    assert.equal(groups.length, 1);
    assert.equal(groups[0]?.axis, 'row');
    assert.equal(groups[0]?.start, 1);
    assert.equal(groups[0]?.end, 4);
  });

  it('transposeSelection swaps rows and columns through matrix.transpose', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'A' }, { value: 'B' }],
        [{ value: 1 }, { value: 2 }],
      ],
    });
    selectRange(app, 0, 0, 1, 1);
    app.transposeSelection();

    const sheet = app.getWorkbook().getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'A');
    assert.equal(sheet.cells.get(0, 1)?.value, 1);
    assert.equal(sheet.cells.get(1, 0)?.value, 'B');
    assert.equal(sheet.cells.get(1, 1)?.value, 2);
  });

  it('applyDataSubtotal writes grouped summary rows', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Group' }, { value: 'Amount' }],
        [{ value: 'East' }, { value: 10 }],
        [{ value: 'East' }, { value: 5 }],
        [{ value: 'West' }, { value: 7 }],
      ],
    });
    selectRange(app, 0, 0, 3, 1);
    app.applyDataSubtotal();

    const sheet = app.getWorkbook().getSheet(sheetId);
    assert.equal(sheet.cells.get(5, 0)?.value, 'Group');
    assert.equal(sheet.cells.get(6, 0)?.value, 'East');
    assert.equal(sheet.cells.get(6, 1)?.value, 15);
    assert.equal(sheet.cells.get(7, 0)?.value, 'West');
    assert.equal(sheet.cells.get(7, 1)?.value, 7);
  });
});
