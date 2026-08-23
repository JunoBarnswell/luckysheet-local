import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

function selectRange(
  app: WorkbookSession,
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): void {
  const sheetId = app.getActiveSheetId();
  app.runCommand('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow, endRow, startColumn, endColumn }],
    primaryRangeIndex: 0,
    primaryRowIndex: startRow,
    primaryColumnIndex: startColumn,
  });
}

describe('WorkbookSession data objects integration', () => {
  it('createSheetTableFromSelection registers a sheet table with headers', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'Name' }, { value: 'Amount' }],
        [{ value: 'Alpha' }, { value: 10 }],
        [{ value: 'Beta' }, { value: 20 }],
      ],
    });
    selectRange(app, 0, 0, 2, 1);
    app.createSheetTableFromSelection();

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.sheetTables.length, 1);
    assert.equal(sheet.sheetTables[0]?.name, 'Table1');
    assert.equal(sheet.sheetTables[0]?.columns.length, 2);
    assert.ok(sheet.filter);
  });

  it('addConditionalFormat stores rules on the worksheet', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addConditionalFormat({
      id: 'cf-1',
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }],
      type: 'highlight',
      operator: 'greaterThan',
      value1: 100,
      style: { background: '#FFEB9C' },
    });

    const rules = app['runtime'].model.getSheet(sheetId).conditionalFormats;
    assert.equal(rules.length, 1);
    assert.equal(rules[0]?.id, 'cf-1');
    assert.equal(rules[0]?.operator, 'greaterThan');
  });

  it('addDataValidation stores validation rules on the worksheet', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.addDataValidation({
      id: 'dv-1',
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 10, startColumn: 0, endColumn: 0 }],
      type: 'list',
      formula1: 'Yes,No',
      showDropdown: true,
    });

    const rules = app['runtime'].model.getSheet(sheetId).dataValidations;
    assert.equal(rules.length, 1);
    assert.equal(rules[0]?.type, 'list');
  });

  it('sortRange reorders values through sheet.sort.multi', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.range.set', {
      sheetId,
      startRow: 0,
      startColumn: 0,
      values: [
        [{ value: 'B' }, { value: 2 }],
        [{ value: 'A' }, { value: 1 }],
        [{ value: 'C' }, { value: 3 }],
      ],
    });
    selectRange(app, 0, 0, 2, 1);
    app.sortRange([{ colIdx: 0, ascending: true }], false);

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(0, 0)?.value, 'A');
    assert.equal(sheet.cells.get(1, 0)?.value, 'B');
    assert.equal(sheet.cells.get(2, 0)?.value, 'C');
  });

  it('applyFilter sets filter criteria on the sheet', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.applyFilter(0, { selectedValues: ['East', 'West'] });

    const filter = app['runtime'].model.getSheet(sheetId).filter;
    assert.ok(filter);
    assert.deepEqual(filter?.criteria[0]?.selectedValues, ['East', 'West']);
  });
});
