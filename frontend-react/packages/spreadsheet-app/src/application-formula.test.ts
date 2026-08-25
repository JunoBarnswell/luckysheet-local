import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPasteSpecialSpec } from '@react-sheets/sheet-features';
import { WorkbookSession } from './workbook-session';

function cellValue(app: WorkbookSession, row: number, column: number): string {
  return app.getUiSnapshot().selectedSheet.getCell(row, column)?.value ?? '';
}

describe('WorkbookSession formula integration', () => {
  it('recalculates dependent formulas automatically when source values change', async () => {
    const app = new WorkbookSession();
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
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '6');

    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 5 },
    });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('manual recalculation mode defers updates until recalculateFormulas()', async () => {
    const app = new WorkbookSession();
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

    await app.recalculateFormulas();
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('resolves defined names through workbook.name.set', async () => {
    const app = new WorkbookSession();
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
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 3, 3), '10');

    app.runCommand('workbook.name.set', { name: 'TaxRate', value: '0.2' });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 3, 3), '20');
  });

  it('exposes canonical defined-name CRUD and calculation state through Session APIs', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const created = app.setDefinedName({ name: 'LocalRate', formula: '0.25', scope: 'sheet', sheetId });
    assert.equal(created.scope, 'sheet');
    assert.equal(app.getDefinedName('LocalRate', sheetId)?.formula, '0.25');
    assert.equal(app.listDefinedNames(sheetId).some((entry) => entry.name === 'LocalRate'), true);

    app.setRecalculationMode('manual');
    assert.equal(app.getRecalculationMode(), 'manual');
    app.removeDefinedName('LocalRate', 'sheet', sheetId);
    assert.equal(app.getDefinedName('LocalRate', sheetId), undefined);
    assert.equal(app.hasPendingFormulaRecalculation(), false);
    await app.waitForFormulaCalculation();
  });

  it('resolves Go To through the active sheet scoped name before workbook scope', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.setDefinedName({ name: 'Target', formula: 'C3', scope: 'workbook' });
    app.setDefinedName({ name: 'Target', formula: 'D4', scope: 'sheet', sheetId });
    app.runCommand('navigation.goto', { sheetId, reference: 'Target' });
    assert.deepEqual(app.getSelection().activeCell, { row: 3, column: 3 });
  });

  it('tracks dynamic-array spill ranges and child values', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=SEQUENCE(2,2,1,1)' },
    });

    await app.waitForFormulaCalculation();
    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.spillRanges.length, 1);
    assert.equal(sheet.spillRanges[0]?.state, 'ok');
    assert.equal(cellValue(app, 0, 0), '1');
    assert.equal(cellValue(app, 0, 1), '2');
    assert.equal(cellValue(app, 1, 0), '3');
    assert.equal(cellValue(app, 1, 1), '4');
  });

  it('synchronizes formulas after range paste and cell insert shifts', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 2 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A1*3', value: null } });
    app.runCommand('sheet.range.paste', {
      sheetId,
      targetOrigin: { row: 1, column: 0 },
      clipboard: { range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, values: [[{ value: 4 }]], transfer: 'copy', rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] } },
      transfer: 'copy',
      spec: createPasteSpecialSpec(),
    });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(1, 0)?.value, 4);

    app.runCommand('sheet.cells.insert', {
      sheetId,
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      operation: 'insert',
      axis: 'row',
    });
    const moved = app['runtime'].model.getSheet(sheetId).cells.get(1, 1);
    assert.equal(moved?.formula, '=A2*3');
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 1, 1), '6');
  });

  it('synchronizes clear and emits #REF! after a deleted-row reference', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 7 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A2*2', value: null } });
    app.runCommand('sheet.range.clear', {
      sheetId,
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      mode: 'contents',
    });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.formula, '=A2*2');

    app.runCommand('sheet.rows.delete', { sheetId, at: 1, count: 1 });
    const formula = app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.formula;
    assert.equal(formula, '=#REF!*2');
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '#REF!');
  });

  it('rejects direct writes into a dynamic-array spill child and rolls back the model', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=SEQUENCE(2,2,1,1)', value: null } });
    await app.waitForFormulaCalculation();
    assert.throws(() => app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 99 } }), /Spill cells are read-only/);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1), undefined);
    assert.equal(cellValue(app, 0, 1), '2');
  });
});
