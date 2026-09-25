import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import { normalizeSheetTables, type SheetTableRef } from './sheet-table-resolver';

const sampleTable: SheetTableRef = {
  id: 't1',
  sheetId: 'Sheet1',
  name: 'Sales',
  range: { sheetId: 'Sheet1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
  hasHeaderRow: true,
  hasTotalRow: false,
  columns: [
    { id: 'c1', name: 'Product' },
    { id: 'c2', name: 'Amount' },
  ],
};

test('Sheet Table formula context rejects duplicate workbook identities', () => {
  assert.throws(() => normalizeSheetTables([
    sampleTable,
    { ...sampleTable, id: 't2', sheetId: 'Sheet2', name: 'sales' },
  ]), /identities must be unique/);
  assert.throws(() => normalizeSheetTables([
    sampleTable,
    { ...sampleTable, sheetId: 'Sheet2', name: 'Orders' },
  ]), /identities must be unique/);
});

test('FormulaEngine resolves structured table column references', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSheetTables([sampleTable]);
  engine.setValue('A2', 'Apple');
  engine.setValue('B2', 10);
  engine.setValue('A3', 'Banana');
  engine.setValue('B3', 20);
  engine.setFormula('D1', '=SUM(Sales[Amount])');
  assert.equal(engine.getCellValue('D1'), 30);
});

test('FormulaEngine resolves this-row structured table references', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSheetTables([sampleTable]);
  engine.setValue('A2', 'Apple');
  engine.setValue('B2', 10);
  engine.setFormula('C2', '=Sales[@Amount]*2');
  assert.equal(engine.getCellValue('C2'), 20);
});

test('FormulaEngine evaluates SUBTOTAL on structured table columns', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSheetTables([sampleTable]);
  engine.setValue('A2', 'Apple');
  engine.setValue('B2', 10);
  engine.setValue('A3', 'Banana');
  engine.setValue('B3', 20);
  engine.setFormula('D1', '=SUBTOTAL(109,Sales[Amount])');
  assert.equal(engine.getCellValue('D1'), 30);
});

test('FormulaEngine resolves #Data and #Headers structured table specifiers', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSheetTables([sampleTable]);
  engine.setValue('A1', 'Product');
  engine.setValue('B1', 'Amount');
  engine.setValue('A2', 'Apple');
  engine.setValue('B2', 10);
  engine.setValue('A3', 'Banana');
  engine.setValue('B3', 20);
  engine.setFormula('D1', '=SUM(Sales[#Data])');
  assert.equal(engine.getCellValue('D1'), 30);
  engine.setFormula('D2', '=Sales[[#Headers],[Amount]]');
  assert.equal(engine.getCellValue('D2'), 'Amount');
});

test('FormulaEngine resolves #All structured table specifier', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSheetTables([sampleTable]);
  engine.setValue('A1', 'Product');
  engine.setValue('B1', 'Amount');
  engine.setValue('A2', 10);
  engine.setValue('B2', 20);
  engine.setFormula('C1', '=SUM(Sales[#All])');
  assert.equal(engine.getCellValue('C1'), 30);
});

test('FormulaEngine refreshes only formulas bound to a changed table, including through a defined name', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  const initialTable: SheetTableRef = {
    ...sampleTable,
    range: { ...sampleTable.range, endRow: 2 },
  };
  engine.setSheetTables([initialTable]);
  engine.setDefinedNameModels([{ name: 'Revenue', formula: 'SUM(Sales[Amount])', scope: 'workbook' }]);
  engine.setValue('B2', 10);
  engine.setValue('B3', 20);
  engine.setValue('B4', 30);
  engine.setFormula('D1', '=Revenue');
  engine.setFormula('E1', '=1+1');
  assert.equal(engine.getCellValue('D1'), 30);

  const report = engine.setSheetTables([{
    ...initialTable,
    range: { ...initialTable.range, endRow: 3 },
  }]);

  assert.equal(engine.getCellValue('D1'), 60);
  assert.ok(report.recalculated.some(({ row, column }) => row === 0 && column === 3));
  assert.ok(!report.recalculated.some(({ row, column }) => row === 0 && column === 4));
});
