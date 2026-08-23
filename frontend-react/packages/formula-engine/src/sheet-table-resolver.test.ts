import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import type { SheetTableRef } from './sheet-table-resolver';

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
