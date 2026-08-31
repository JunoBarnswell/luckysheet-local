import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import { normalizeDefinedNames, resolveDefinedNameSource } from './defined-names';

test('normalizeDefinedNames uppercases keys', () => {
  assert.deepEqual(normalizeDefinedNames({ TaxRate: '0.15' }), { TAXRATE: '0.15' });
});

test('resolveDefinedNameSource supports scalar, range, and formula values', () => {
  const context = {
    currentCell: { sheetId: 'Sheet1', row: 4, column: 2 },
    readCell: (address: { row: number; column: number }) => (address.row === 0 && address.column === 0 ? 10 : null),
    readRangeMatrix: () => [[1, 2], [3, 4]],
  };
  assert.equal(resolveDefinedNameSource('0.25', context), 0.25);
  assert.deepEqual(resolveDefinedNameSource('A1:B2', context), [[1, 2], [3, 4]]);
  assert.equal(resolveDefinedNameSource('=A1+5', context), 15);
});

test('FormulaEngine resolves defined names in formulas with current cell context', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNames({ TaxRate: '0.1', BaseCell: 'C3' });
  engine.setValue('C3', 100);
  engine.setFormula('D3', '=BaseCell*TaxRate');
  assert.equal(engine.getCellValue('D3'), 10);
});

test('FormulaEngine recalculates formulas when a defined name changes', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNames({ Rate: '2' });
  engine.setValue('A1', 5);
  engine.setFormula('B1', '=A1*Rate');
  assert.equal(engine.getCellValue('B1'), 10);
  engine.setDefinedNames({ Rate: '3' });
  assert.equal(engine.getCellValue('B1'), 15);
});

test('FormulaEngine resolves sheet-scoped names before workbook-scoped names', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNameModels([
    { name: 'Rate', formula: '2', scope: 'workbook' },
    { name: 'Rate', formula: '3', scope: 'sheet', sheetId: 'Sheet2' },
  ]);
  engine.setFormula({ sheetId: 'Sheet1', row: 0, column: 0 }, '=Rate');
  engine.setFormula({ sheetId: 'Sheet2', row: 0, column: 0 }, '=Rate');
  assert.equal(engine.getCellValue({ sheetId: 'Sheet1', row: 0, column: 0 }), 2);
  assert.equal(engine.getCellValue({ sheetId: 'Sheet2', row: 0, column: 0 }), 3);
});

test('scoped names survive the persistent calculation bootstrap boundary', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setDefinedNameModels([
    { name: 'Rate', formula: '2', scope: 'workbook' },
    { name: 'Rate', formula: '4', scope: 'sheet', sheetId: 'Sheet2' },
  ]);
  engine.setFormula({ sheetId: 'Sheet2', row: 0, column: 0 }, '=Rate');
  const restored = FormulaEngine.fromCalculationBootstrap(engine.exportCalculationBootstrap());
  restored.executeCalculationDelta({
    protocol: 'react-sheets.formula-delta',
    version: 1,
    kind: 'calculation.delta',
    sessionId: 'scoped-name-session',
    taskId: 'scoped-name-delta',
    revision: 1,
    generation: 1,
    delta: {},
    forceRecalculate: true,
  });
  assert.equal(restored.getCellValue({ sheetId: 'Sheet2', row: 0, column: 0 }), 4);
});

test('FormulaEngine recalculates volatile formulas when dependencies change', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setFormula('A1', '=RAND()');
  const first = engine.getCellValue('A1');
  const second = engine.getCellValue('A1');
  assert.equal(first, second);
  engine.setValue('B1', 1);
  const third = engine.getCellValue('A1');
  assert.notEqual(second, third);
});
