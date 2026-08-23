import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';

test('manual recalculation defers dependent formula updates until recalculate()', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  assert.equal(engine.getCellValue('B1'), 6);

  engine.setValue('A1', 5);
  assert.equal(engine.getCellValue('B1'), 6);
  assert.equal(engine.hasPendingRecalculation(), true);

  engine.recalculate();
  assert.equal(engine.getCellValue('B1'), 15);
  assert.equal(engine.hasPendingRecalculation(), false);
});

test('partial recalculate updates only the requested dependency subtree', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 1);
  engine.setFormula('B1', '=A1+1');
  engine.setFormula('C1', '=B1+1');
  engine.setValue('A1', 9);
  engine.recalculateCell('B1');
  assert.equal(engine.getCellValue('B1'), 10);
  assert.equal(engine.getCellValue('C1'), 3);
});
