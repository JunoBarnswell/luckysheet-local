import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { FormulaEngine } from '../formula-engine';

test('SJS.TABLE returns a dynamic array without writing result cells', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setFormula('C1', '=B1*2');
  engine.setValue('A2', 3);
  engine.setValue('A3', 4);
  engine.setFormula('D1', '=SJS.TABLE(C1,A2:A3,B1)');
  engine.setValue('B1', 2);
  engine.recalculate();

  assert.deepEqual(engine.getCellValue('D1'), [[6], [8]]);
  assert.equal(engine.getCellValue('D2'), null);
});

test('SJS.TABLE rejects non-paired input arguments', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setFormula('D1', '=SJS.TABLE(C1,A2:A3)');

  assert.equal(engine.getCellValue('D1')?.kind, 'error');
  assert.equal(engine.getCellValue('D1')?.code, '#VALUE!');
});
