import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from '../formula-engine';
import { isArrayValue, isFormulaError } from '../values';

test('dynamic array functions: FILTER UNIQUE SORT SEQUENCE', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 'East');
  engine.setValue('A2', 'West');
  engine.setValue('A3', 'East');
  engine.setValue('B1', 100);
  engine.setValue('B2', 200);
  engine.setValue('B3', 50);
  engine.setValue('C1', true);
  engine.setValue('C2', false);
  engine.setValue('C3', true);

  const filtered = engine.setFormula('D1', '=FILTER(A1:B3, C1:C3)').value;
  assert.ok(isArrayValue(filtered));
  assert.equal((filtered as unknown[][]).length, 2);

  const unique = engine.setFormula('E1', '=UNIQUE(A1:A3)').value;
  assert.ok(isArrayValue(unique));
  assert.equal((unique as unknown[][]).length, 2);

  const sorted = engine.setFormula('F1', '=SORT(A1:B3, 2, -1)').value;
  assert.ok(isArrayValue(sorted));
  assert.equal((sorted as unknown[][])[0]?.[1], 200);

  const sequence = engine.setFormula('G1', '=SEQUENCE(2, 3, 1, 1)').value;
  assert.ok(isArrayValue(sequence));
  assert.deepEqual(sequence, [
    [1, 2, 3],
    [4, 5, 6],
  ]);
});

test('dynamic array functions: XMATCH HSTACK VSTACK TAKE DROP', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 'Alpha');
  engine.setValue('A2', 'Beta');
  engine.setValue('A3', 'Gamma');
  engine.setValue('B1', 10);
  engine.setValue('B2', 20);
  engine.setValue('C1', 1);
  engine.setValue('C2', 2);

  assert.equal(engine.setFormula('D1', '=XMATCH("Beta", A1:A3, 0)').value, 2);
  assert.equal(engine.setFormula('D2', '=XMATCH(15, B1:B2, 1)').value, 2);

  const stacked = engine.setFormula('E1', '=HSTACK(A1:A2, B1:B2)').value;
  assert.ok(isArrayValue(stacked));
  assert.deepEqual(stacked, [
    ['Alpha', 10],
    ['Beta', 20],
  ]);

  const vstacked = engine.setFormula('F1', '=VSTACK(C1:C2, B1:B2)').value;
  assert.ok(isArrayValue(vstacked));
  assert.equal((vstacked as unknown[][]).length, 4);

  const taken = engine.setFormula('G1', '=TAKE(A1:B3, 2)').value;
  assert.ok(isArrayValue(taken));
  assert.equal((taken as unknown[][]).length, 2);

  const dropped = engine.setFormula('H1', '=DROP(A1:B3, 1)').value;
  assert.ok(isArrayValue(dropped));
  assert.equal((dropped as unknown[][]).length, 2);
});

test('dynamic array functions: SORTBY and RANDARRAY', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 10, columnCount: 10, isOccupied: () => false });
  engine.setValue('A1', 'b');
  engine.setValue('A2', 'a');
  engine.setValue('B1', 2);
  engine.setValue('B2', 1);

  const sorted = engine.setFormula('C1', '=SORTBY(A1:A2,B1:B2)').value;
  assert.ok(isArrayValue(sorted));
  assert.equal((sorted as unknown[][])[0]?.[0], 'a');

  const random = engine.setFormula('D1', '=RANDARRAY(2,2,1,9,1)').value;
  assert.ok(isArrayValue(random));
  assert.equal((random as unknown[][]).length, 2);
});

test('dynamic array functions: FILTER empty uses if_empty', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 1);
  engine.setValue('B1', false);
  const result = engine.setFormula('C1', '=FILTER(A1, B1, "none")').value;
  assert.deepEqual(result, [['none']]);
});

test('dynamic array functions: FILTER no match returns #CALC!', () => {
  const engine = new FormulaEngine();
  engine.setValue('A1', 1);
  engine.setValue('B1', false);
  const result = engine.setFormula('C1', '=FILTER(A1, B1)').value;
  assert.ok(isFormulaError(result));
  if (!isFormulaError(result)) throw new Error('expected error');
  assert.equal(result.code, '#CALC!');
});
