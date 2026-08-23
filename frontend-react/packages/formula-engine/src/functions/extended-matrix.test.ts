import test from 'node:test';
import assert from 'node:assert/strict';
import { createFormulaCapabilities, FormulaEngine } from '../index';
import { getBuiltinFunction } from './index';
import { isFormulaError } from '../values';

test('GROUPBY aggregates by group column', () => {
  const fn = getBuiltinFunction('GROUPBY', createFormulaCapabilities({ enabled: ['GROUPBY'] }));
  assert.ok(fn);
  const result = fn!([
    [['A'], ['A'], ['B']],
    [[10], [20], [5]],
    1,
  ]);
  assert.ok(Array.isArray(result));
  const matrix = result as unknown[][];
  assert.equal(matrix.length, 2);
});

test('PIVOTBY returns matrix with header', () => {
  const fn = getBuiltinFunction('PIVOTBY', createFormulaCapabilities({ enabled: ['PIVOTBY'] }));
  assert.ok(fn);
  const result = fn!([
    [['East'], ['East'], ['West']],
    [['Q1'], ['Q2'], ['Q1']],
    [[100], [200], [50]],
    1,
  ]);
  assert.ok(Array.isArray(result));
  const matrix = result as unknown[][];
  assert.ok(matrix.length >= 2);
});

test('GROUPBY rejects invalid args', () => {
  const fn = getBuiltinFunction('GROUPBY', createFormulaCapabilities({ enabled: ['GROUPBY'] }));
  const result = fn!([]);
  assert.ok(isFormulaError(result));
});

test('gated matrix functions fail closed by default instead of returning a stub matrix', () => {
  const groupBy = getBuiltinFunction('GROUPBY');
  const pivotBy = getBuiltinFunction('PIVOTBY');
  assert.ok(groupBy);
  assert.ok(pivotBy);

  const groupByResult = groupBy!([
    [['A'], ['A'], ['B']],
    [[10], [20], [5]],
    1,
  ]);
  const pivotByResult = pivotBy!([
    [['East'], ['East'], ['West']],
    [['Q1'], ['Q2'], ['Q1']],
    [[100], [200], [50]],
    1,
  ]);

  assert.equal(isFormulaError(groupByResult), true);
  assert.equal(isFormulaError(pivotByResult), true);
  if (!isFormulaError(groupByResult) || !isFormulaError(pivotByResult)) {
    throw new Error('Expected disabled functions to return formula errors');
  }
  assert.equal(groupByResult.code, '#BLOCKED!');
  assert.equal(pivotByResult.code, '#BLOCKED!');
  assert.match(groupByResult.message, /GROUPBY.*disabled/i);
  assert.match(pivotByResult.message, /PIVOTBY.*disabled/i);
});

test('FormulaEngine applies the capability gate at the formula execution entry', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 'A');
  engine.setValue('A2', 'A');
  engine.setValue('A3', 'B');
  engine.setValue('B1', 10);
  engine.setValue('B2', 20);
  engine.setValue('B3', 5);

  const result = engine.setFormula('C1', '=GROUPBY(A1:A3, B1:B3, 1)').value;
  assert.equal(isFormulaError(result), true);
  if (!isFormulaError(result)) throw new Error('Expected GROUPBY to be blocked');
  assert.equal(result.code, '#BLOCKED!');
  assert.match(result.message, /not implemented/i);
});

test('FormulaEngine executes an explicitly injected capability', () => {
  const engine = new FormulaEngine({
    defaultSheetId: 'Sheet1',
    capabilities: createFormulaCapabilities({ enabled: ['GROUPBY'] }),
  });
  engine.setValue('A1', 'A');
  engine.setValue('A2', 'A');
  engine.setValue('A3', 'B');
  engine.setValue('B1', 10);
  engine.setValue('B2', 20);
  engine.setValue('B3', 5);

  const result = engine.setFormula('C1', '=GROUPBY(A1:A3, B1:B3, 1)').value;
  assert.ok(Array.isArray(result));
  assert.deepEqual(result, [['A', 30], ['B', 5]]);
});
