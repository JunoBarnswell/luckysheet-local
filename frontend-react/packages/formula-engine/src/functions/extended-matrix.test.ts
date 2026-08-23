import test from 'node:test';
import assert from 'node:assert/strict';
import { getBuiltinFunction } from './index';
import { isFormulaError } from '../values';

test('GROUPBY aggregates by group column', () => {
  const fn = getBuiltinFunction('GROUPBY');
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
  const fn = getBuiltinFunction('PIVOTBY');
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
  const fn = getBuiltinFunction('GROUPBY');
  const result = fn!([]);
  assert.ok(isFormulaError(result));
});
