import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from '../index';
import { getBuiltinFunction } from './index';
import { isFormulaError } from '../values';

test('GROUPBY is enabled by default and aggregates multiple field and value columns deterministically', () => {
  const groupBy = getBuiltinFunction('GROUPBY');
  assert.ok(groupBy);

  const result = groupBy!([
    [['East', 'A'], ['East', 'A'], ['West', 'A'], ['East', 'B']],
    [[10, 1], [20, 2], [5, 3], [7, 4]],
    'SUM',
  ]);

  assert.deepEqual(result, [
    ['East', 'A', 30, 3],
    ['West', 'A', 5, 3],
    ['East', 'B', 7, 4],
  ]);
});

test('GROUPBY supports the five deterministic aggregations and treats blanks as non-numeric', () => {
  const groupBy = getBuiltinFunction('GROUPBY');
  assert.ok(groupBy);
  const fields = [['A'], ['A'], ['B'], ['B'], ['C']];
  const values = [[10], [null], [5], ['ignored'], [null]];

  assert.deepEqual(groupBy!([fields, values, 'SUM']), [['A', 10], ['B', 5], ['C', 0]]);
  assert.deepEqual(groupBy!([fields, values, 'COUNT']), [['A', 1], ['B', 1], ['C', 0]]);
  assert.deepEqual(groupBy!([fields, values, 'MIN']), [['A', 10], ['B', 5], ['C', 0]]);
  assert.deepEqual(groupBy!([fields, values, 'MAX']), [['A', 10], ['B', 5], ['C', 0]]);
  assert.deepEqual(groupBy!([fields, values, 1]), [['A', 10], ['B', 5], ['C', 0]]);

  const average = groupBy!([fields, values, 'AVERAGE']);
  assert.equal(Array.isArray(average), true);
  if (!Array.isArray(average)) throw new Error('Expected a matrix');
  assert.equal(average[0]?.[1], 10);
  assert.equal(average[1]?.[1], 5);
  assert.equal(isFormulaError(average[2]?.[1]), true);
});

test('GROUPBY uses typed keys and combines null with empty text as one blank group', () => {
  const groupBy = getBuiltinFunction('GROUPBY');
  assert.ok(groupBy);
  const result = groupBy!([
    [[null], [''], [1], ['1']],
    [[2], [3], [5], [7]],
    'SUM',
  ]);
  assert.deepEqual(result, [
    [null, 5],
    [1, 5],
    ['1', 7],
  ]);
});

test('GROUPBY validates shapes and propagates input errors instead of truncating or coercing rows', () => {
  const groupBy = getBuiltinFunction('GROUPBY');
  assert.ok(groupBy);
  const mismatch = groupBy!([[['A'], ['B']], [[1]], 'SUM']);
  assert.equal(isFormulaError(mismatch), true);
  if (!isFormulaError(mismatch)) throw new Error('Expected a shape error');
  assert.equal(mismatch.code, '#VALUE!');

  const sourceError = { kind: 'error' as const, code: '#REF!' as const, message: 'deleted' };
  const propagated = groupBy!([[['A']], [[sourceError]], 'SUM']);
  assert.deepEqual(propagated, sourceError);

  const unsupported = groupBy!([[['A']], [[1]], 'MEDIAN']);
  assert.equal(isFormulaError(unsupported), true);
  if (!isFormulaError(unsupported)) throw new Error('Expected an aggregate error');
  assert.equal(unsupported.code, '#VALUE!');
});

test('PIVOTBY builds stable row/column field headers and cross-aggregates every matching source row', () => {
  const pivotBy = getBuiltinFunction('PIVOTBY');
  assert.ok(pivotBy);

  const result = pivotBy!([
    [['East'], ['East'], ['West'], ['East']],
    [['Q1'], ['Q2'], ['Q1'], ['Q1']],
    [[100], [200], [50], [25]],
    'SUM',
  ]);

  assert.deepEqual(result, [
    [null, 'Q1', 'Q2'],
    ['East', 125, 200],
    ['West', 50, 0],
  ]);
});

test('PIVOTBY preserves multi-field and multi-value dimensions without stringifying keys', () => {
  const pivotBy = getBuiltinFunction('PIVOTBY');
  assert.ok(pivotBy);

  const result = pivotBy!([
    [['East', 2025], ['East', 2025], ['West', 2025]],
    [['Q1', true], ['Q2', true], ['Q1', true]],
    [[10, 1], [20, 2], [5, 3]],
    'SUM',
  ]);

  assert.deepEqual(result, [
    [null, null, 'Q1', 'Q1', 'Q2', 'Q2'],
    [null, null, true, true, true, true],
    [null, null, 'SUM 1', 'SUM 2', 'SUM 1', 'SUM 2'],
    ['East', 2025, 10, 1, 20, 2],
    ['West', 2025, 5, 3, 0, 0],
  ]);
});

test('PIVOTBY rejects mismatched input heights and unsupported optional arguments', () => {
  const pivotBy = getBuiltinFunction('PIVOTBY');
  assert.ok(pivotBy);
  const mismatch = pivotBy!([[['East'], ['West']], [['Q1']], [[1], [2]], 'SUM']);
  assert.equal(isFormulaError(mismatch), true);
  if (!isFormulaError(mismatch)) throw new Error('Expected a shape error');
  assert.equal(mismatch.code, '#VALUE!');

  const unsupported = pivotBy!([[['East']], [['Q1']], [[1]], 'SUM', 3]);
  assert.equal(isFormulaError(unsupported), true);
  if (!isFormulaError(unsupported)) throw new Error('Expected an unsupported-arity error');
  assert.equal(unsupported.code, '#VALUE!');
});

test('FormulaEngine accepts bare aggregate identifiers, spills GROUPBY, and computes PIVOTBY by default', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 20, columnCount: 12, isOccupied: () => false });
  engine.setValue('A1', 'East');
  engine.setValue('A2', 'East');
  engine.setValue('A3', 'West');
  engine.setValue('B1', 10);
  engine.setValue('B2', 20);
  engine.setValue('B3', 5);
  engine.setValue('C1', 'Q1');
  engine.setValue('C2', 'Q2');
  engine.setValue('C3', 'Q1');

  const grouped = engine.setFormula('E1', '=GROUPBY(A1:A3,B1:B3,SUM)').value;
  assert.deepEqual(grouped, [['East', 30], ['West', 5]]);
  assert.equal(engine.getSpillsForSheet('Sheet1')[0]?.state, 'ok');

  const pivoted = engine.setFormula('H1', '=PIVOTBY(A1:A3,C1:C3,B1:B3,COUNT)').value;
  assert.deepEqual(pivoted, [
    [null, 'Q1', 'Q2'],
    ['East', 1, 1],
    ['West', 1, 0],
  ]);
});
