import assert from 'node:assert/strict';
import test from 'node:test';
import { FormulaEngine, FUNCTION_DESCRIPTORS, FUNCTION_LIBRARY, FUNCTION_ARGUMENT_CONTRACTS, isFormulaError, type FormulaValue } from './index';
import { CALCULATION_TASK_PROTOCOL, CALCULATION_TASK_VERSION } from './calculation-task-port';
import { FUNCTION_CORPUS } from './fixtures/function-corpus';

function equalValue(actual: FormulaValue, expected: FormulaValue): void {
  if (typeof actual === 'number' && typeof expected === 'number') { assert.ok(Math.abs(actual - expected) <= Math.max(1, Math.abs(expected)) * 1e-13, `${actual} != ${expected}`); return; }
  if (Array.isArray(actual) && Array.isArray(expected)) {
    assert.equal(actual.length, expected.length);
    actual.forEach((row, r) => { assert.equal(row.length, expected[r]!.length); row.forEach((value, c) => equalValue(value, expected[r]![c]!)); });
    return;
  }
  assert.deepEqual(actual, expected);
}
function engineForCase() {
  const engine = new FormulaEngine();
  const values = [[2, 'x', true], [4, 'y', false], [6, 'x', true]] as const;
  engine.synchronizeInputs(values.flatMap((row, r) => row.map((value, column) => ({ address: { sheetId: 'Sheet1', row: r, column }, input: { kind: 'value' as const, value } }))));
  engine.setValue('D1', 10);
  return engine;
}
const entropy = { cycleId: 1, entropySeed: 'function-corpus', passIndex: 0, calculationTimeUtcMs: Date.UTC(2024, 0, 2, 12), calculationTimeZoneOffsetMinutes: 0 };
const request = { protocol: CALCULATION_TASK_PROTOCOL, version: CALCULATION_TASK_VERSION, taskId: 'corpus', kind: 'recalculate' as const, revision: 1, full: true, calculationEntropy: entropy };

test('every executable function is discoverable, has an argument contract and a concrete acceptance vector', () => {
  const ids = [...FUNCTION_DESCRIPTORS.keys()].sort();
  assert.deepEqual(FUNCTION_LIBRARY.map(({ id }) => id).sort(), ids);
  assert.deepEqual(Object.keys(FUNCTION_CORPUS).sort(), ids);
  assert.deepEqual([...FUNCTION_ARGUMENT_CONTRACTS.keys()].sort(), ids);
  for (const id of ['SORT', 'SORTBY', 'FILTER', 'UNIQUE', 'GROUPBY', 'PIVOTBY', 'SEQUENCE', 'LET', 'LAMBDA']) assert.equal(FUNCTION_DESCRIPTORS.get(id)?.streaming, false, id);
});
for (const [id, fixture] of Object.entries(FUNCTION_CORPUS)) test(`function catalog: ${id} inline and Worker calculation`, () => {
  const engine = engineForCase();
  engine.setFormula('Z1', fixture.formula);
  engine.executeCalculationTask(request);
  const value = engine.getCellValue('Z1');
  assert.equal(isFormulaError(value), false, `${id}: ${JSON.stringify(value)}`);
  if (id === 'RAND') assert.ok(typeof value === 'number' && value >= 0 && value < 1);
  else equalValue(value, fixture.expected);
  const worker = FormulaEngine.fromCalculationSnapshot(structuredClone(engine.exportCalculationSnapshot()));
  worker.executeCalculationTask(request);
  equalValue(worker.getCellValue('Z1'), value);
  assert.equal(engine.getCellValue('D1'), 10, 'scenario evaluation must preserve the input value');
  for (const cell of ['A1', 'A2', 'A3']) assert.equal(worker.getCellValue(cell), engine.getCellValue(cell));
});
for (const id of FUNCTION_DESCRIPTORS.keys()) test(`function catalog: ${id} rejects excessive arguments`, () => {
  const engine = new FormulaEngine();
  const value = engine.setFormula('A1', `=${id}(${new Array(256).fill('1').join(',')})`).value;
  assert.equal(isFormulaError(value) ? value.code : value, '#VALUE!');
});
