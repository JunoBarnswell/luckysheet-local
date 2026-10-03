import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine, FUNCTION_ARGUMENT_CONTRACTS, isFormulaError, isFormulaInputFault,
  consumeBrowserCalculationTask, CALCULATION_TASK_PROTOCOL, CALCULATION_TASK_VERSION,
  createSnapshotVisibilityResolver,
  type FormulaValue, type ExternalCalculationLink,
} from '../index';
import { FINANCIAL_FUNCTION_CORPUS } from '../fixtures/financial-corpus';

function close(actual: FormulaValue, expected: number): void {
  assert.equal(typeof actual, 'number', JSON.stringify(actual));
  assert.ok(Math.abs((actual as number) - expected) <= Math.max(1, Math.abs(expected)) * 1e-13, `${actual} != ${expected}`);
}
function workerValue(engine: FormulaEngine): FormulaValue {
  const result = consumeBrowserCalculationTask({
    snapshot: structuredClone(engine.exportCalculationSnapshot()),
    protocol: CALCULATION_TASK_PROTOCOL, version: CALCULATION_TASK_VERSION, taskId: 'financial', kind: 'recalculate', revision: 1, full: true,
  });
  assert.equal(result.status, 'completed', JSON.stringify(result.error));
  const cell = result.report?.results.find(cell => cell.address.row === 0 && cell.address.column === 25);
  assert.ok(cell, 'Worker must actually calculate the financial formula');
  return cell.value;
}
const golden: readonly (readonly [string, number])[] = [
  ['PV(0,4,-25)', 100], ['FV(0,4,-25,100)', 0], ['PMT(0,4,100)', -25], ['NPER(0,-25,100)', 4],
  ['FV(0.5,2,-100,0,1)', 375], ['PMT(0.5,2,250,0,1)', -150], ['NPER(0.5,-100,0,375,1)', 2],
  ['IPMT(0.5,1,2,250,0,1)', 0], ['IPMT(0.5,2,2,250,0,1)', -50], ['PPMT(0.5,2,2,250,0,1)', -100],
  ['IPMT(0,2,4,100)', 0], ['PPMT(0,2,4,100)', -25], ['ISPMT(0.5,1,2,250)', -62.5],
  ['PV(-0.5,2,-50)', 300], ['FV(-0.5,2,-50)', 75], ['PMT(-0.5,2,250)', -41.6666666666667],
  ['PMT(0.000000000001,4,100)', -25.0000000000625], ['NPER(0.000000000001,-25,100)', 4.00000000001],
  ['PV(1000,1000,-100)', 0.1], ['PMT(1000,1000,100)', -100000],
  ['EFFECT(0.2,2.9)', 0.21], ['NOMINAL(0.21,2.9)', 0.2],
  ['DOLLARDE(-1.02,16.9)', -1.125], ['DOLLARFR(-1.125,16.9)', -1.02],
  ['SLN(1000,100,-3)', -300], ['SLN(0,100,2)', -50], ['PMT("0",4,"100")', -25],
];
for (const [formula, expected] of golden) test(`financial golden ${formula}: inline and real Worker entry`, () => {
  const engine = new FormulaEngine();
  engine.setFormula('Z1', `=${formula}`);
  close(engine.getCellValue('Z1'), expected); close(workerValue(engine), expected);
});
const rejections = [
  ['PMT(0,0,100)', '#DIV/0!'], ['NPER(0,0,100)', '#DIV/0!'], ['PV(-1,2,-10)', '#DIV/0!'],
  ['FV(-2,0.5,-1)', '#NUM!'], ['FV(1000,1000,-1)', '#NUM!'], ['NPER(-1,-100,100)', '#NUM!'],
  ['NPER(0.1,-10,100)', '#NUM!'], ['IPMT(0.1,0,2,100)', '#NUM!'], ['PPMT(0.1,3,2,100)', '#NUM!'],
  ['FV(0,2,-1,0,2)', '#NUM!'], ['PV(0,2,-1,0,-1)', '#NUM!'], ['PMT(0,2,100,0,2)', '#NUM!'],
  ['NPER(0,-1,100,0,2)', '#NUM!'], ['IPMT(0,1,2,100,0,2)', '#NUM!'], ['PPMT(0,1,2,100,0,2)', '#NUM!'],
  ['ISPMT(0.1,0,0,100)', '#DIV/0!'], ['EFFECT(0,2)', '#NUM!'], ['EFFECT(0.1,0.9)', '#NUM!'],
  ['NOMINAL(-0.1,2)', '#NUM!'], ['NOMINAL(0.1,0.9)', '#NUM!'], ['SLN(100,10,0)', '#DIV/0!'],
  ['DOLLARDE(1.02,-0.5)', '#NUM!'], ['DOLLARFR(1.02,-1)', '#NUM!'], ['DOLLARDE(1.02,0.9)', '#DIV/0!'],
  ['DOLLARFR(1.02,0)', '#DIV/0!'],
] as const;
for (const [formula, code] of rejections) test(`financial rejection ${formula}: inline and Worker errors`, () => {
  const engine = new FormulaEngine(); engine.setFormula('Z1', `=${formula}`);
  for (const value of [engine.getCellValue('Z1'), workerValue(engine)]) assert.equal(isFormulaError(value) ? value.code : value, code);
});
for (const id of Object.keys(FINANCIAL_FUNCTION_CORPUS)) test(`financial ${id}: coercion, normal errors and lower arity`, () => {
  const minimum = FUNCTION_ARGUMENT_CONTRACTS.get(id)!.minimum, engine = new FormulaEngine();
  for (const [head, code] of [['"bad"', '#VALUE!'], ['#N/A', '#N/A']] as const) {
    engine.setFormula('Z1', `=${id}(${[head, ...new Array(minimum - 1).fill('1')].join(',')})`);
    for (const value of [engine.getCellValue('Z1'), workerValue(engine)]) assert.equal(isFormulaError(value) ? value.code : value, code);
  }
  const value = engine.setFormula('Z1', `=${id}(${new Array(minimum - 1).fill('1').join(',')})`).value;
  assert.equal(isFormulaError(value) ? value.code : value, '#VALUE!');
});
test('financial scalar arrays broadcast through the one evaluator and keep shape/error behavior', () => {
  const engine = new FormulaEngine(); engine.setValue('A1', 0); engine.setValue('A2', 0.5);
  engine.setValue('B1', 100); engine.setValue('C1', 250);
  engine.setFormula('Z1', '=PMT(A1:A2,2,B1:C1)');
  const expected = [[-50, -125], [-90, -225]];
  assert.deepEqual(engine.getCellValue('Z1'), expected); assert.deepEqual(workerValue(engine), expected);
  engine.setFormula('Z1', '=PMT(A1:A2,2,A1:A3)');
  for (const value of [engine.getCellValue('Z1'), workerValue(engine)]) assert.equal(isFormulaError(value) ? value.code : value, '#VALUE!');
  engine.setFormula('A1', '=#N/A');
  engine.setFormula('Z1', '=IFERROR(PMT(A1:A2,2,100),7)');
  assert.deepEqual(engine.getCellValue('Z1'), [[7], [-90]]); assert.deepEqual(workerValue(engine), [[7], [-90]]);
});
test('interest and principal obey the same annuity owner for both payment timings', () => {
  const engine = new FormulaEngine();
  for (const type of [0, 1]) for (const period of [1, 2]) {
    engine.setFormula('Z1', `=IPMT(0.5,${period},2,250,0,${type})+PPMT(0.5,${period},2,250,0,${type})-PMT(0.5,2,250,0,${type})`);
    close(engine.getCellValue('Z1'), 0); close(workerValue(engine), 0);
  }
});
test('financial reads consume canonical values independently of every row visibility reason', () => {
  for (const reason of ['manualHidden', 'filterHidden', 'outlineHidden'] as const) {
    const engine = new FormulaEngine({ rowVisibilityResolver: createSnapshotVisibilityResolver({ revision: 1, rows: [{
      sheetId: 'Sheet1', row: 0, manualHidden: false, filterHidden: false, outlineHidden: false, [reason]: true,
    }] }) });
    engine.setValue('A1', 0.5);
    for (const fixture of Object.values(FINANCIAL_FUNCTION_CORPUS)) {
      engine.setFormula('Z1', fixture.formula.replace(/\((?:-?[\d.]+)/, '(A1'));
      const value = engine.getCellValue('Z1'); assert.equal(typeof value, 'number');
      assert.deepEqual(workerValue(engine), value);
    }
  }
});
test('all twelve financial functions retain consumed source faults across IFERROR and Worker, then recover', async () => {
  const source: ExternalCalculationLink = {
    id: 'rates', token: 'Rates.xlsx', sourceUnitId: 'rates-source', subject: 'reader', sourceRevision: 1, accessRevision: 1, state: 'connected',
    sheets: [{ id: 'rates-sheet', name: 'Rates', rowCount: 10, columnCount: 10 }],
    cells: [{ address: { sheetId: 'rates-sheet', row: 0, column: 0 }, value: 0.5 }],
  };
  for (const fixture of Object.values(FINANCIAL_FUNCTION_CORPUS)) {
    const engine = new FormulaEngine(), external = fixture.formula.replace(/\((?:-?[\d.]+)/, '([Rates.xlsx]Rates!A1');
    engine.applyExternalCalculationLinks([source]); engine.setFormula('Z1', `=IFERROR(${external.slice(1)},7)`);
    await engine.recalculateAsync(); const before = engine.getCellValue('Z1'); assert.equal(isFormulaError(before), false);
    for (const state of ['denied', 'unavailable', 'broken'] as const) {
      engine.applyExternalCalculationLinks([{ ...source, state, accessRevision: 2, sheets: [], cells: [] }]); await engine.recalculateAsync();
      for (const value of [engine.getCellValue('Z1'), workerValue(engine)]) {
        assert.ok(isFormulaInputFault(value), `${fixture.formula}: ${JSON.stringify(value)}`);
        assert.equal(value.inputFault.source, 'rates-source');
      }
    }
    engine.applyExternalCalculationLinks([{ ...source, sourceRevision: 2, accessRevision: 3 }]); await engine.recalculateAsync();
    assert.deepEqual(engine.getCellValue('Z1'), before); assert.deepEqual(workerValue(engine), before);
  }
  const blocked = new FormulaEngine({ blockedRanges: [{ sheetId: 'Sheet1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] });
  blocked.setValue('A1', 0.5); blocked.setFormula('Z1', '=IFERROR(PMT(A1,2,250),7)');
  assert.ok(isFormulaInputFault(blocked.getCellValue('Z1'))); assert.ok(isFormulaInputFault(workerValue(blocked)));
});
