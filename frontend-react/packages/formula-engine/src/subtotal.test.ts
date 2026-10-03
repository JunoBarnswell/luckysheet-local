import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine, isFormulaError, createSnapshotVisibilityResolver,
  consumeBrowserCalculationTask, CALCULATION_TASK_PROTOCOL, CALCULATION_TASK_VERSION,
  type FormulaValue,
} from './index';

function values(engine: FormulaEngine): readonly FormulaValue[] {
  const worker = consumeBrowserCalculationTask({ snapshot: structuredClone(engine.exportCalculationSnapshot()),
    protocol: CALCULATION_TASK_PROTOCOL, version: CALCULATION_TASK_VERSION, taskId: 'subtotal', kind: 'recalculate', revision: 1, full: true });
  assert.equal(worker.status, 'completed', JSON.stringify(worker.error));
  const cell = worker.report?.results.find(cell => cell.address.row === 0 && cell.address.column === 25);
  assert.ok(cell, 'The Worker must calculate the actual SUBTOTAL formula');
  return [engine.getCellValue('Z1'), cell.value];
}
function expectValues(engine: FormulaEngine, expected: number | string): void {
  for (const value of values(engine)) {
    if (typeof expected === 'number') {
      assert.equal(typeof value, 'number', JSON.stringify(value));
      assert.ok(Math.abs(Number(value) - expected) <= Math.max(1, Math.abs(expected)) * 1e-13, `${JSON.stringify(value)} != ${expected}`);
    } else { assert.ok(isFormulaError(value)); assert.equal(value.code, expected); }
  }
}

test('all eleven subtotal functions use Excel reference coercion and exclude nested totals in inline and Worker', () => {
  const expected = [3, 2, 5, 4, 2, 8, Math.SQRT2, 1, 6, 2, 1];
  const engine = new FormulaEngine();
  engine.setValue('A1', 2); engine.setValue('A2', 4); engine.setValue('A3', null);
  engine.setValue('A4', '7'); engine.setValue('A5', true); engine.setFormula('A6', '=""');
  engine.setValue('B1', 100); engine.setFormula('A7', '=SUBTOTAL(9,B1:B1)'); engine.setFormula('A8', '=AGGREGATE(9,0,B1:B1)');
  for (let code = 1; code <= 11; code++) {
    engine.setFormula('Z1', `=SUBTOTAL(${code},A1:A8)`);
    expectValues(engine, expected[code - 1]!);
  }
});

test('subtotal uses three independent visibility reasons for both function-number families in inline and Worker', () => {
  const visibility = createSnapshotVisibilityResolver({ revision: 1, rows: [
    { sheetId: 'Sheet1', row: 1, manualHidden: true, filterHidden: false, outlineHidden: false },
    { sheetId: 'Sheet1', row: 2, manualHidden: false, filterHidden: false, outlineHidden: true },
    { sheetId: 'Sheet1', row: 3, manualHidden: false, filterHidden: true, outlineHidden: false },
  ] });
  const engine = new FormulaEngine({ rowVisibilityResolver: visibility });
  for (const [index, value] of [2, 4, 8, 16].entries()) engine.setValue(`A${index + 1}`, value);
  const ordinary = [14 / 3, 3, 3, 8, 2, 64, Math.sqrt(28 / 3), Math.sqrt(56 / 9), 14, 28 / 3, 56 / 9];
  const excluding = [2, 1, 1, 2, 2, 2, '#DIV/0!', 0, 2, '#DIV/0!', 0];
  for (let code = 1; code <= 11; code++) {
    engine.setFormula('Z1', `=SUBTOTAL(${code},A1:A4)`); expectValues(engine, ordinary[code - 1]!);
    engine.setFormula('Z1', `=SUBTOTAL(${code + 100},A1:A4)`); expectValues(engine, excluding[code - 1]!);
  }
});

test('subtotal preserves empty and error behavior for every aggregate in inline and Worker', () => {
  const engine = new FormulaEngine();
  for (let code = 1; code <= 11; code++) {
    engine.setFormula('Z1', `=SUBTOTAL(${code},A1:A3)`);
    expectValues(engine, [1, 7, 8, 10, 11].includes(code) ? '#DIV/0!' : 0);
  }
  engine.setValue('A1', 2); engine.setFormula('A2', '=1/0'); engine.setFormula('A3', '=""');
  for (let code = 1; code <= 11; code++) {
    engine.setFormula('Z1', `=SUBTOTAL(${code},A1:A3)`);
    expectValues(engine, code === 2 ? 1 : code === 3 ? 3 : '#DIV/0!');
  }
});
