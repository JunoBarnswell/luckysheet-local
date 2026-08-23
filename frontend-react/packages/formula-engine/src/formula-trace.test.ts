import assert from 'node:assert/strict';
import test from 'node:test';
import { FormulaEngine, evaluateFormulaWithTrace, parseFormula } from './index';

test('formula trace evaluates real AST nodes in order', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 2);
  engine.setValue('A2', 3);
  engine.setFormula('B1', '=SUM(A1:A2)*2');

  const trace = engine.evaluateFormulaWithTrace('B1');
  assert.ok(trace);
  assert.equal(trace?.value, 10);
  assert.ok((trace?.steps.length ?? 0) >= 5);
  assert.equal(trace?.steps.at(-1)?.expression, '=SUM(A1:A2)*2');

  const standalone = evaluateFormulaWithTrace(parseFormula('=A1+2'), {
    currentCell: { sheetId: 'Sheet1', row: 0, column: 1 },
    readCell: () => 3,
    readRange: () => [],
  });
  assert.equal(standalone.value, 5);
  assert.equal(standalone.steps.at(-1)?.value, 5);
});
