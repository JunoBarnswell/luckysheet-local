import assert from 'node:assert/strict';
import test from 'node:test';
import {
  consumeCalculationTask,
  FormulaEngine,
  installCalculationWorkerEntry,
  type CalculationTaskResult,
  type CalculationWorkerScope,
} from './index';

test('calculation worker entry consumes valid tasks without host indirection', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 4);
  engine.setFormula('B1', '=A1*2');

  const result = consumeCalculationTask(engine, {
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'worker-task-1',
    kind: 'recalculate',
    revision: 9,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  });

  assert.equal(result.status, 'completed');
  assert.equal(result.taskId, 'worker-task-1');
  assert.equal(result.revision, 9);
  assert.equal(result.report?.results.some((entry) => entry.value === 8), true);
  assert.doesNotThrow(() => JSON.stringify(result));
});

test('calculation worker entry returns a failed result for malformed tasks', () => {
  const result = consumeCalculationTask(new FormulaEngine(), { taskId: 'bad', revision: 2 });
  assert.equal(result.status, 'failed');
  assert.equal(result.taskId, 'bad');
  assert.equal(result.revision, 2);
  assert.match(result.error?.message ?? '', /protocol/i);
});

test('calculation worker entry installs and restores a direct message handler', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setValue('A1', 3);
  engine.setFormula('B1', '=A1+1');
  const sent: CalculationTaskResult[] = [];
  const previous = () => undefined;
  const scope: CalculationWorkerScope = {
    onmessage: previous,
    postMessage: (message) => sent.push(message),
  };

  const uninstall = installCalculationWorkerEntry(engine, scope);
  scope.onmessage?.({
    data: {
      protocol: 'react-sheets.formula-calculation',
      version: 1,
      taskId: 'worker-task-2',
      kind: 'recalculate',
      revision: 10,
    },
  });
  assert.equal(sent[0]?.status, 'completed');
  uninstall();
  assert.equal(scope.onmessage, previous);
});
