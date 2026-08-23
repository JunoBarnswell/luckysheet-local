import assert from 'node:assert/strict';
import test from 'node:test';
import {
  consumeBrowserCalculationTask,
  consumeCalculationTask,
  FormulaEngine,
  installCalculationWorkerEntry,
  type CalculationBrowserWorker,
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

test('browser task port posts a calculation snapshot to a Worker and applies a manual partial result', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  engine.setValue('A1', 5);
  assert.equal(engine.getCellValue('B1'), 6);

  const worker = new ImmediateCalculationWorker();
  const port = engine.createCalculationTaskPort({ workerFactory: () => worker });
  const result = await port.submit({
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'browser-worker-task',
    kind: 'recalculate',
    revision: 11,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  });

  assert.equal(result.status, 'completed');
  assert.equal(worker.calculationPosts, 1);
  assert.equal(engine.getCellValue('B1'), 15);
  port.dispose?.();
});

test('recalculateAsync uses the supplied Worker task port instead of synchronous evaluation', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  engine.setValue('A1', 5);
  const worker = new ImmediateCalculationWorker();
  const port = engine.createCalculationTaskPort({ workerFactory: () => worker });

  const report = await engine.recalculateAsync('A1', port);

  assert.equal(worker.calculationPosts, 1);
  assert.equal([...report.results.values()].some((result) => result.value === 15), true);
  assert.equal(engine.getCellValue('B1'), 15);
  port.dispose?.();
});

test('late browser Worker output cannot overwrite a newer formula input generation', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  engine.setValue('A1', 5);
  const worker = new DeferredCalculationWorker();
  const port = engine.createCalculationTaskPort({ workerFactory: () => worker });
  const pending = port.submit({
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'stale-worker-task',
    kind: 'recalculate',
    revision: 12,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  });

  engine.setValue('A1', 8);
  worker.completeNext();
  const result = await pending;

  assert.equal(result.status, 'completed');
  assert.equal(engine.getCellValue('B1'), 6);
  port.dispose?.();
});

test('browser task cancellation settles immediately and ignores a late Worker result', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  engine.setValue('A1', 5);
  const worker = new DeferredCalculationWorker();
  const port = engine.createCalculationTaskPort({ workerFactory: () => worker });
  const pending = port.submit({
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'cancelled-worker-task',
    kind: 'recalculate',
    revision: 13,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  });

  port.cancel('cancelled-worker-task');
  const result = await pending;
  worker.completeNext();

  assert.equal(result.status, 'cancelled');
  assert.equal(engine.getCellValue('B1'), 6);
  port.dispose?.();
});

test('browser Worker snapshots calculate GROUPBY with the same default semantics', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setSpillEnvironment('Sheet1', { rowCount: 12, columnCount: 8, isOccupied: () => false });
  engine.setValue('A1', 'East');
  engine.setValue('A2', 'East');
  engine.setValue('A3', 'West');
  engine.setValue('B1', 10);
  engine.setValue('B2', 20);
  engine.setValue('B3', 5);
  engine.setFormula('D1', '=GROUPBY(A1:A3,B1:B3,SUM)');
  const worker = new ImmediateCalculationWorker();
  const port = engine.createCalculationTaskPort({ workerFactory: () => worker });

  const result = await port.submit({
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'groupby-worker-task',
    kind: 'recalculate',
    revision: 14,
    roots: [{ sheetId: 'Sheet1', row: 0, column: 3 }],
  });

  assert.equal(result.status, 'completed');
  assert.equal(worker.calculationPosts, 1);
  assert.deepEqual(engine.getCellResult('D1')?.value, [['East', 30], ['West', 5]]);
  port.dispose?.();
});

abstract class BaseCalculationWorker implements CalculationBrowserWorker {
  private readonly listeners = {
    message: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
    error: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
    messageerror: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
  };

  abstract postMessage(message: unknown): void;

  terminate(): void {
    for (const listeners of Object.values(this.listeners)) listeners.clear();
  }

  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    this.listeners[type].add(listener);
  }

  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void {
    this.listeners[type].delete(listener);
  }

  protected emit(type: 'message' | 'error' | 'messageerror', event: { readonly data?: unknown; readonly message?: string }): void {
    for (const listener of this.listeners[type]) listener(event);
  }
}

class ImmediateCalculationWorker extends BaseCalculationWorker {
  calculationPosts = 0;

  postMessage(message: unknown): void {
    if (isCalculationMessage(message)) {
      this.calculationPosts += 1;
      queueMicrotask(() => this.emit('message', { data: consumeBrowserCalculationTask(message) }));
    }
  }
}

class DeferredCalculationWorker extends BaseCalculationWorker {
  private pending: unknown[] = [];

  postMessage(message: unknown): void {
    if (isCalculationMessage(message)) this.pending.push(message);
  }

  completeNext(): void {
    const message = this.pending.shift();
    if (!message) throw new Error('Expected a queued calculation task');
    this.emit('message', { data: consumeBrowserCalculationTask(message) });
  }
}

function isCalculationMessage(value: unknown): boolean {
  return typeof value === 'object'
    && value !== null
    && 'kind' in value
    && (value as { kind?: unknown }).kind === 'recalculate';
}
