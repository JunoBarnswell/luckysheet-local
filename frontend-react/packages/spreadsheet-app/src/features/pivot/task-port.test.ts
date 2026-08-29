import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { pivotSourceIdentity, WorkbookModel } from '@react-sheets/core-model';
import { preparePivotTaskInput } from './engine';
import { buildPivotModel } from './helpers';
import { BrowserPivotTaskPort, InlinePivotTaskPort, type PivotBrowserWorker } from './task-port';
import { createPivotCalculateRequest, createPivotSourceRegisterRequest, type PivotTaskRequest } from './task-protocol';
import { PivotTaskEvaluator } from './task-worker-entry';

type WorkerListener = (event: { readonly data?: unknown; readonly message?: string }) => void;

class FakePivotBrowserWorker implements PivotBrowserWorker {
  readonly listeners = new Map<'message' | 'error' | 'messageerror', Set<WorkerListener>>();
  terminated = false;
  private readonly evaluator = new PivotTaskEvaluator();

  constructor(private readonly calculateDelayMs: number) {}

  postMessage(message: unknown): void {
    const request = message as PivotTaskRequest;
    const delay = request.kind === 'calculate' ? this.calculateDelayMs : 0;
    setTimeout(() => {
      if (this.terminated) return;
      const result = this.evaluator.consume(request);
      this.listeners.get('message')?.forEach((listener) => listener({ data: result }));
    }, delay);
  }

  terminate(): void {
    this.terminated = true;
  }

  addEventListener(type: 'message' | 'error' | 'messageerror', listener: WorkerListener): void {
    const listeners = this.listeners.get(type) ?? new Set<WorkerListener>();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: WorkerListener): void {
    this.listeners.get(type)?.delete(listener);
  }
}

function preparedTask() {
  const workbook = new WorkbookModel('pivot-task', 'Pivot task');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 30]].forEach((values, row) => {
    values.forEach((value, column) => sheet.cells.set(row, column, { value }));
  });
  const pivot = buildPivotModel(workbook, sheet.id, 'pivot-task-1', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
  assert.ok(pivot);
  const region = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!;
  const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
  pivot.layout.rows = [{ fieldId: region.fieldId }];
  pivot.layout.values = [{ valueId: 'amount:sum', fieldId: amount.fieldId, summarizeBy: 'sum' }];
  return { input: preparePivotTaskInput(workbook, pivot), sourceIdentity: pivotSourceIdentity(pivot.source) };
}

describe('Pivot task port', () => {
  it('registers one source revision and evaluates through the worker contract', async () => {
    const port = new InlinePivotTaskPort();
    const { input, sourceIdentity } = preparedTask();
    const registered = await port.submit(createPivotSourceRegisterRequest('register-1', 1, sourceIdentity, input.revisions.sourceRevision, input.source));
    assert.equal(registered.status, 'accepted');
    const result = await port.submit(createPivotCalculateRequest('calculate-1', 1, sourceIdentity, input.definition, input.controls, input.revisions, input.targetBounds));
    assert.equal(result.status, 'completed');
    if (result.status === 'completed') assert.equal(result.result.grandTotal?.values[0], 60);
    port.dispose();
  });

  it('rejects a calculation whose source revision is not registered', async () => {
    const port = new InlinePivotTaskPort();
    const { input, sourceIdentity } = preparedTask();
    await port.submit(createPivotSourceRegisterRequest('register-2', 1, sourceIdentity, input.revisions.sourceRevision, input.source));
    const result = await port.submit(createPivotCalculateRequest('calculate-2', 2, sourceIdentity, input.definition, input.controls, { ...input.revisions, sourceRevision: 'stale' }, input.targetBounds));
    assert.equal(result.status, 'failed');
    if (result.status === 'failed') assert.equal(result.error.code, 'PIVOT_TASK_REVISION_MISMATCH');
    port.dispose();
  });

  it('settles a queued task as cancelled and never publishes its result', async () => {
    const port = new InlinePivotTaskPort();
    const { input, sourceIdentity } = preparedTask();
    await port.submit(createPivotSourceRegisterRequest('register-3', 1, sourceIdentity, input.revisions.sourceRevision, input.source));
    const pending = port.submit(createPivotCalculateRequest('calculate-3', 3, sourceIdentity, input.definition, input.controls, input.revisions, input.targetBounds));
    port.cancel('calculate-3');
    assert.equal((await pending).status, 'cancelled');
    port.dispose();
  });

  it('interrupts an in-flight browser calculation by replacing the worker', async () => {
    const workers: FakePivotBrowserWorker[] = [];
    const createWorker = () => {
      const worker = new FakePivotBrowserWorker(100);
      workers.push(worker);
      return worker;
    };
    const port = new BrowserPivotTaskPort(createWorker(), 1_000, createWorker);
    const { input, sourceIdentity } = preparedTask();
    assert.equal((await port.submit(createPivotSourceRegisterRequest('browser-register', 1, sourceIdentity, input.revisions.sourceRevision, input.source))).status, 'accepted');
    const pending = port.submit(createPivotCalculateRequest('browser-cancel', 1, sourceIdentity, input.definition, input.controls, input.revisions, input.targetBounds));
    port.cancel('browser-cancel');
    assert.equal((await pending).status, 'cancelled');
    assert.equal(workers[0]!.terminated, true);
    assert.equal(port.sourceRegistrationEpoch, 1);
    port.dispose();
  });

  it('fails timed-out browser work and releases the blocked worker turn', async () => {
    const workers: FakePivotBrowserWorker[] = [];
    const createWorker = () => {
      const worker = new FakePivotBrowserWorker(100);
      workers.push(worker);
      return worker;
    };
    const port = new BrowserPivotTaskPort(createWorker(), 20, createWorker);
    const { input, sourceIdentity } = preparedTask();
    assert.equal((await port.submit(createPivotSourceRegisterRequest('browser-timeout-register', 1, sourceIdentity, input.revisions.sourceRevision, input.source))).status, 'accepted');
    const timedOut = await port.submit(createPivotCalculateRequest('browser-timeout', 1, sourceIdentity, input.definition, input.controls, input.revisions, input.targetBounds));
    assert.equal(timedOut.status, 'failed');
    if (timedOut.status === 'failed') assert.equal(timedOut.error.code, 'PIVOT_TASK_TIMEOUT');
    assert.equal(workers[0]!.terminated, true);
    assert.equal(port.sourceRegistrationEpoch, 1);
    port.dispose();
  });
});
