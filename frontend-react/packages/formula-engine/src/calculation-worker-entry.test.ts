import assert from 'node:assert/strict';
import test from 'node:test';
import {
  CALCULATION_DELTA_PROTOCOL,
  CALCULATION_DELTA_VERSION,
  FormulaEngine,
  consumeCalculationSession,
  installBrowserCalculationWorkerEntry,
  type CalculationBrowserWorker,
  type CalculationSessionRequest,
  type CalculationSessionResult,
} from './index';

test('persistent calculation worker opens once and applies only canonical deltas afterwards', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setValue('A1', 2);
  engine.setFormula('B1', '=A1*3');
  const worker = new ImmediateCalculationWorker();
  const port = engine.createCalculationSessionPort({ workerFactory: () => worker });

  await engine.recalculateAsync(undefined, port);
  assert.equal(worker.calculationPosts, 1);
  assert.equal(engine.getCellValue('B1'), 6);

  engine.setValue('A1', 5);
  await engine.recalculateAsync('A1', port);
  assert.equal(worker.calculationPosts, 2);
  assert.equal(engine.getCellValue('B1'), 15);
  const delta = worker.messages[1] as CalculationSessionRequest;
  assert.equal(delta.kind, 'calculation.delta');
  if (delta.kind !== 'calculation.delta') throw new Error('Expected calculation delta');
  assert.deepEqual(delta.delta.cells?.map((entry) => entry.address.row), [0]);
  port.dispose();
});

test('worker entry rejects the removed snapshot task protocol', () => {
  const result = consumeCalculationSession(new Map(), {
    protocol: 'react-sheets.formula-calculation',
    version: 1,
    taskId: 'legacy',
    kind: 'recalculate',
    revision: 1,
  });
  assert.equal(result.status, 'failed');
  assert.match(result.error?.message ?? '', /protocol/i);
});

test('worker entry retains one FormulaEngine per session and restores its handler', () => {
  const sent: CalculationSessionResult[] = [];
  const previous = () => undefined;
  const scope = new TestScope(sent, previous);
  const uninstall = installBrowserCalculationWorkerEntry(scope);
  scope.onmessage?.({ data: openRequest() });
  scope.onmessage?.({ data: deltaRequest(2, 8) });
  assert.equal(sent[0]?.status, 'completed');
  assert.equal(sent[1]?.status, 'completed');
  assert.equal(sent[1]?.report?.results.some((entry) => entry.value === 24), true);
  uninstall();
  assert.equal(scope.onmessage, previous);
});

function openRequest(): CalculationSessionRequest {
  const source = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  source.setValue('A1', 2);
  source.setFormula('B1', '=A1*3');
  return {
    protocol: CALCULATION_DELTA_PROTOCOL,
    version: CALCULATION_DELTA_VERSION,
    kind: 'session.open',
    sessionId: 'session-test',
    taskId: 'task-open',
    revision: 1,
    generation: source.getCalculationGeneration(),
    bootstrap: source.exportCalculationBootstrap(),
  };
}

function deltaRequest(revision: number, value: number): CalculationSessionRequest {
  return {
    protocol: CALCULATION_DELTA_PROTOCOL,
    version: CALCULATION_DELTA_VERSION,
    kind: 'calculation.delta',
    sessionId: 'session-test',
    taskId: `task-${revision}`,
    revision,
    generation: 8,
    delta: { cells: [{ kind: 'set-value', address: { sheetId: 'Sheet1', row: 0, column: 0 }, value }] },
    roots: [{ sheetId: 'Sheet1', row: 0, column: 0 }],
  };
}

abstract class BaseCalculationWorker implements CalculationBrowserWorker {
  private readonly listeners = {
    message: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
    error: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
    messageerror: new Set<(event: { readonly data?: unknown; readonly message?: string }) => void>(),
  };

  abstract postMessage(message: unknown): void;

  terminate(): void { for (const listeners of Object.values(this.listeners)) listeners.clear(); }
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void { this.listeners[type].add(listener); }
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void { this.listeners[type].delete(listener); }
  protected emit(type: 'message' | 'error' | 'messageerror', event: { readonly data?: unknown; readonly message?: string }): void { for (const listener of this.listeners[type]) listener(event); }
}

class ImmediateCalculationWorker extends BaseCalculationWorker {
  calculationPosts = 0;
  readonly messages: unknown[] = [];
  private readonly sessions = new Map<string, FormulaEngine>();
  postMessage(message: unknown): void {
    if (!isCalculationMessage(message)) return;
    this.calculationPosts += 1;
    this.messages.push(structuredClone(message));
    queueMicrotask(() => this.emit('message', { data: consumeCalculationSession(this.sessions, message) }));
  }
}

class TestScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  constructor(private readonly sent: CalculationSessionResult[], previous: () => void) {
    this.onmessage = previous;
  }
  postMessage(message: CalculationSessionResult): void { this.sent.push(message); }
  terminate(): void {}
}

function isCalculationMessage(value: unknown): boolean {
  return typeof value === 'object' && value !== null && 'kind' in value && ((value as { kind?: unknown }).kind === 'session.open' || (value as { kind?: unknown }).kind === 'calculation.delta');
}
