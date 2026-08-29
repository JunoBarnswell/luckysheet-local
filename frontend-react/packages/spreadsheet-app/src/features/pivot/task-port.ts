import { pivotSourceIndexTransferables } from './source-index';
import { PivotTaskEvaluator } from './task-worker-entry';
import {
  assertPivotTaskResult,
  createPivotTaskCancelRequest,
  PIVOT_TASK_PROTOCOL,
  PIVOT_TASK_VERSION,
  pivotTaskFailure,
  type PivotCalculateRequest,
  type PivotTaskErrorCode,
  type PivotTaskRequest,
  type PivotTaskResult,
} from './task-protocol';

export interface PivotBrowserWorker {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): void;
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
}

export interface PivotTaskPort {
  /**
   * Increments whenever the backing worker is replaced.  The session uses this
   * epoch to invalidate its source-registration cache; a restarted worker must
   * never receive a calculate request for a source it has not registered.
   */
  readonly sourceRegistrationEpoch: number;
  submit(request: Exclude<PivotTaskRequest, { kind: 'cancel' }>): Promise<PivotTaskResult>;
  cancel(taskId: string): void;
  dispose(): void;
}

interface PendingTask {
  request: Exclude<PivotTaskRequest, { kind: 'cancel' }>;
  resolve: (result: PivotTaskResult) => void;
  timeout: ReturnType<typeof setTimeout>;
}

/** Browser Worker transport. It owns transferred source buffers and never computes on the main thread. */
export class BrowserPivotTaskPort implements PivotTaskPort {
  private readonly pending = new Map<string, PendingTask>();
  private readonly workerFactory: (() => PivotBrowserWorker) | null;
  private worker: PivotBrowserWorker;
  private workerEpoch = 0;
  private disposed = false;

  constructor(worker: PivotBrowserWorker, timeoutMs = 30_000, workerFactory?: () => PivotBrowserWorker) {
    this.worker = worker;
    this.timeoutMs = timeoutMs;
    this.workerFactory = workerFactory ?? null;
    this.attachWorker(this.worker);
  }

  private readonly timeoutMs: number;

  get sourceRegistrationEpoch(): number {
    return this.workerEpoch;
  }

  submit(request: Exclude<PivotTaskRequest, { kind: 'cancel' }>): Promise<PivotTaskResult> {
    if (this.disposed) return Promise.resolve(failedResult(request, 'PIVOT_TASK_FAILED', 'Pivot worker has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'PIVOT_TASK_PROTOCOL_ERROR', `Pivot task already exists: ${request.taskId}`));
    return new Promise<PivotTaskResult>((resolve) => {
      const timeout = setTimeout(() => {
        const current = this.pending.get(request.taskId);
        if (!current) return;
        if (this.workerFactory) {
          this.interruptPendingTasks(request.taskId, 'PIVOT_TASK_TIMEOUT', `Pivot task exceeded ${String(this.timeoutMs)} ms`);
          return;
        }
        this.pending.delete(request.taskId);
        try { this.worker.postMessage(createPivotTaskCancelRequest(request.taskId, request.generation)); } catch { /* caller receives timeout below */ }
        resolve(failedResult(request, 'PIVOT_TASK_TIMEOUT', `Pivot task exceeded ${String(this.timeoutMs)} ms`));
      }, this.timeoutMs);
      this.pending.set(request.taskId, { request, resolve, timeout });
      try {
        const transfer = request.kind === 'source-register' ? pivotSourceIndexTransferables(request.source) : [];
        this.worker.postMessage(request, transfer);
      } catch (error) {
        this.pending.delete(request.taskId);
        clearTimeout(timeout);
        resolve(failedResult(request, 'PIVOT_TASK_FAILED', error instanceof Error ? error.message : 'Pivot worker postMessage failed'));
      }
    });
  }

  cancel(taskId: string): void {
    const pending = this.pending.get(taskId);
    if (!pending) return;
    if (this.workerFactory) {
      this.interruptPendingTasks(taskId, 'PIVOT_TASK_CANCELLED', 'Pivot task was cancelled');
      return;
    }
    this.pending.delete(taskId);
    clearTimeout(pending.timeout);
    try { this.worker.postMessage(createPivotTaskCancelRequest(taskId, pending.request.generation)); } catch { /* stale worker results remain ignored */ }
    pending.resolve({ protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation: pending.request.generation, status: 'cancelled' });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const taskId of [...this.pending.keys()]) this.cancel(taskId);
    this.detachWorker(this.worker);
    this.worker.terminate();
  }

  private attachWorker(worker: PivotBrowserWorker): void {
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleFailure);
    worker.addEventListener('messageerror', this.handleFailure);
  }

  private detachWorker(worker: PivotBrowserWorker): void {
    worker.removeEventListener('message', this.handleMessage);
    worker.removeEventListener('error', this.handleFailure);
    worker.removeEventListener('messageerror', this.handleFailure);
  }

  /**
   * Worker calculations are synchronous inside the worker event turn.  A
   * queued cancel message therefore cannot interrupt a long calculation.  In
   * production the port owns a factory and replaces the worker at the cancel
   * or timeout boundary, which is the only reliable interruption mechanism;
   * all requests queued behind the interrupted turn settle explicitly.
   */
  private interruptPendingTasks(primaryTaskId: string, primaryCode: PivotTaskErrorCode, primaryMessage: string): void {
    const pending = [...this.pending.values()];
    this.pending.clear();
    for (const entry of pending) clearTimeout(entry.timeout);
    let replacementError: unknown;
    try {
      this.replaceWorker();
    } catch (error) {
      replacementError = error;
    }
    for (const entry of pending) {
      if (replacementError) {
        entry.resolve(failedResult(entry.request, 'PIVOT_TASK_FAILED', replacementError instanceof Error ? replacementError.message : 'Pivot worker could not be restarted'));
        continue;
      }
      if (entry.request.taskId === primaryTaskId && primaryCode !== 'PIVOT_TASK_CANCELLED') {
        entry.resolve(failedResult(entry.request, primaryCode, primaryMessage));
      } else if (primaryCode === 'PIVOT_TASK_CANCELLED') {
        entry.resolve(cancelledResult(entry.request));
      } else {
        entry.resolve(failedResult(entry.request, 'PIVOT_TASK_CANCELLED', 'Pivot worker was restarted after another task timed out'));
      }
    }
  }

  private replaceWorker(): void {
    if (!this.workerFactory || this.disposed) return;
    this.detachWorker(this.worker);
    this.worker.terminate();
    this.worker = this.workerFactory();
    this.workerEpoch += 1;
    this.attachWorker(this.worker);
  }

  private readonly handleMessage = (event: { readonly data?: unknown }): void => {
    let result: PivotTaskResult;
    try {
      assertPivotTaskResult(event.data);
      result = event.data;
    } catch {
      this.failAll('PIVOT_TASK_PROTOCOL_ERROR', 'Pivot worker returned an invalid result');
      return;
    }
    const pending = this.pending.get(result.taskId);
    if (!pending) return;
    this.pending.delete(result.taskId);
    clearTimeout(pending.timeout);
    if (result.generation !== pending.request.generation) {
      pending.resolve(failedResult(pending.request, 'PIVOT_TASK_REVISION_MISMATCH', 'Pivot worker returned a mismatched generation'));
      return;
    }
    pending.resolve(result);
  };

  private readonly handleFailure = (event: { readonly message?: string }): void => {
    this.failAll('PIVOT_TASK_FAILED', event.message ?? 'Pivot worker failed');
    if (!this.disposed && this.workerFactory) {
      try {
        this.replaceWorker();
      } catch {
        // The pending callers already received the worker failure.  A future
        // submit will surface the replacement failure synchronously.
      }
    }
  };

  private failAll(code: PivotTaskErrorCode, message: string): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout);
      pending.resolve(failedResult(pending.request, code, message));
    }
    this.pending.clear();
  }
}

/** Test/non-browser host for the exact worker evaluator; production never selects it as a fallback. */
export class InlinePivotTaskPort implements PivotTaskPort {
  private readonly evaluator = new PivotTaskEvaluator();
  private readonly pending = new Map<string, { request: Exclude<PivotTaskRequest, { kind: 'cancel' }>; resolve: (result: PivotTaskResult) => void; timer: ReturnType<typeof setTimeout> }>();
  private disposed = false;

  readonly sourceRegistrationEpoch = 0;

  submit(request: Exclude<PivotTaskRequest, { kind: 'cancel' }>): Promise<PivotTaskResult> {
    if (this.disposed) return Promise.resolve(failedResult(request, 'PIVOT_TASK_FAILED', 'Inline Pivot task port has been disposed'));
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        if (!this.pending.delete(request.taskId)) return;
        resolve(this.evaluator.consume(request));
      }, 0);
      this.pending.set(request.taskId, { request, resolve, timer });
    });
  }

  cancel(taskId: string): void {
    const pending = this.pending.get(taskId);
    if (!pending) return;
    this.pending.delete(taskId);
    clearTimeout(pending.timer);
    this.evaluator.consume(createPivotTaskCancelRequest(taskId, pending.request.generation));
    pending.resolve({ protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation: pending.request.generation, status: 'cancelled' });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const taskId of [...this.pending.keys()]) this.cancel(taskId);
  }
}

export function createBrowserPivotTaskPort(): BrowserPivotTaskPort | null {
  if (typeof Worker === 'undefined') return null;
  const createWorker = (): PivotBrowserWorker => new Worker(new URL('./pivot-browser-worker.ts', import.meta.url), { type: 'module', name: 'pivot-runtime' }) as unknown as PivotBrowserWorker;
  return new BrowserPivotTaskPort(createWorker(), 30_000, createWorker);
}

function cancelledResult(request: Exclude<PivotTaskRequest, { kind: 'cancel' }>): PivotTaskResult {
  return {
    protocol: PIVOT_TASK_PROTOCOL,
    version: PIVOT_TASK_VERSION,
    taskId: request.taskId,
    generation: request.generation,
    status: 'cancelled',
  };
}

function failedResult(
  request: Exclude<PivotTaskRequest, { kind: 'cancel' }>,
  code: PivotTaskErrorCode,
  message: string,
): PivotTaskResult {
  if (request.kind === 'calculate') return pivotTaskFailure(request as PivotCalculateRequest, new Error(message), code);
  return {
    protocol: PIVOT_TASK_PROTOCOL,
    version: PIVOT_TASK_VERSION,
    taskId: request.taskId,
    generation: request.generation,
    status: 'failed',
    error: {
      code,
      message,
      pivotId: 'unbound',
      sourceIdentity: request.sourceIdentity,
      sourceRevision: request.sourceRevision,
      recovery: code === 'PIVOT_SOURCE_INVALID' || code === 'PIVOT_SOURCE_UNAVAILABLE' ? 'fix-source' : 'retry',
    },
  };
}
