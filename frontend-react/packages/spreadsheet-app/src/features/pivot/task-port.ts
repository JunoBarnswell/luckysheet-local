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
  private disposed = false;

  constructor(private readonly worker: PivotBrowserWorker, private readonly timeoutMs = 30_000) {
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleFailure);
    worker.addEventListener('messageerror', this.handleFailure);
  }

  submit(request: Exclude<PivotTaskRequest, { kind: 'cancel' }>): Promise<PivotTaskResult> {
    if (this.disposed) return Promise.resolve(failedResult(request, 'PIVOT_TASK_FAILED', 'Pivot worker has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'PIVOT_TASK_PROTOCOL_ERROR', `Pivot task already exists: ${request.taskId}`));
    return new Promise<PivotTaskResult>((resolve) => {
      const timeout = setTimeout(() => {
        const current = this.pending.get(request.taskId);
        if (!current) return;
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
    this.pending.delete(taskId);
    clearTimeout(pending.timeout);
    try { this.worker.postMessage(createPivotTaskCancelRequest(taskId, pending.request.generation)); } catch { /* stale worker results remain ignored */ }
    pending.resolve({ protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation: pending.request.generation, status: 'cancelled' });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const taskId of [...this.pending.keys()]) this.cancel(taskId);
    this.worker.removeEventListener('message', this.handleMessage);
    this.worker.removeEventListener('error', this.handleFailure);
    this.worker.removeEventListener('messageerror', this.handleFailure);
    this.worker.terminate();
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
  const worker = new Worker(new URL('./pivot-browser-worker.ts', import.meta.url), { type: 'module', name: 'pivot-runtime' }) as unknown as PivotBrowserWorker;
  return new BrowserPivotTaskPort(worker);
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
