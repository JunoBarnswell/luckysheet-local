import {
  assertCalculationSessionRequest,
  assertCalculationSessionResult,
  CALCULATION_DELTA_PROTOCOL,
  CALCULATION_DELTA_VERSION,
  type CalculationSessionPort,
  type CalculationSessionRequest,
  type CalculationSessionResult,
} from './calculation-task-port';

/** Minimal real Worker surface so browser tests can provide an in-memory worker. */
export interface CalculationBrowserWorker {
  postMessage(message: unknown): void;
  terminate(): void;
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
}

export type CalculationBrowserWorkerFactory = () => CalculationBrowserWorker;

interface PendingCalculation {
  readonly request: CalculationSessionRequest;
  readonly resolve: (result: CalculationSessionResult) => void;
}

/**
 * Persistent browser calculation session. The Worker owns one FormulaEngine
 * for the lifetime of this port. Only the initial session.open carries a
 * bootstrap; later messages are calculation deltas.
 */
export class BrowserCalculationSessionPort implements CalculationSessionPort {
  readonly protocol = CALCULATION_DELTA_PROTOCOL;
  readonly version = CALCULATION_DELTA_VERSION;

  private readonly pending = new Map<string, PendingCalculation>();
  private disposed = false;

  constructor(private readonly worker: CalculationBrowserWorker) {
    this.worker.addEventListener('message', this.handleMessage);
    this.worker.addEventListener('error', this.handleFailure);
    this.worker.addEventListener('messageerror', this.handleFailure);
  }

  submit(request: CalculationSessionRequest): Promise<CalculationSessionResult> {
    try {
      assertCalculationSessionRequest(request);
    } catch (error) {
      return Promise.resolve(failedResult(request, 'CALCULATION_DELTA_INVALID', errorMessage(error)));
    }
    if (this.disposed) return Promise.resolve(failedResult(request, 'CALCULATION_WORKER_DISPOSED', 'Calculation worker has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'CALCULATION_DELTA_DUPLICATE', `Calculation task already exists: ${request.taskId}`));

    return new Promise<CalculationSessionResult>((resolve) => {
      this.pending.set(request.taskId, { request, resolve });
      try {
        this.worker.postMessage(request);
      } catch (error) {
        this.pending.delete(request.taskId);
        resolve(failedResult(request, 'CALCULATION_WORKER_POST_FAILED', errorMessage(error)));
      }
    });
  }

  cancel(taskId: string): void {
    if (!taskId) return;
    const pending = this.pending.get(taskId);
    if (!pending) return;
    this.pending.delete(taskId);
    const request = pending.request;
    const cancellation = {
      protocol: CALCULATION_DELTA_PROTOCOL,
      version: CALCULATION_DELTA_VERSION,
      kind: 'calculation.cancel' as const,
      sessionId: request.sessionId,
      taskId,
      revision: request.revision,
      generation: request.generation,
    };
    try {
      this.worker.postMessage(cancellation);
    } catch {
      // The pending entry has already been removed, so a broken Worker cannot
      // apply the cancelled request to the host session.
    }
    pending.resolve({
      protocol: CALCULATION_DELTA_PROTOCOL,
      version: CALCULATION_DELTA_VERSION,
      kind: 'calculation.result',
      sessionId: request.sessionId,
      taskId,
      revision: request.revision,
      generation: request.generation,
      status: 'cancelled',
    });
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
    let result: CalculationSessionResult;
    try {
      assertCalculationSessionResult(event.data);
      result = event.data;
    } catch {
      this.failAll('CALCULATION_WORKER_PROTOCOL_ERROR', 'Calculation worker returned an invalid delta result');
      return;
    }
    const pending = this.pending.get(result.taskId);
    if (!pending) return;
    this.pending.delete(result.taskId);
    if (result.sessionId !== pending.request.sessionId || result.revision !== pending.request.revision || result.generation !== pending.request.generation) {
      pending.resolve(failedResult(pending.request, 'CALCULATION_WORKER_GENERATION_MISMATCH', 'Calculation worker returned a result for another session generation'));
      return;
    }
    pending.resolve(result);
  };

  private readonly handleFailure = (event: { readonly message?: string }): void => {
    this.failAll('CALCULATION_WORKER_FAILED', event.message ?? 'Calculation worker failed');
  };

  private failAll(code: string, message: string): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const pending of this.pending.values()) pending.resolve(failedResult(pending.request, code, message));
    this.pending.clear();
    this.worker.terminate();
  }
}

/** Creates the production browser Worker. No inline calculation fallback exists. */
export function createBrowserCalculationWorker(): CalculationBrowserWorker | null {
  if (typeof Worker === 'undefined') return null;
  return new Worker(new URL('./calculation-browser-worker.ts', import.meta.url), {
    type: 'module',
    name: 'formula-calculation',
  }) as unknown as CalculationBrowserWorker;
}

function failedResult(
  request: CalculationSessionRequest,
  code: string,
  message: string,
): CalculationSessionResult {
  return {
    protocol: CALCULATION_DELTA_PROTOCOL,
    version: CALCULATION_DELTA_VERSION,
    kind: 'calculation.failed',
    sessionId: request.sessionId,
    taskId: request.taskId,
    revision: request.revision,
    generation: request.generation,
    status: 'failed',
    error: { code, message },
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'Calculation worker failed';
}
