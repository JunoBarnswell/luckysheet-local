import {
  buildFacadePlanForBounds,
  checkFacadeExecution,
  parseFacadeScript,
  validateFacadePlan,
  type FacadeDslLimits,
  type FacadeExecutionControl,
  type FacadePlan,
  type FacadeSheetBounds,
} from './dsl';

export const AUTOMATION_WORKER_PROTOCOL = 'react-sheets.automation-worker';

export interface AutomationWorkerRequest {
  readonly protocol: typeof AUTOMATION_WORKER_PROTOCOL;
  readonly taskId: string;
  readonly kind: 'plan';
  readonly source: string;
  readonly bounds: FacadeSheetBounds;
  readonly limits: FacadeDslLimits;
  readonly maxOperations: number;
  readonly maxDurationMs: number;
}

export interface AutomationWorkerCancel {
  readonly protocol: typeof AUTOMATION_WORKER_PROTOCOL;
  readonly taskId: string;
  readonly kind: 'cancel';
}

export interface AutomationWorkerFailure {
  readonly code: string;
  readonly message: string;
}

export type AutomationWorkerResult =
  | {
    readonly protocol: typeof AUTOMATION_WORKER_PROTOCOL;
    readonly taskId: string;
    readonly status: 'completed';
    readonly plan: FacadePlan;
  }
  | {
    readonly protocol: typeof AUTOMATION_WORKER_PROTOCOL;
    readonly taskId: string;
    readonly status: 'cancelled';
  }
  | {
    readonly protocol: typeof AUTOMATION_WORKER_PROTOCOL;
    readonly taskId: string;
    readonly status: 'failed';
    readonly error: AutomationWorkerFailure;
  };

export interface AutomationWorkerSurface {
  postMessage(message: unknown): void;
  terminate(): void;
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
}

export type AutomationWorkerFactory = () => AutomationWorkerSurface;

/**
 * Browser-only production Worker creation. There is intentionally no
 * synchronous or main-thread fallback: callers must handle an unavailable
 * Worker as an explicit execution error.
 */
export function createAutomationWorker(): AutomationWorkerSurface {
  if (typeof Worker === 'undefined') {
    throw new Error('AUTOMATION_WORKER_UNAVAILABLE: browser Worker support is required');
  }
  return new Worker(new URL('./automation-worker-entry.ts', import.meta.url), {
    type: 'module',
    name: 'spreadsheet-automation',
  }) as unknown as AutomationWorkerSurface;
}

export function isAutomationWorkerRequest(value: unknown): value is AutomationWorkerRequest {
  if (!isRecord(value)) return false;
  return value.protocol === AUTOMATION_WORKER_PROTOCOL
    && typeof value.taskId === 'string'
    && value.taskId.length > 0
    && value.kind === 'plan'
    && typeof value.source === 'string'
    && isSheetBounds(value.bounds)
    && isDslLimits(value.limits)
    && isPositiveSafeInteger(value.maxOperations)
    && isPositiveSafeInteger(value.maxDurationMs);
}

export function isAutomationWorkerCancel(value: unknown): value is AutomationWorkerCancel {
  return isRecord(value)
    && value.protocol === AUTOMATION_WORKER_PROTOCOL
    && value.kind === 'cancel'
    && typeof value.taskId === 'string'
    && value.taskId.length > 0;
}

export function isAutomationWorkerResult(value: unknown): value is AutomationWorkerResult {
  if (!isRecord(value) || value.protocol !== AUTOMATION_WORKER_PROTOCOL || typeof value.taskId !== 'string') return false;
  if (value.status === 'cancelled') return true;
  if (value.status === 'failed') {
    return isRecord(value.error) && typeof value.error.code === 'string' && typeof value.error.message === 'string';
  }
  return value.status === 'completed' && isRecord(value.plan) && Array.isArray(value.plan.operations);
}

/** Worker entry logic, exported for a real Worker and deterministic tests. */
export function consumeAutomationWorkerRequest(payload: unknown): AutomationWorkerResult {
  const taskId = isRecord(payload) && typeof payload.taskId === 'string' ? payload.taskId : 'invalid-task';
  try {
    if (!isAutomationWorkerRequest(payload)) throw workerError('AUTOMATION_WORKER_PROTOCOL', 'Invalid automation Worker request');
    const startedAt = Date.now();
    const control: FacadeExecutionControl = { deadlineAt: startedAt + payload.maxDurationMs };
    checkFacadeExecution(control);
    const program = parseFacadeScript(payload.source, payload.limits, control);
    checkFacadeExecution(control);
    const plan = buildFacadePlanForBounds(payload.bounds, program, control);
    if (plan.operations.length > payload.maxOperations) {
      throw workerError('AUTOMATION_WORKER_QUOTA', `Automation plan exceeds ${payload.maxOperations} operations`);
    }
    validateFacadePlan(plan, payload.bounds);
    if (!isSerializable(plan)) throw workerError('AUTOMATION_WORKER_RESULT', 'Automation plan is not serializable');
    checkFacadeExecution(control);
    return { protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'completed', plan };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const code = message.includes('timed out') ? 'AUTOMATION_WORKER_TIMEOUT'
      : message.includes('cancelled') ? 'AUTOMATION_WORKER_CANCELLED'
        : message.includes('exceeds') ? 'AUTOMATION_WORKER_QUOTA'
          : 'AUTOMATION_WORKER_FAILED';
    return {
      protocol: AUTOMATION_WORKER_PROTOCOL,
      taskId,
      status: 'failed',
      error: { code, message },
    };
  }
}

/**
 * Main-thread client. A task owns one Worker lifetime. On timeout/cancel the
 * Worker is terminated so a stuck plan can never fall back to this thread.
 */
export class AutomationWorkerClient {
  private worker: AutomationWorkerSurface | null = null;
  private sequence = 0;

  constructor(private readonly factory: AutomationWorkerFactory = createAutomationWorker) {}

  submit(
    source: string,
    bounds: FacadeSheetBounds,
    sandboxPolicy: { readonly limits: FacadeDslLimits; readonly maxOperations: number; readonly maxDurationMs: number },
    signal?: AbortSignal,
  ): Promise<AutomationWorkerResult> {
    const taskId = `automation-${Date.now().toString(36)}-${(this.sequence += 1).toString(36)}`;
    let worker: AutomationWorkerSurface;
    try {
      worker = this.factory();
      this.worker = worker;
    } catch (error) {
      return Promise.resolve({
        protocol: AUTOMATION_WORKER_PROTOCOL,
        taskId,
        status: 'failed',
        error: { code: 'AUTOMATION_WORKER_UNAVAILABLE', message: error instanceof Error ? error.message : String(error) },
      });
    }

    const request: AutomationWorkerRequest = {
      protocol: AUTOMATION_WORKER_PROTOCOL,
      taskId,
      kind: 'plan',
      source,
      bounds,
      limits: sandboxPolicy.limits,
      maxOperations: sandboxPolicy.maxOperations,
      maxDurationMs: sandboxPolicy.maxDurationMs,
    };

    return new Promise<AutomationWorkerResult>((resolve) => {
      let settled = false;
      const finish = (result: AutomationWorkerResult): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        signal?.removeEventListener('abort', cancel);
        worker.removeEventListener('message', onMessage);
        worker.removeEventListener('error', onError);
        worker.removeEventListener('messageerror', onError);
        if (this.worker === worker) this.worker = null;
        resolve(result);
      };
      const cancel = (): void => {
        try { worker.postMessage({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, kind: 'cancel' } satisfies AutomationWorkerCancel); } catch { /* worker is terminated below */ }
        worker.terminate();
        finish({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'cancelled' });
      };
      const onMessage = (event: { readonly data?: unknown }): void => {
        if (!isAutomationWorkerResult(event.data) || event.data.taskId !== taskId) {
          finish({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'failed', error: { code: 'AUTOMATION_WORKER_PROTOCOL', message: 'Automation Worker returned an invalid result' } });
          return;
        }
        finish(event.data);
      };
      const onError = (event: { readonly message?: string }): void => {
        finish({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'failed', error: { code: 'AUTOMATION_WORKER_FAILED', message: event.message ?? 'Automation Worker failed' } });
      };
      const timer = setTimeout(() => {
        try { worker.terminate(); } finally {
          finish({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'failed', error: { code: 'AUTOMATION_WORKER_TIMEOUT', message: 'Automation Worker exceeded its wall-clock limit' } });
        }
      }, sandboxPolicy.maxDurationMs);
      worker.addEventListener('message', onMessage);
      worker.addEventListener('error', onError);
      worker.addEventListener('messageerror', onError);
      signal?.addEventListener('abort', cancel, { once: true });
      if (signal?.aborted) {
        cancel();
        return;
      }
      try {
        worker.postMessage(request);
      } catch (error) {
        finish({ protocol: AUTOMATION_WORKER_PROTOCOL, taskId, status: 'failed', error: { code: 'AUTOMATION_WORKER_POST', message: error instanceof Error ? error.message : String(error) } });
      }
    });
  }

  dispose(): void {
    this.worker?.terminate();
    this.worker = null;
  }
}

function workerError(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string };
  error.code = code;
  return error;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isSheetBounds(value: unknown): value is FacadeSheetBounds {
  if (!isRecord(value) || typeof value.sheetId !== 'string') return false;
  return isPositiveSafeInteger(value.rowCount) && isPositiveSafeInteger(value.columnCount);
}

function isDslLimits(value: unknown): value is FacadeDslLimits {
  if (!isRecord(value)) return false;
  return isPositiveSafeInteger(value.maxSourceLength)
    && isPositiveSafeInteger(value.maxStatements)
    && isPositiveSafeInteger(value.maxCells);
}

function isPositiveSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0;
}

function isSerializable(value: unknown, seen = new Set<object>()): boolean {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true;
  if (typeof value === 'number') return Number.isFinite(value);
  if (typeof value !== 'object' || seen.has(value)) return false;
  seen.add(value);
  if (Array.isArray(value)) return value.every((entry) => isSerializable(entry, seen));
  return Object.values(value as Record<string, unknown>).every((entry) => isSerializable(entry, seen));
}
