import {
  assertCalculationSessionRequest,
  CALCULATION_DELTA_PROTOCOL,
  CALCULATION_DELTA_VERSION,
  type CalculationSessionPort,
  type CalculationSessionRequest,
  type CalculationSessionResult,
} from './calculation-task-port';
import { consumeCalculationSession } from './calculation-worker-entry';
import { FormulaEngine } from './formula-engine';

/**
 * Explicit in-process calculation transport for Node/SSR and deterministic
 * harnesses. It uses the exact Worker entry protocol and keeps one engine per
 * session; production browser code must inject the Worker-backed port instead.
 */
export class LocalCalculationSessionPort implements CalculationSessionPort {
  readonly protocol = CALCULATION_DELTA_PROTOCOL;
  readonly version = CALCULATION_DELTA_VERSION;

  private readonly sessions = new Map<string, FormulaEngine>();
  private readonly pending = new Set<string>();
  private readonly cancelled = new Set<string>();
  private disposed = false;

  submit(request: CalculationSessionRequest): Promise<CalculationSessionResult> {
    try {
      assertCalculationSessionRequest(request);
    } catch (error) {
      return Promise.resolve(failedResult(request, 'CALCULATION_DELTA_INVALID', errorMessage(error)));
    }
    if (this.disposed) return Promise.resolve(failedResult(request, 'CALCULATION_SESSION_DISPOSED', 'Local calculation session has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'CALCULATION_DELTA_DUPLICATE', `Calculation task already exists: ${request.taskId}`));

    this.pending.add(request.taskId);
    return Promise.resolve().then(() => {
      this.pending.delete(request.taskId);
      if (this.cancelled.delete(request.taskId)) return cancelledResult(request);
      return consumeCalculationSession(this.sessions, request);
    });
  }

  cancel(taskId: string): void {
    if (!taskId || !this.pending.has(taskId)) return;
    this.cancelled.add(taskId);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.pending.clear();
    this.cancelled.clear();
    for (const engine of this.sessions.values()) engine.disposeCalculationTasks();
    this.sessions.clear();
  }
}

/** Explicit factory; no runtime path calls this implicitly. */
export function createLocalCalculationSessionPort(): LocalCalculationSessionPort {
  return new LocalCalculationSessionPort();
}

function failedResult(request: CalculationSessionRequest, code: string, message: string): CalculationSessionResult {
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

function cancelledResult(request: CalculationSessionRequest): CalculationSessionResult {
  return {
    protocol: CALCULATION_DELTA_PROTOCOL,
    version: CALCULATION_DELTA_VERSION,
    kind: 'calculation.result',
    sessionId: request.sessionId,
    taskId: request.taskId,
    revision: request.revision,
    generation: request.generation,
    status: 'cancelled',
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'Local calculation session failed';
}
