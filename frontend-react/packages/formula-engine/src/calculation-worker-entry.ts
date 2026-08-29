import {
  assertCalculationSessionRequest,
  CALCULATION_DELTA_PROTOCOL,
  CALCULATION_DELTA_VERSION,
  type CalculationCancelRequest,
  type CalculationSessionRequest,
  type CalculationSessionResult,
} from './calculation-task-port';
import { FormulaEngine } from './formula-engine';
import { assertFormulaCalculationBootstrap } from './calculation-state';

/** Minimal Worker-like surface used by the entry point; no DOM/global is required. */
export interface CalculationWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: CalculationSessionResult): void;
}

export interface BrowserCalculationWorkerScope extends CalculationWorkerScope {}

/**
 * Consume one canonical delta message in a persistent session map. The map is
 * deliberately owned by the Worker entry; a new FormulaEngine is created only
 * for session.open and is reused for every later delta.
 */
export function consumeCalculationSession(
  sessions: Map<string, FormulaEngine>,
  payload: unknown,
): CalculationSessionResult {
  const request = payload as Partial<CalculationSessionRequest> | null;
  const sessionId = typeof request?.sessionId === 'string' && request.sessionId.length > 0 ? request.sessionId : 'invalid-session';
  const taskId = typeof request?.taskId === 'string' && request.taskId.length > 0 ? request.taskId : 'invalid-task';
  const revision = readRevision(request?.revision);
  const generation = readRevision(request?.generation);
  try {
    assertCalculationSessionRequest(request as CalculationSessionRequest);
    const canonical = request as CalculationSessionRequest;
    if (canonical.kind === 'session.open') {
      if (sessions.has(canonical.sessionId)) throw new Error(`Calculation session already exists: ${canonical.sessionId}`);
      assertFormulaCalculationBootstrap(canonical.bootstrap);
      const engine = FormulaEngine.fromCalculationBootstrap(canonical.bootstrap);
      sessions.set(canonical.sessionId, engine);
      const report = engine.executeCalculationDelta({
        protocol: CALCULATION_DELTA_PROTOCOL,
        version: CALCULATION_DELTA_VERSION,
        kind: 'calculation.delta',
        sessionId: canonical.sessionId,
        taskId: canonical.taskId,
        revision: canonical.revision,
        generation: canonical.generation,
        delta: {},
        forceRecalculate: true,
      });
      return {
        protocol: CALCULATION_DELTA_PROTOCOL,
        version: CALCULATION_DELTA_VERSION,
        kind: 'calculation.result',
        sessionId,
        taskId,
        revision,
        generation,
        status: 'completed',
        report,
      };
    }
    const engine = sessions.get(canonical.sessionId);
    if (!engine) throw new Error(`Calculation session is not open: ${canonical.sessionId}`);
    if (canonical.kind === 'calculation.cancel') {
      return cancelledResult(canonical.sessionId, taskId, revision, generation);
    }
    if (canonical.kind === 'session.close') {
      sessions.delete(canonical.sessionId);
      engine.disposeCalculationTasks();
      return {
        protocol: CALCULATION_DELTA_PROTOCOL,
        version: CALCULATION_DELTA_VERSION,
        kind: 'session.closed',
        sessionId,
        taskId,
        revision,
        generation,
        status: 'closed',
      };
    }
    const report = engine.executeCalculationDelta(canonical);
    return {
      protocol: CALCULATION_DELTA_PROTOCOL,
      version: CALCULATION_DELTA_VERSION,
      kind: 'calculation.result',
      sessionId,
      taskId,
      revision,
      generation,
      status: 'completed',
      report,
    };
  } catch (error) {
    return {
      protocol: CALCULATION_DELTA_PROTOCOL,
      version: CALCULATION_DELTA_VERSION,
      kind: 'calculation.failed',
      sessionId,
      taskId,
      revision,
      generation,
      status: 'failed',
      error: {
        code: 'CALCULATION_DELTA_FAILED',
        message: error instanceof Error ? error.message : 'Calculation delta failed',
        recovery: 'Discard the session and reopen it with a canonical bootstrap.',
      },
    };
  }
}

/** Install the production browser Worker entry with one persistent session map. */
export function installBrowserCalculationWorkerEntry(scope: BrowserCalculationWorkerScope): () => void {
  const previous = scope.onmessage;
  const sessions = new Map<string, FormulaEngine>();
  const openingTasks = new Map<string, string>();
  const cancelled = new Set<string>();
  scope.onmessage = (event) => {
    const request = readRequest(event.data);
    if (isValidCancellationRequest(request)) {
      cancelled.add(request.taskId);
      if (openingTasks.get(request.sessionId) === request.taskId) {
        sessions.get(request.sessionId)?.disposeCalculationTasks();
        sessions.delete(request.sessionId);
        openingTasks.delete(request.sessionId);
      }
      scope.postMessage(cancelledResult(request.sessionId, request.taskId, request.revision, request.generation));
      return;
    }
    if (request?.kind === 'session.open') openingTasks.set(request.sessionId!, request.taskId!);
    const result = consumeCalculationSession(sessions, event.data);
    if (request?.kind === 'session.open') openingTasks.delete(result.sessionId);
    if (cancelled.delete(result.taskId)) {
      scope.postMessage(cancelledResult(result.sessionId, result.taskId, result.revision, result.generation));
      return;
    }
    scope.postMessage(result);
  };
  return () => {
    scope.onmessage = previous;
    for (const engine of sessions.values()) engine.disposeCalculationTasks();
    sessions.clear();
    openingTasks.clear();
  };
}

function readRequest(value: unknown): Partial<CalculationSessionRequest> | null {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? value as Partial<CalculationSessionRequest>
    : null;
}

function readRevision(value: unknown): number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : 0;
}

function isValidCancellationRequest(
  request: Partial<CalculationSessionRequest> | null,
): request is CalculationCancelRequest {
  return request?.kind === 'calculation.cancel'
    && request.protocol === CALCULATION_DELTA_PROTOCOL
    && request.version === CALCULATION_DELTA_VERSION
    && typeof request.sessionId === 'string'
    && request.sessionId.trim().length > 0
    && typeof request.taskId === 'string'
    && request.taskId.trim().length > 0
    && readRevision(request.revision) === request.revision
    && readRevision(request.generation) === request.generation;
}

function cancelledResult(sessionId: string, taskId: string, revision: number, generation: number): CalculationSessionResult {
  return {
    protocol: CALCULATION_DELTA_PROTOCOL,
    version: CALCULATION_DELTA_VERSION,
    kind: 'calculation.result',
    sessionId,
    taskId,
    revision,
    generation,
    status: 'cancelled',
  };
}
