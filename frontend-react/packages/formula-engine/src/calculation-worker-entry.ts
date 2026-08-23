import {
  assertCalculationTaskRequest,
  CALCULATION_TASK_PROTOCOL,
  CALCULATION_TASK_VERSION,
  type CalculationTaskRequest,
  type CalculationTaskReport,
  type CalculationTaskResult,
} from './calculation-task-port';
import type { FormulaEngine } from './formula-engine';

/** Minimal Worker-like surface used by the entry point; no DOM/global is required. */
export interface CalculationWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: CalculationTaskResult): void;
}

/**
 * Consume one structured-clone-safe task payload and return one structured
 * clone-safe result. This function is the independent worker entry contract;
 * a browser Worker, Node worker thread or a test harness can call it directly.
 */
export function consumeCalculationTask(engine: FormulaEngine, payload: unknown): CalculationTaskResult {
  const request = payload as Partial<CalculationTaskRequest> | null;
  const taskId = typeof request?.taskId === 'string' && request.taskId.length > 0 ? request.taskId : 'invalid-task';
  const revision = Number.isSafeInteger(request?.revision) && (request?.revision ?? -1) >= 0
    ? request!.revision!
    : 0;
  try {
    assertCalculationTaskRequest(request as CalculationTaskRequest);
    const report = engine.executeCalculationTask(request as CalculationTaskRequest);
    return {
      protocol: CALCULATION_TASK_PROTOCOL,
      version: CALCULATION_TASK_VERSION,
      taskId,
      revision,
      status: 'completed',
      report,
    };
  } catch (error) {
    return {
      protocol: CALCULATION_TASK_PROTOCOL,
      version: CALCULATION_TASK_VERSION,
      taskId,
      revision,
      status: 'failed',
      error: {
        code: 'CALCULATION_TASK_FAILED',
        message: error instanceof Error ? error.message : 'Calculation task failed',
      },
    };
  }
}

/** Install a direct `onmessage` worker entry and return an uninstall function. */
export function installCalculationWorkerEntry(
  engine: FormulaEngine,
  scope: CalculationWorkerScope,
): () => void {
  const previous = scope.onmessage;
  scope.onmessage = (event) => {
    scope.postMessage(consumeCalculationTask(engine, event.data));
  };
  return () => {
    scope.onmessage = previous;
  };
}

/** Type-only helper for hosts which need to construct a report in a test port. */
export type CalculationWorkerReport = CalculationTaskReport;
