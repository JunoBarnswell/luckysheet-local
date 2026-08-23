import {
  assertCalculationTaskRequest,
  isCalculationTaskCancellation,
  CALCULATION_TASK_PROTOCOL,
  CALCULATION_TASK_VERSION,
  type CalculationTaskCancellation,
  type CalculationTaskRequest,
  type CalculationTaskReport,
  type CalculationTaskResult,
} from './calculation-task-port';
import { FormulaEngine } from './formula-engine';
import { assertFormulaCalculationSnapshot } from './calculation-state';
import type { CalculationWorkerTaskRequest } from './calculation-browser-task-port';

/** Minimal Worker-like surface used by the entry point; no DOM/global is required. */
export interface CalculationWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: CalculationTaskResult): void;
}

/** Browser Worker scope that receives calculation snapshots and cancellations. */
export interface BrowserCalculationWorkerScope extends CalculationWorkerScope {
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

/**
 * Consume a browser task using only its structured-clone-safe calculation
 * snapshot. The FormulaEngine constructed here is isolated to this Worker.
 */
export function consumeBrowserCalculationTask(payload: unknown): CalculationTaskResult {
  const request = payload as Partial<CalculationWorkerTaskRequest> | null;
  const taskId = typeof request?.taskId === 'string' && request.taskId.length > 0 ? request.taskId : 'invalid-task';
  const revision = Number.isSafeInteger(request?.revision) && (request?.revision ?? -1) >= 0
    ? request!.revision!
    : 0;
  try {
    assertCalculationTaskRequest(request as CalculationTaskRequest);
    assertFormulaCalculationSnapshot(request?.snapshot);
    const engine = FormulaEngine.fromCalculationSnapshot(request.snapshot);
    return consumeCalculationTask(engine, request);
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

/**
 * Install the production browser Worker entry. Cancellation removes a queued
 * result before it can be posted; the host also settles and ignores cancelled
 * tasks immediately, so a late worker message cannot update the live engine.
 */
export function installBrowserCalculationWorkerEntry(scope: BrowserCalculationWorkerScope): () => void {
  const previous = scope.onmessage;
  const cancelled = new Set<string>();
  scope.onmessage = (event) => {
    if (isCalculationTaskCancellation(event.data)) {
      cancelled.add(event.data.taskId);
      return;
    }
    const taskId = readTaskId(event.data);
    const revision = readRevision(event.data);
    if (taskId && cancelled.delete(taskId)) {
      scope.postMessage(cancelledResult(taskId, revision));
      return;
    }
    const result = consumeBrowserCalculationTask(event.data);
    if (cancelled.delete(result.taskId)) {
      scope.postMessage(cancelledResult(result.taskId, result.revision));
      return;
    }
    scope.postMessage(result);
  };
  return () => {
    scope.onmessage = previous;
  };
}

/** Type-only helper for hosts which need to construct a report in a test port. */
export type CalculationWorkerReport = CalculationTaskReport;

function readTaskId(value: unknown): string | null {
  if (!isRecord(value) || typeof value.taskId !== 'string' || value.taskId.length === 0) return null;
  return value.taskId;
}

function readRevision(value: unknown): number {
  if (!isRecord(value) || typeof value.revision !== 'number') return 0;
  return Number.isSafeInteger(value.revision) && value.revision >= 0 ? value.revision : 0;
}

function cancelledResult(taskId: string, revision: number): CalculationTaskResult {
  return {
    protocol: CALCULATION_TASK_PROTOCOL,
    version: CALCULATION_TASK_VERSION,
    taskId,
    revision,
    status: 'cancelled',
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
