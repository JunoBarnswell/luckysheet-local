import { exportXlsx } from './export';
import { importXlsx } from './import';
import {
  assertXlsxWorkerRequest,
  XLSX_WORKER_PROTOCOL,
  XLSX_WORKER_VERSION,
  type XlsxWorkerResult,
  type XlsxWorkerRequest,
} from './worker-protocol';

export interface XlsxWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: XlsxWorkerResult, transfer?: Transferable[]): void;
}

export async function consumeXlsxWorkerTask(payload: unknown): Promise<XlsxWorkerResult> {
  const taskId = readTaskId(payload);
  const revision = readRevision(payload);
  try {
    assertXlsxWorkerRequest(payload);
    if (payload.kind === 'cancel') {
      return cancelledResult(payload.taskId, revision);
    }
    const result = payload.kind === 'import'
      ? await importXlsx(payload.payload)
      : await exportXlsx({
        snapshot: payload.payload.snapshot,
        fileName: payload.payload.fileName,
        options: payload.payload.options,
        ...(payload.payload.nativePackage ? { nativePackage: payload.payload.nativePackage } : {}),
      });
    return {
      protocol: XLSX_WORKER_PROTOCOL,
      version: XLSX_WORKER_VERSION,
      taskId: payload.taskId,
      revision: payload.revision,
      status: 'completed',
      result,
    };
  } catch (error) {
    return {
      protocol: XLSX_WORKER_PROTOCOL,
      version: XLSX_WORKER_VERSION,
      taskId,
      revision,
      status: 'failed',
      error: {
        code: 'XLSX_WORKER_TASK_FAILED',
        message: error instanceof Error ? error.message : 'XLSX worker task failed',
      },
    };
  }
}

export function installXlsxWorkerEntry(scope: XlsxWorkerScope): () => void {
  const previous = scope.onmessage;
  const cancelled = new Set<string>();
  scope.onmessage = (event) => {
    const value = event.data;
    if (isCancel(value)) {
      cancelled.add(value.taskId);
      scope.postMessage(cancelledResult(value.taskId, readRevision(value)));
      return;
    }
    const taskId = readTaskId(value);
    const revision = readRevision(value);
    if (taskId !== 'invalid-task') {
      const stage = typeof value === 'object' && value !== null && (value as { kind?: unknown }).kind === 'export' ? 'serialize' : 'read';
      scope.postMessage(progressResult(taskId, revision, stage, 5));
    }
    void consumeXlsxWorkerTask(value).then((result) => {
      if (cancelled.delete(result.taskId)) {
        return;
      }
      const transfers: Transferable[] = [];
      if (result.status === 'completed') {
        const completed = result.result;
        if ('buffer' in completed) transfers.push(completed.buffer);
        else transfers.push(completed.nativePackage.sourceBytes);
      }
      scope.postMessage(result, transfers);
    });
  };
  return () => { scope.onmessage = previous; };
}

function progressResult(taskId: string, revision: number, stage: 'validate' | 'read' | 'parse' | 'serialize' | 'complete', percent: number): XlsxWorkerResult {
  return { protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, taskId, revision, status: 'progress', stage, percent };
}

function cancelledResult(taskId: string, revision: number): XlsxWorkerResult {
  return { protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, taskId, revision, status: 'cancelled' };
}

function isCancel(value: unknown): value is Extract<XlsxWorkerRequest, { kind: 'cancel' }> {
  return typeof value === 'object' && value !== null && (value as { kind?: unknown }).kind === 'cancel' && typeof (value as { taskId?: unknown }).taskId === 'string';
}

function readTaskId(value: unknown): string {
  return typeof value === 'object' && value !== null && typeof (value as { taskId?: unknown }).taskId === 'string'
    ? (value as { taskId: string }).taskId
    : 'invalid-task';
}

function readRevision(value: unknown): number {
  const revision = typeof value === 'object' && value !== null ? (value as { revision?: unknown }).revision : undefined;
  return typeof revision === 'number' && Number.isSafeInteger(revision) && revision >= 0 ? revision : 0;
}
