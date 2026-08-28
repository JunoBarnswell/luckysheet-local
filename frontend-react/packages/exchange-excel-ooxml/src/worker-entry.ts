import { nativeDocumentCodecRegistry } from './codec-registry';
import {
  assertNativeDocumentWorkerRequest,
  NATIVE_DOCUMENT_WORKER_PROTOCOL,
  NATIVE_DOCUMENT_WORKER_VERSION,
  type NativeDocumentWorkerResult,
  type NativeDocumentWorkerRequest,
} from './worker-protocol';
import { NativeDocumentError } from './native-document-error';

export interface NativeDocumentWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: NativeDocumentWorkerResult, transfer?: Transferable[]): void;
}

export async function consumeNativeDocumentWorkerTask(payload: unknown): Promise<NativeDocumentWorkerResult> {
  const taskId = readTaskId(payload);
  const revision = readRevision(payload);
  try {
    assertNativeDocumentWorkerRequest(payload);
    if (payload.kind === 'cancel') {
      return cancelledResult(payload.taskId, revision);
    }
    const result = payload.kind === 'import'
      ? await nativeDocumentCodecRegistry.import({
        fileName: payload.payload.fileName,
        buffer: payload.payload.buffer,
        options: payload.payload.options,
        execution: 'inline-test',
        revision: payload.revision,
      })
      : await nativeDocumentCodecRegistry.export({
        snapshot: payload.payload.snapshot,
        fileName: payload.payload.fileName,
        options: payload.payload.options,
        ...(payload.payload.mode ? { mode: payload.payload.mode } : {}),
        ...(payload.payload.artifact ? { artifact: payload.payload.artifact } : {}),
        execution: 'inline-test',
        revision: payload.revision,
      });
    return {
      protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL,
      version: NATIVE_DOCUMENT_WORKER_VERSION,
      taskId: payload.taskId,
      revision: payload.revision,
      status: 'completed',
      result,
    };
  } catch (error) {
    return {
      protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL,
      version: NATIVE_DOCUMENT_WORKER_VERSION,
      taskId,
      revision,
      status: 'failed',
      error: {
        code: error instanceof NativeDocumentError ? error.code : 'NATIVE_DOCUMENT_WORKER_TASK_FAILED',
        message: error instanceof NativeDocumentError ? error.message.replace(`${error.code}: `, '') : error instanceof Error ? error.message : 'Native document worker task failed',
        ...(error instanceof NativeDocumentError && error.format ? { format: error.format } : {}),
        ...(error instanceof NativeDocumentError && error.location ? { location: error.location } : {}),
        ...(error instanceof NativeDocumentError && error.recovery ? { recovery: error.recovery } : {}),
      },
    };
  }
}

export function installNativeDocumentWorkerEntry(scope: NativeDocumentWorkerScope): () => void {
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
    void consumeNativeDocumentWorkerTask(value).then((result) => {
      if (cancelled.delete(result.taskId)) {
        return;
      }
      const transfers: Transferable[] = [];
      if (result.status === 'completed') {
        const completed = result.result;
        if ('buffer' in completed) transfers.push(completed.buffer);
        else transfers.push(completed.artifact.sourceBytes);
      }
      scope.postMessage(result, transfers);
    });
  };
  return () => { scope.onmessage = previous; };
}

function progressResult(taskId: string, revision: number, stage: 'validate' | 'read' | 'parse' | 'serialize' | 'complete', percent: number): NativeDocumentWorkerResult {
  return { protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, taskId, revision, status: 'progress', stage, percent };
}

function cancelledResult(taskId: string, revision: number): NativeDocumentWorkerResult {
  return { protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, taskId, revision, status: 'cancelled' };
}

function isCancel(value: unknown): value is Extract<NativeDocumentWorkerRequest, { kind: 'cancel' }> {
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
