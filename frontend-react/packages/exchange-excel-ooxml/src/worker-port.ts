import type { NativeDocumentExportRequest } from './export';
import type { NativeDocumentImportRequest } from './import';
import {
  assertNativeDocumentWorkerResult,
  createNativeDocumentCancelRequest,
  createNativeDocumentExportRequest,
  createNativeDocumentImportRequest,
  type NativeDocumentWorkerExportRequest,
  type NativeDocumentWorkerImportRequest,
  type NativeDocumentWorkerRequest,
  type NativeDocumentWorkerResult,
  NATIVE_DOCUMENT_WORKER_PROTOCOL,
  NATIVE_DOCUMENT_WORKER_VERSION,
} from './worker-protocol';
import type { NativeDocumentExportResult, NativeDocumentImportResult } from './types';
import { NativeDocumentError } from './native-document-error';

export interface NativeDocumentBrowserWorker {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): void;
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
}

export interface NativeDocumentWorkerPort {
  submit(request: NativeDocumentWorkerImportRequest | NativeDocumentWorkerExportRequest): Promise<NativeDocumentWorkerResult>;
  cancel(taskId: string): void;
  dispose(): void;
}

interface PendingTask {
  request: NativeDocumentWorkerImportRequest | NativeDocumentWorkerExportRequest;
  resolve: (result: NativeDocumentWorkerResult) => void;
  timeout: ReturnType<typeof setTimeout>;
}

/** Real browser Worker transport. It has no main-thread fallback. */
export class BrowserNativeDocumentWorkerPort implements NativeDocumentWorkerPort {
  private readonly pending = new Map<string, PendingTask>();
  private disposed = false;

  constructor(private readonly worker: NativeDocumentBrowserWorker, private readonly timeoutMs = 120_000) {
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleFailure);
    worker.addEventListener('messageerror', this.handleFailure);
  }

  submit(request: NativeDocumentWorkerImportRequest | NativeDocumentWorkerExportRequest): Promise<NativeDocumentWorkerResult> {
    if (this.disposed) return Promise.resolve(failedResult(request, 'NATIVE_DOCUMENT_WORKER_DISPOSED', 'Native document worker has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'NATIVE_DOCUMENT_TASK_DUPLICATE', `Native document task already exists: ${request.taskId}`));
    return new Promise<NativeDocumentWorkerResult>((resolve) => {
      const timeout = setTimeout(() => {
        const current = this.pending.get(request.taskId);
        if (!current) return;
        this.pending.delete(request.taskId);
        try { this.worker.postMessage(createNativeDocumentCancelRequest(request.taskId, request.revision)); } catch { /* failure is reported to the caller below */ }
        resolve(failedResult(request, 'NATIVE_DOCUMENT_WORKER_TIMEOUT', `Native document worker task exceeded ${this.timeoutMs} ms`));
      }, this.timeoutMs);
      this.pending.set(request.taskId, { request, resolve, timeout });
      try {
        const transfers: Transferable[] = request.kind === 'import' ? [request.payload.buffer] : [];
        this.worker.postMessage(request, transfers);
      } catch (error) {
        this.pending.delete(request.taskId);
        clearTimeout(timeout);
        resolve(failedResult(request, 'NATIVE_DOCUMENT_WORKER_POST_FAILED', errorMessage(error)));
      }
    });
  }

  cancel(taskId: string): void {
    const pending = this.pending.get(taskId);
    if (!pending) return;
    this.pending.delete(taskId);
    clearTimeout(pending.timeout);
    try {
      this.worker.postMessage(createNativeDocumentCancelRequest(taskId, pending.request.revision));
    } catch {
      // The local promise still settles as cancelled; the worker can no longer
      // mutate application state because the task is no longer pending.
    }
    pending.resolve({ protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, taskId, revision: pending.request.revision, status: 'cancelled' });
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
    let result: NativeDocumentWorkerResult;
    try {
      assertNativeDocumentWorkerResult(event.data);
      result = event.data;
    } catch {
      this.failAll('NATIVE_DOCUMENT_WORKER_PROTOCOL_ERROR', 'Native document worker returned an invalid result');
      return;
    }
    const pending = this.pending.get(result.taskId);
    if (!pending) return;
    if (result.status === 'progress') return;
    this.pending.delete(result.taskId);
    clearTimeout(pending.timeout);
    if (result.revision !== pending.request.revision) {
      pending.resolve(failedResult(pending.request, 'NATIVE_DOCUMENT_WORKER_REVISION_MISMATCH', 'Native document worker returned a mismatched revision'));
      return;
    }
    pending.resolve(result);
  };

  private readonly handleFailure = (event: { readonly message?: string }): void => {
    this.failAll('NATIVE_DOCUMENT_WORKER_FAILED', event.message ?? 'Native document worker failed');
  };

  private failAll(code: string, message: string): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout);
      pending.resolve(failedResult(pending.request, code, message));
    }
    this.pending.clear();
  }
}

/** Vite turns this direct constructor into an isolated module worker. */
export function createBrowserNativeDocumentWorker(): NativeDocumentBrowserWorker | null {
  if (typeof Worker === 'undefined') return null;
  return new Worker(new URL('./native-document-worker.ts', import.meta.url), { type: 'module', name: 'native-document-io' }) as unknown as NativeDocumentBrowserWorker;
}

export function createBrowserNativeDocumentWorkerPort(): BrowserNativeDocumentWorkerPort | null {
  const worker = createBrowserNativeDocumentWorker();
  return worker ? new BrowserNativeDocumentWorkerPort(worker) : null;
}

export async function importNativeDocumentWithWorker(request: NativeDocumentImportRequest, port?: NativeDocumentWorkerPort, revision = 0): Promise<NativeDocumentImportResult> {
  const active = port ?? createBrowserNativeDocumentWorkerPort();
  if (!active) throw new Error('Native document import requires a browser Worker; no Worker is available');
  const owned = !port;
  const taskId = `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const bytes = request.buffer instanceof Uint8Array ? request.buffer.slice() : new Uint8Array(request.buffer);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  try {
    const result = await active.submit(createNativeDocumentImportRequest(taskId, revision, { fileName: request.fileName, buffer, options: request.options }));
    return unwrapImportResult(result);
  } finally {
    if (owned) active.dispose();
  }
}

export async function exportNativeDocumentWithWorker(request: NativeDocumentExportRequest, port?: NativeDocumentWorkerPort, revision = 0): Promise<NativeDocumentExportResult> {
  const active = port ?? createBrowserNativeDocumentWorkerPort();
  if (!active) throw new Error('Native document export requires a browser Worker; no Worker is available');
  const owned = !port;
  const taskId = `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  try {
    const result = await active.submit(createNativeDocumentExportRequest(taskId, revision, {
      fileName: request.fileName,
      snapshot: request.snapshot,
      options: request.options,
      ...(request.mode ? { mode: request.mode } : {}),
      ...(request.artifact ? { artifact: request.artifact } : {}),
    }));
    return unwrapExportResult(result);
  } finally {
    if (owned) active.dispose();
  }
}

function unwrapImportResult(result: NativeDocumentWorkerResult): NativeDocumentImportResult {
  if (result.status === 'completed' && 'snapshot' in result.result) return result.result;
  throwWorkerResult(result);
}

function unwrapExportResult(result: NativeDocumentWorkerResult): NativeDocumentExportResult {
  if (result.status === 'completed' && 'buffer' in result.result) return result.result;
  throwWorkerResult(result);
}

function throwWorkerResult(result: NativeDocumentWorkerResult): never {
  if (result.status === 'failed') throw new NativeDocumentError({ code: result.error.code, message: result.error.message, format: result.error.format, location: result.error.location, recovery: result.error.recovery });
  if (result.status === 'cancelled') throw new Error('NATIVE_DOCUMENT_WORKER_CANCELLED: Native document worker task was cancelled');
  throw new Error('NATIVE_DOCUMENT_WORKER_RESULT_INVALID: Native document worker returned an invalid result');
}

function failedResult(request: NativeDocumentWorkerRequest, code: string, message: string): NativeDocumentWorkerResult {
  return {
    protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL,
    version: NATIVE_DOCUMENT_WORKER_VERSION,
    taskId: request.taskId,
    revision: 'revision' in request && typeof request.revision === 'number' ? request.revision : 0,
    status: 'failed',
    error: { code, message },
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'NATIVE_DOCUMENT_WORKER_FAILED: Native document worker failed';
}
