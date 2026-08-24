import type { XlsxExportRequest } from './export';
import type { XlsxImportRequest } from './import';
import {
  assertXlsxWorkerResult,
  createXlsxCancelRequest,
  createXlsxExportRequest,
  createXlsxImportRequest,
  type XlsxWorkerExportRequest,
  type XlsxWorkerImportRequest,
  type XlsxWorkerRequest,
  type XlsxWorkerResult,
  XLSX_WORKER_PROTOCOL,
  XLSX_WORKER_VERSION,
} from './worker-protocol';
import type { XlsxExportResult, XlsxImportResult } from './types';

export interface XlsxBrowserWorker {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): void;
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
  removeEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void;
}

export interface XlsxWorkerPort {
  submit(request: XlsxWorkerImportRequest | XlsxWorkerExportRequest): Promise<XlsxWorkerResult>;
  cancel(taskId: string): void;
  dispose(): void;
}

interface PendingTask {
  request: XlsxWorkerImportRequest | XlsxWorkerExportRequest;
  resolve: (result: XlsxWorkerResult) => void;
  timeout: ReturnType<typeof setTimeout>;
}

/** Real browser Worker transport. It has no main-thread fallback. */
export class BrowserXlsxWorkerPort implements XlsxWorkerPort {
  private readonly pending = new Map<string, PendingTask>();
  private disposed = false;

  constructor(private readonly worker: XlsxBrowserWorker, private readonly timeoutMs = 120_000) {
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleFailure);
    worker.addEventListener('messageerror', this.handleFailure);
  }

  submit(request: XlsxWorkerImportRequest | XlsxWorkerExportRequest): Promise<XlsxWorkerResult> {
    if (this.disposed) return Promise.resolve(failedResult(request, 'XLSX_WORKER_DISPOSED', 'XLSX worker has been disposed'));
    if (this.pending.has(request.taskId)) return Promise.resolve(failedResult(request, 'XLSX_TASK_DUPLICATE', `XLSX task already exists: ${request.taskId}`));
    return new Promise<XlsxWorkerResult>((resolve) => {
      const timeout = setTimeout(() => {
        const current = this.pending.get(request.taskId);
        if (!current) return;
        this.pending.delete(request.taskId);
        try { this.worker.postMessage(createXlsxCancelRequest(request.taskId, request.revision)); } catch { /* failure is reported to the caller below */ }
        resolve(failedResult(request, 'XLSX_WORKER_TIMEOUT', `XLSX worker task exceeded ${this.timeoutMs} ms`));
      }, this.timeoutMs);
      this.pending.set(request.taskId, { request, resolve, timeout });
      try {
        const transfers: Transferable[] = request.kind === 'import' ? [request.payload.buffer] : [];
        this.worker.postMessage(request, transfers);
      } catch (error) {
        this.pending.delete(request.taskId);
        clearTimeout(timeout);
        resolve(failedResult(request, 'XLSX_WORKER_POST_FAILED', errorMessage(error)));
      }
    });
  }

  cancel(taskId: string): void {
    const pending = this.pending.get(taskId);
    if (!pending) return;
    this.pending.delete(taskId);
    clearTimeout(pending.timeout);
    try {
      this.worker.postMessage(createXlsxCancelRequest(taskId, pending.request.revision));
    } catch {
      // The local promise still settles as cancelled; the worker can no longer
      // mutate application state because the task is no longer pending.
    }
    pending.resolve({ protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, taskId, revision: pending.request.revision, status: 'cancelled' });
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
    let result: XlsxWorkerResult;
    try {
      assertXlsxWorkerResult(event.data);
      result = event.data;
    } catch {
      this.failAll('XLSX_WORKER_PROTOCOL_ERROR', 'XLSX worker returned an invalid result');
      return;
    }
    const pending = this.pending.get(result.taskId);
    if (!pending) return;
    if (result.status === 'progress') return;
    this.pending.delete(result.taskId);
    clearTimeout(pending.timeout);
    if (result.revision !== pending.request.revision) {
      pending.resolve(failedResult(pending.request, 'XLSX_WORKER_REVISION_MISMATCH', 'XLSX worker returned a mismatched revision'));
      return;
    }
    pending.resolve(result);
  };

  private readonly handleFailure = (event: { readonly message?: string }): void => {
    this.failAll('XLSX_WORKER_FAILED', event.message ?? 'XLSX worker failed');
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
export function createBrowserXlsxWorker(): XlsxBrowserWorker | null {
  if (typeof Worker === 'undefined') return null;
  return new Worker(new URL('./xlsx-browser-worker.ts', import.meta.url), { type: 'module', name: 'xlsx-exchange' }) as unknown as XlsxBrowserWorker;
}

export function createBrowserXlsxWorkerPort(): BrowserXlsxWorkerPort | null {
  const worker = createBrowserXlsxWorker();
  return worker ? new BrowserXlsxWorkerPort(worker) : null;
}

export async function importXlsxWithWorker(request: XlsxImportRequest, port?: XlsxWorkerPort, revision = 0): Promise<XlsxImportResult> {
  const active = port ?? createBrowserXlsxWorkerPort();
  if (!active) throw new Error('XLSX import requires a browser Worker; no Worker is available in this host');
  const owned = !port;
  const taskId = `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const bytes = request.buffer instanceof Uint8Array ? request.buffer.slice() : new Uint8Array(request.buffer);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  try {
    const result = await active.submit(createXlsxImportRequest(taskId, revision, { fileName: request.fileName, buffer, options: request.options }));
    return unwrapImportResult(result);
  } finally {
    if (owned) active.dispose();
  }
}

export async function exportXlsxWithWorker(request: XlsxExportRequest, port?: XlsxWorkerPort, revision = 0): Promise<XlsxExportResult> {
  const active = port ?? createBrowserXlsxWorkerPort();
  if (!active) throw new Error('XLSX export requires a browser Worker; no Worker is available in this host');
  const owned = !port;
  const taskId = `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  try {
    const result = await active.submit(createXlsxExportRequest(taskId, revision, {
      fileName: request.fileName,
      snapshot: request.snapshot,
      options: request.options,
      ...(request.nativePackage ? { nativePackage: request.nativePackage } : {}),
    }));
    return unwrapExportResult(result);
  } finally {
    if (owned) active.dispose();
  }
}

function unwrapImportResult(result: XlsxWorkerResult): XlsxImportResult {
  if (result.status === 'completed' && 'snapshot' in result.result) return result.result;
  throwWorkerResult(result);
}

function unwrapExportResult(result: XlsxWorkerResult): XlsxExportResult {
  if (result.status === 'completed' && 'buffer' in result.result) return result.result;
  throwWorkerResult(result);
}

function throwWorkerResult(result: XlsxWorkerResult): never {
  if (result.status === 'failed') throw new Error(`${result.error.code}: ${result.error.message}`);
  if (result.status === 'cancelled') throw new Error('XLSX worker task was cancelled');
  throw new Error('XLSX worker returned an invalid result');
}

function failedResult(request: XlsxWorkerRequest, code: string, message: string): XlsxWorkerResult {
  return {
    protocol: XLSX_WORKER_PROTOCOL,
    version: XLSX_WORKER_VERSION,
    taskId: request.taskId,
    revision: 'revision' in request && typeof request.revision === 'number' ? request.revision : 0,
    status: 'failed',
    error: { code, message },
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'XLSX worker failed';
}
