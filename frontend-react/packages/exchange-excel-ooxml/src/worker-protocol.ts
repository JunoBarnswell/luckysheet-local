import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type { XlsxExportOptions, XlsxExportResult, XlsxImportOptions, XlsxImportResult, NativePackageState } from './types';

export const XLSX_WORKER_PROTOCOL = 'react-sheets/xlsx-exchange';
export const XLSX_WORKER_VERSION = 1;

export interface XlsxWorkerImportPayload {
  fileName: string;
  buffer: ArrayBuffer;
  options: XlsxImportOptions;
}

export interface XlsxWorkerExportPayload {
  fileName: string;
  snapshot: WorkbookSnapshot;
  options: XlsxExportOptions;
  nativePackage?: NativePackageState;
}

export interface XlsxWorkerImportRequest {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  kind: 'import';
  taskId: string;
  revision: number;
  payload: XlsxWorkerImportPayload;
}

export interface XlsxWorkerExportRequest {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  kind: 'export';
  taskId: string;
  revision: number;
  payload: XlsxWorkerExportPayload;
}

export interface XlsxWorkerCancelRequest {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  kind: 'cancel';
  taskId: string;
  revision?: number;
}

export type XlsxWorkerRequest = XlsxWorkerImportRequest | XlsxWorkerExportRequest | XlsxWorkerCancelRequest;

export interface XlsxWorkerError {
  code: string;
  message: string;
}

export interface XlsxWorkerCompletedResult {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'completed';
  result: XlsxImportResult | XlsxExportResult;
}

export interface XlsxWorkerFailedResult {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'failed';
  error: XlsxWorkerError;
}

export interface XlsxWorkerCancelledResult {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'cancelled';
}

export interface XlsxWorkerProgressResult {
  protocol: typeof XLSX_WORKER_PROTOCOL;
  version: typeof XLSX_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'progress';
  stage: 'validate' | 'read' | 'parse' | 'serialize' | 'complete';
  percent: number;
}

export type XlsxWorkerResult = XlsxWorkerCompletedResult | XlsxWorkerFailedResult | XlsxWorkerCancelledResult | XlsxWorkerProgressResult;

export function createXlsxImportRequest(taskId: string, revision: number, payload: XlsxWorkerImportPayload): XlsxWorkerImportRequest {
  return { protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, kind: 'import', taskId, revision, payload };
}

export function createXlsxExportRequest(taskId: string, revision: number, payload: XlsxWorkerExportPayload): XlsxWorkerExportRequest {
  return { protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, kind: 'export', taskId, revision, payload };
}

export function createXlsxCancelRequest(taskId: string, revision?: number): XlsxWorkerCancelRequest {
  return { protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, kind: 'cancel', taskId, ...(revision === undefined ? {} : { revision }) };
}

export function assertXlsxWorkerRequest(value: unknown): asserts value is XlsxWorkerRequest {
  if (!isRecord(value) || value.protocol !== XLSX_WORKER_PROTOCOL || value.version !== XLSX_WORKER_VERSION) {
    throw new Error('Invalid XLSX worker protocol or version');
  }
  if (typeof value.taskId !== 'string' || value.taskId.length === 0) throw new Error('XLSX worker taskId is required');
  if (value.kind === 'cancel') {
    if (value.revision !== undefined && (typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0)) throw new Error('XLSX worker cancellation revision is invalid');
    return;
  }
  if (typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0) throw new Error('XLSX worker revision must be a non-negative integer');
  if (!isRecord(value.payload)) throw new Error('XLSX worker payload is required');
  if (value.kind === 'import') {
    if (typeof value.payload.fileName !== 'string' || !isArrayBuffer(value.payload.buffer) || !isRecord(value.payload.options)) throw new Error('Invalid XLSX import worker payload');
    assertXlsxOptions(value.payload.options, 'import');
    return;
  }
  if (value.kind === 'export') {
    if (typeof value.payload.fileName !== 'string' || !isRecord(value.payload.snapshot) || value.payload.snapshot.schema !== 'WorkbookSnapshot' || !isRecord(value.payload.options)) throw new Error('Invalid XLSX export worker payload');
    assertXlsxOptions(value.payload.options, 'export');
    return;
  }
  throw new Error('Unknown XLSX worker request kind');
}

export function assertXlsxWorkerResult(value: unknown): asserts value is XlsxWorkerResult {
  if (!isRecord(value) || value.protocol !== XLSX_WORKER_PROTOCOL || value.version !== XLSX_WORKER_VERSION || typeof value.taskId !== 'string' || value.taskId.length === 0 || typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0) {
    throw new Error('Invalid XLSX worker result envelope');
  }
  if (value.status === 'completed' && isRecord(value.result)) return;
  if (value.status === 'cancelled') return;
  if (value.status === 'progress' && (value.stage === 'validate' || value.stage === 'read' || value.stage === 'parse' || value.stage === 'serialize' || value.stage === 'complete') && typeof value.percent === 'number' && Number.isFinite(value.percent) && value.percent >= 0 && value.percent <= 100) return;
  if (value.status === 'failed' && isRecord(value.error) && typeof value.error.code === 'string' && typeof value.error.message === 'string') return;
  throw new Error('Invalid XLSX worker result payload');
}

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return value instanceof ArrayBuffer;
}

function assertXlsxOptions(value: Record<string, any>, kind: 'import' | 'export'): void {
  if (value.compatibilityTarget !== 'A' && value.compatibilityTarget !== 'B' && value.compatibilityTarget !== 'C') throw new Error(`Invalid XLSX ${kind} compatibility target`);
  if (value.compatibilityMode !== undefined && !['strict', 'balanced', 'best-effort'].includes(value.compatibilityMode)) throw new Error(`Invalid XLSX ${kind} compatibility mode`);
  if (value.dateSystem !== undefined && value.dateSystem !== '1900' && value.dateSystem !== '1904') throw new Error(`Invalid XLSX ${kind} date system`);
  if (value.preserveMacros !== undefined && typeof value.preserveMacros !== 'boolean') throw new Error(`Invalid XLSX ${kind} preserveMacros flag`);
  if (kind === 'import' && value.limits !== undefined) {
    if (!isRecord(value.limits)) throw new Error('Invalid XLSX import limits');
    for (const [key, limit] of Object.entries(value.limits)) if (!['maxArchiveBytes', 'maxEntries', 'maxEntryBytes', 'maxUncompressedBytes', 'maxCompressionRatio'].includes(key) || typeof limit !== 'number' || !Number.isSafeInteger(limit) || limit <= 0) throw new Error(`Invalid XLSX import limit: ${key}`);
  }
}

function isRecord(value: unknown): value is Record<string, any> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
