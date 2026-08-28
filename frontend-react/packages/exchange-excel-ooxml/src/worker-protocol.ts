import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type { NativeDocumentExportOptions, NativeDocumentExportResult, NativeDocumentImportOptions, NativeDocumentImportResult, NativeDocumentArtifact } from './types';

export const NATIVE_DOCUMENT_WORKER_PROTOCOL = 'react-sheets/native-document-io';
export const NATIVE_DOCUMENT_WORKER_VERSION = 1;

export interface NativeDocumentWorkerImportPayload {
  fileName: string;
  buffer: ArrayBuffer;
  options: NativeDocumentImportOptions;
}

export interface NativeDocumentWorkerExportPayload {
  fileName: string;
  snapshot: WorkbookSnapshot;
  options: NativeDocumentExportOptions;
  artifact?: NativeDocumentArtifact;
  mode?: 'save' | 'save-as' | 'export';
}

export interface NativeDocumentWorkerImportRequest {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  kind: 'import';
  taskId: string;
  revision: number;
  payload: NativeDocumentWorkerImportPayload;
}

export interface NativeDocumentWorkerExportRequest {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  kind: 'export';
  taskId: string;
  revision: number;
  payload: NativeDocumentWorkerExportPayload;
}

export interface NativeDocumentWorkerCancelRequest {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  kind: 'cancel';
  taskId: string;
  revision?: number;
}

export type NativeDocumentWorkerRequest = NativeDocumentWorkerImportRequest | NativeDocumentWorkerExportRequest | NativeDocumentWorkerCancelRequest;

export interface NativeDocumentWorkerError {
  code: string;
  message: string;
  format?: import('./types').NativeDocumentFormat;
  location?: string;
  recovery?: string;
}

export interface NativeDocumentWorkerCompletedResult {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'completed';
  result: NativeDocumentImportResult | NativeDocumentExportResult;
}

export interface NativeDocumentWorkerFailedResult {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'failed';
  error: NativeDocumentWorkerError;
}

export interface NativeDocumentWorkerCancelledResult {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'cancelled';
}

export interface NativeDocumentWorkerProgressResult {
  protocol: typeof NATIVE_DOCUMENT_WORKER_PROTOCOL;
  version: typeof NATIVE_DOCUMENT_WORKER_VERSION;
  taskId: string;
  revision: number;
  status: 'progress';
  stage: 'validate' | 'read' | 'parse' | 'serialize' | 'complete';
  percent: number;
}

export type NativeDocumentWorkerResult = NativeDocumentWorkerCompletedResult | NativeDocumentWorkerFailedResult | NativeDocumentWorkerCancelledResult | NativeDocumentWorkerProgressResult;

export function createNativeDocumentImportRequest(taskId: string, revision: number, payload: NativeDocumentWorkerImportPayload): NativeDocumentWorkerImportRequest {
  return { protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, kind: 'import', taskId, revision, payload };
}

export function createNativeDocumentExportRequest(taskId: string, revision: number, payload: NativeDocumentWorkerExportPayload): NativeDocumentWorkerExportRequest {
  return { protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, kind: 'export', taskId, revision, payload };
}

export function createNativeDocumentCancelRequest(taskId: string, revision?: number): NativeDocumentWorkerCancelRequest {
  return { protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, kind: 'cancel', taskId, ...(revision === undefined ? {} : { revision }) };
}

export function assertNativeDocumentWorkerRequest(value: unknown): asserts value is NativeDocumentWorkerRequest {
  if (!isRecord(value) || value.protocol !== NATIVE_DOCUMENT_WORKER_PROTOCOL || value.version !== NATIVE_DOCUMENT_WORKER_VERSION) {
    throw new Error('Invalid native document worker protocol or version');
  }
  if (typeof value.taskId !== 'string' || value.taskId.length === 0) throw new Error('Native document worker taskId is required');
  if (value.kind === 'cancel') {
    if (value.revision !== undefined && (typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0)) throw new Error('Native document worker cancellation revision is invalid');
    return;
  }
    if (typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0) throw new Error('Native document worker revision must be a non-negative integer');
  if (!isRecord(value.payload)) throw new Error('Native document worker payload is required');
  if (value.kind === 'import') {
    if (typeof value.payload.fileName !== 'string' || !isArrayBuffer(value.payload.buffer) || !isRecord(value.payload.options)) throw new Error('Invalid native document import worker payload');
    assertNativeDocumentOptions(value.payload.options, 'import');
    return;
  }
  if (value.kind === 'export') {
    if (typeof value.payload.fileName !== 'string' || !isRecord(value.payload.snapshot) || value.payload.snapshot.schema !== 'WorkbookSnapshot' || !isRecord(value.payload.options)) throw new Error('Invalid native document export worker payload');
    assertNativeDocumentOptions(value.payload.options, 'export');
    if (value.payload.mode !== undefined && !['save', 'save-as', 'export'].includes(value.payload.mode)) throw new Error('Invalid native document export mode');
    return;
  }
  throw new Error('Unknown native document worker request kind');
}

export function assertNativeDocumentWorkerResult(value: unknown): asserts value is NativeDocumentWorkerResult {
  if (!isRecord(value) || value.protocol !== NATIVE_DOCUMENT_WORKER_PROTOCOL || value.version !== NATIVE_DOCUMENT_WORKER_VERSION || typeof value.taskId !== 'string' || value.taskId.length === 0 || typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0) {
    throw new Error('Invalid native document worker result envelope');
  }
  if (value.status === 'completed' && isRecord(value.result)) return;
  if (value.status === 'cancelled') return;
  if (value.status === 'progress' && (value.stage === 'validate' || value.stage === 'read' || value.stage === 'parse' || value.stage === 'serialize' || value.stage === 'complete') && typeof value.percent === 'number' && Number.isFinite(value.percent) && value.percent >= 0 && value.percent <= 100) return;
  if (value.status === 'failed' && isRecord(value.error) && typeof value.error.code === 'string' && typeof value.error.message === 'string' && (value.error.location === undefined || typeof value.error.location === 'string') && (value.error.recovery === undefined || typeof value.error.recovery === 'string')) return;
  throw new Error('Invalid native document worker result payload');
}

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return value instanceof ArrayBuffer;
}

function assertNativeDocumentOptions(value: Record<string, any>, kind: 'import' | 'export'): void {
  if (value.compatibilityTarget !== 'A' && value.compatibilityTarget !== 'B' && value.compatibilityTarget !== 'C') throw new Error(`Invalid native document ${kind} compatibility target`);
  if (value.compatibilityMode !== undefined && !['strict', 'balanced', 'best-effort'].includes(value.compatibilityMode)) throw new Error(`Invalid native document ${kind} compatibility mode`);
  if (value.dateSystem !== undefined && value.dateSystem !== '1900' && value.dateSystem !== '1904') throw new Error(`Invalid native document ${kind} date system`);
  if (value.preserveMacros !== undefined && typeof value.preserveMacros !== 'boolean') throw new Error(`Invalid native document ${kind} preserveMacros flag`);
  if (value.limits !== undefined) {
    if (!isRecord(value.limits)) throw new Error(`Invalid native document ${kind} limits`);
    for (const [key, limit] of Object.entries(value.limits)) if (!['maxArchiveBytes', 'maxEntries', 'maxEntryBytes', 'maxUncompressedBytes', 'maxCompressionRatio', 'maxCfbStreams', 'maxStreamBytes', 'maxRecordCount', 'maxXmlDepth', 'maxXmlBytes', 'maxCells'].includes(key) || typeof limit !== 'number' || !Number.isSafeInteger(limit) || limit <= 0) throw new Error(`Invalid native document ${kind} limit: ${key}`);
  }
}

function isRecord(value: unknown): value is Record<string, any> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
