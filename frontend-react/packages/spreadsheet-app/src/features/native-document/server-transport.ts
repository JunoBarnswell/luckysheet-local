import type { WorkbookImportResponse, WorkbookSourceArtifactMetadata } from '@react-sheets/protocol';
import type {
  CompatibilityIssue,
  CompatibilityLevel,
  CompatibilityReport,
  DateSystem,
  NativeDocumentArtifact,
  NativeDocumentFormat,
  NativeDocumentTransport,
} from '@react-sheets/exchange-excel-ooxml';

/** The server rejects chunks larger than this bound. Keep the client bound in sync. */
export const NATIVE_DOCUMENT_UPLOAD_CHUNK_BYTES = 8 * 1024 * 1024;

export type NativeDocumentImportTaskState = 'uploading' | 'importing' | 'completed' | 'failed' | 'cancelled';

export interface NativeDocumentImportTaskRequest {
  readonly fileName: string;
  readonly name?: string;
  readonly spaceId?: string;
  readonly folderId?: string;
  readonly byteLength: number;
  readonly sha256: string;
}

/** Wire response returned by the chunk task endpoints. */
export interface NativeDocumentImportTaskResponse {
  readonly taskId: string;
  readonly state: NativeDocumentImportTaskState;
  readonly uploadedBytes: number;
  readonly byteLength: number;
  readonly result?: WorkbookImportResponse | null;
  readonly errorCode?: string | null;
  readonly errorMessage?: string | null;
}

/**
 * Server-owned native document API. The task methods are deliberately the
 * only import entrypoint exposed here; multipart import is not a browser
 * runtime path. Artifact methods are used for the server export transaction.
 */
export interface NativeDocumentServerApi {
  createNativeDocumentTask(request: NativeDocumentImportTaskRequest): Promise<NativeDocumentImportTaskResponse>;
  uploadNativeDocumentTaskChunk(taskId: string, offset: number, bytes: ArrayBuffer): Promise<NativeDocumentImportTaskResponse>;
  commitNativeDocumentTask(taskId: string): Promise<NativeDocumentImportTaskResponse>;
  cancelNativeDocumentTask(taskId: string): Promise<NativeDocumentImportTaskResponse>;
  saveNativeDocumentArtifact(unitId: string, request: { revision: number; fileName: string; format: string }): Promise<WorkbookSourceArtifactMetadata>;
  getWorkbookSourceArtifact(unitId: string): Promise<{ artifact: Blob; metadata: WorkbookSourceArtifactMetadata }>;
}

interface NativeFeatureMetadata {
  readonly feature: string;
  readonly support: 'edit' | 'preserve' | 'export-only' | 'unsupported';
  readonly reason: string;
}

interface NativeArtifactMetadata {
  readonly revision: number;
  readonly checksum: string;
  readonly byteLength: number;
  readonly format: string;
  readonly codecRevision: number;
  readonly documentMetadata: {
    readonly dateSystem: string;
    readonly features: readonly NativeFeatureMetadata[];
  };
}

export interface NativeDocumentImportDestination {
  readonly name?: string;
  readonly spaceId?: string;
  readonly folderId?: string;
}

function requireRecord(value: unknown, message: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(message);
  return value as Record<string, unknown>;
}

function parseMetadata(metadata: WorkbookSourceArtifactMetadata): NativeArtifactMetadata {
  const native = requireRecord(metadata.nativeMetadata, 'NATIVE_DOCUMENT_ARTIFACT_METADATA_MISSING');
  const documentMetadata = requireRecord(native.documentMetadata, 'NATIVE_DOCUMENT_DOCUMENT_METADATA_MISSING');
  const features = documentMetadata.features;
  if (native.revision !== metadata.revision || native.checksum !== metadata.checksum
    || native.byteLength !== metadata.byteLength || typeof native.format !== 'string'
    || !Number.isSafeInteger(native.codecRevision) || Number(native.codecRevision) < 1
    || typeof documentMetadata.dateSystem !== 'string' || !Array.isArray(features)) {
    throw new Error('NATIVE_DOCUMENT_ARTIFACT_METADATA_INVALID');
  }
  for (const entry of features) {
    const feature = requireRecord(entry, 'NATIVE_DOCUMENT_FEATURE_METADATA_INVALID');
    if (typeof feature.feature !== 'string' || typeof feature.reason !== 'string'
      || !['edit', 'preserve', 'export-only', 'unsupported'].includes(String(feature.support))) {
      throw new Error('NATIVE_DOCUMENT_FEATURE_METADATA_INVALID');
    }
  }
  return native as unknown as NativeArtifactMetadata;
}

function parseFormat(value: string): NativeDocumentFormat {
  if (['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'].includes(value)) {
    return { family: 'ooxml', profile: 'transitional', variant: value as 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'xlam' };
  }
  throw new Error(`NATIVE_DOCUMENT_FORMAT_UNSUPPORTED: ${value}`);
}

function parseDateSystem(value: string): DateSystem {
  if (value === 'excel1900' || value === '1900') return '1900';
  if (value === 'excel1904' || value === '1904') return '1904';
  throw new Error(`NATIVE_DOCUMENT_DATE_SYSTEM_INVALID: ${value}`);
}

function compatibilityIssue(feature: NativeFeatureMetadata): CompatibilityIssue | null {
  if (feature.support === 'edit') return null;
  const unsupported = feature.support === 'unsupported';
  return {
    level: unsupported ? 'A' : 'B',
    severity: unsupported ? 'error' : 'warning',
    feature: feature.feature,
    message: feature.reason,
    preserved: feature.support === 'preserve' || feature.support === 'export-only',
    status: unsupported ? 'unsupported' : 'preserved-only',
    projection: unsupported ? 'unsupported' : 'preserved',
    reason: feature.reason,
  };
}

function buildReport(fileName: string, target: CompatibilityLevel, dateSystem: DateSystem, features: readonly NativeFeatureMetadata[]): CompatibilityReport {
  const issues = features.map(compatibilityIssue).filter((issue): issue is CompatibilityIssue => Boolean(issue));
  return {
    schema: 'CompatibilityReport', fileName, importLevel: target, exportLevel: target, dateSystem, issues,
    summary: {
      editableFeatures: features.filter((feature) => feature.support === 'edit').length,
      preservedOnly: features.filter((feature) => feature.support === 'preserve' || feature.support === 'export-only').length,
      unsupported: features.filter((feature) => feature.support === 'unsupported').length,
    },
  };
}

function buildArtifact(metadata: WorkbookSourceArtifactMetadata, compatibilityTarget: CompatibilityLevel): NativeDocumentArtifact {
  const native = parseMetadata(metadata);
  const dateSystem = parseDateSystem(native.documentMetadata.dateSystem);
  const report = buildReport(metadata.fileName, compatibilityTarget, dateSystem, native.documentMetadata.features);
  return {
    schema: 'NativeDocumentArtifact', unitId: metadata.unitId, fileName: metadata.fileName,
    format: parseFormat(native.format), checksum: metadata.checksum, byteLength: metadata.byteLength,
    sourceRevision: metadata.revision, dateSystem,
    detectedFeatures: native.documentMetadata.features.map((feature) => feature.feature),
    codecRevision: native.codecRevision, compatibility: report,
  };
}

function exportFormat(fileName: string): string {
  const extension = fileName.split('.').at(-1)?.toLowerCase();
  if (!extension || !['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'].includes(extension)) {
    throw new Error(`NATIVE_DOCUMENT_FORMAT_UNSUPPORTED: ${fileName}`);
  }
  return extension;
}

async function checksum(buffer: ArrayBuffer): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('NATIVE_DOCUMENT_CRYPTO_UNAVAILABLE: SHA-256 is required to verify native artifacts');
  const digest = await globalThis.crypto.subtle.digest('SHA-256', buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
}

function assertTaskResponse(value: unknown, expected: { taskId: string; byteLength: number; uploadedBytes?: number }): NativeDocumentImportTaskResponse {
  const task = requireRecord(value, 'NATIVE_DOCUMENT_IMPORT_TASK_INVALID');
  if ((expected.taskId && task.taskId !== expected.taskId)
    || !task.taskId || !['uploading', 'importing', 'completed', 'failed', 'cancelled'].includes(String(task.state))
    || !Number.isSafeInteger(task.uploadedBytes) || Number(task.uploadedBytes) < 0
    || !Number.isSafeInteger(task.byteLength) || Number(task.byteLength) !== expected.byteLength
    || Number(task.uploadedBytes) > expected.byteLength
    || (expected.uploadedBytes !== undefined && Number(task.uploadedBytes) !== expected.uploadedBytes)) {
    throw new Error('NATIVE_DOCUMENT_IMPORT_TASK_INVALID');
  }
  const result = task as unknown as NativeDocumentImportTaskResponse;
  if (result.state === 'failed') {
    throw new Error(`NATIVE_DOCUMENT_IMPORT_TASK_FAILED: ${result.errorCode ?? 'IMPORT_FAILED'}: ${result.errorMessage ?? 'native import task failed'}`);
  }
  if (result.state === 'cancelled') throw new Error('NATIVE_DOCUMENT_IMPORT_TASK_CANCELLED');
  return result;
}

function requireImportResult(task: NativeDocumentImportTaskResponse): WorkbookImportResponse {
  if (task.state !== 'completed' || !task.result) throw new Error('NATIVE_DOCUMENT_IMPORT_TASK_INCOMPLETE: commit did not publish an import result');
  const result = requireRecord(task.result, 'NATIVE_DOCUMENT_IMPORT_RESULT_INVALID');
  if (typeof result.unitId !== 'string' || !Number.isSafeInteger(result.revision)
    || !result.manifest || typeof result.manifest !== 'object' || !result.artifact || typeof result.artifact !== 'object') {
    throw new Error('NATIVE_DOCUMENT_IMPORT_RESULT_INVALID');
  }
  return task.result;
}

async function contentBuffer(content: Blob | ArrayBuffer): Promise<ArrayBuffer> {
  return typeof Blob !== 'undefined' && content instanceof Blob ? content.arrayBuffer() : content.slice(0);
}

/** Production native I/O adapter. Browser code only uploads/downloads opaque bytes. */
export class WorkbookApiNativeDocumentTransport implements NativeDocumentTransport {
  constructor(
    private readonly api: NativeDocumentServerApi,
    private readonly importDestination: NativeDocumentImportDestination = {},
  ) {}

  async import(request: Parameters<NativeDocumentTransport['import']>[0]): ReturnType<NativeDocumentTransport['import']> {
    const content = await contentBuffer(request.content);
    if (content.byteLength < 1) throw new Error('NATIVE_DOCUMENT_IMPORT_EMPTY: native documents must contain bytes');
    const sha256 = await checksum(content);
    const task = assertTaskResponse(await this.api.createNativeDocumentTask({
      fileName: request.fileName, name: this.importDestination.name, spaceId: this.importDestination.spaceId,
      folderId: this.importDestination.folderId, byteLength: content.byteLength, sha256,
    }), { taskId: '', byteLength: content.byteLength });
    const taskId = task.taskId;
    let offset = task.uploadedBytes;
    try {
      if (task.state !== 'uploading' || task.uploadedBytes !== 0) {
        throw new Error('NATIVE_DOCUMENT_IMPORT_TASK_INVALID: a new import task must start at offset zero');
      }
      while (offset < content.byteLength) {
        const end = Math.min(content.byteLength, offset + NATIVE_DOCUMENT_UPLOAD_CHUNK_BYTES);
        const next = assertTaskResponse(await this.api.uploadNativeDocumentTaskChunk(taskId, offset, content.slice(offset, end)), {
          taskId, byteLength: content.byteLength, uploadedBytes: end,
        });
        offset = next.uploadedBytes;
      }
      const committed = assertTaskResponse(await this.api.commitNativeDocumentTask(taskId), {
        taskId, byteLength: content.byteLength, uploadedBytes: content.byteLength,
      });
      const result = requireImportResult(committed);
      const artifact = buildArtifact(result.artifact, request.options.compatibilityTarget);
      if (artifact.unitId !== result.unitId || artifact.sourceRevision !== result.revision
        || result.manifest.unitId !== result.unitId || result.artifact.checksum !== sha256) {
        throw new Error('NATIVE_DOCUMENT_IMPORT_IDENTITY_MISMATCH');
      }
      return { unitId: result.unitId, manifest: result.manifest, report: artifact.compatibility, artifact };
    } catch (failure) {
      try { await this.api.cancelNativeDocumentTask(taskId); } catch { /* preserve the original task failure */ }
      throw failure;
    }
  }

  async export(request: Parameters<NativeDocumentTransport['export']>[0]): ReturnType<NativeDocumentTransport['export']> {
    const metadata = await this.api.saveNativeDocumentArtifact(request.unitId, {
      revision: request.revision, fileName: request.fileName, format: exportFormat(request.fileName),
    });
    const downloaded = await this.api.getWorkbookSourceArtifact(request.unitId);
    if (metadata.unitId !== request.unitId || metadata.revision !== request.revision
      || downloaded.metadata.unitId !== request.unitId || downloaded.metadata.fileName !== metadata.fileName
      || downloaded.metadata.revision !== request.revision || downloaded.metadata.checksum !== metadata.checksum
      || downloaded.artifact.size !== metadata.byteLength) {
      throw new Error('NATIVE_DOCUMENT_EXPORT_IDENTITY_MISMATCH');
    }
    const content = await downloaded.artifact.arrayBuffer();
    if (await checksum(content) !== metadata.checksum) throw new Error('NATIVE_DOCUMENT_EXPORT_CHECKSUM_MISMATCH');
    const artifact = buildArtifact(metadata, request.options.compatibilityTarget);
    return {
      unitId: request.unitId, revision: request.revision, content, fileName: metadata.fileName,
      report: artifact.compatibility, artifact,
    };
  }
}
