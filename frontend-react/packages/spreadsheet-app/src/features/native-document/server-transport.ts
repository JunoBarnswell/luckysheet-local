import type { WorkbookApiClient, WorkbookImportRequest, WorkbookSourceArtifactMetadata } from '@react-sheets/protocol';
import type {
  CompatibilityIssue,
  CompatibilityLevel,
  CompatibilityReport,
  DateSystem,
  NativeDocumentArtifact,
  NativeDocumentFormat,
  NativeDocumentImportOptions,
  NativeDocumentTransport,
} from '@react-sheets/exchange-excel-ooxml';

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

function buildReport(
  fileName: string,
  target: CompatibilityLevel,
  dateSystem: DateSystem,
  features: readonly NativeFeatureMetadata[],
): CompatibilityReport {
  const issues = features.map(compatibilityIssue).filter((issue): issue is CompatibilityIssue => Boolean(issue));
  return {
    schema: 'CompatibilityReport',
    fileName,
    importLevel: target,
    exportLevel: target,
    dateSystem,
    issues,
    summary: {
      editableFeatures: features.filter((feature) => feature.support === 'edit').length,
      preservedOnly: features.filter((feature) => feature.support === 'preserve' || feature.support === 'export-only').length,
      unsupported: features.filter((feature) => feature.support === 'unsupported').length,
    },
  };
}

function buildArtifact(
  metadata: WorkbookSourceArtifactMetadata,
  compatibilityTarget: CompatibilityLevel,
): NativeDocumentArtifact {
  const native = parseMetadata(metadata);
  const dateSystem = parseDateSystem(native.documentMetadata.dateSystem);
  const report = buildReport(metadata.fileName, compatibilityTarget, dateSystem, native.documentMetadata.features);
  return {
    schema: 'NativeDocumentArtifact',
    unitId: metadata.unitId,
    fileName: metadata.fileName,
    format: parseFormat(native.format),
    checksum: metadata.checksum,
    byteLength: metadata.byteLength,
    sourceRevision: metadata.revision,
    dateSystem,
    detectedFeatures: native.documentMetadata.features.map((feature) => feature.feature),
    codecRevision: native.codecRevision,
    compatibility: report,
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
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
}

/** Production native I/O adapter. Browser code transfers bytes and consumes server identities only. */
export class WorkbookApiNativeDocumentTransport implements NativeDocumentTransport {
  constructor(
    private readonly api: Pick<WorkbookApiClient, 'createWorkbookImport' | 'saveNativeDocumentArtifact' | 'getWorkbookSourceArtifact'>,
    private readonly importDestination: NativeDocumentImportDestination = {},
  ) {}

  async import(request: Parameters<NativeDocumentTransport['import']>[0]): ReturnType<NativeDocumentTransport['import']> {
    const input: WorkbookImportRequest = {
      content: request.content,
      fileName: request.fileName,
      name: this.importDestination.name,
      spaceId: this.importDestination.spaceId,
      folderId: this.importDestination.folderId,
    };
    const response = await this.api.createWorkbookImport(input);
    const artifact = buildArtifact(response.artifact, request.options.compatibilityTarget);
    if (artifact.unitId !== response.unitId || artifact.sourceRevision !== response.revision) {
      throw new Error('NATIVE_DOCUMENT_IMPORT_IDENTITY_MISMATCH');
    }
    return {
      unitId: response.unitId,
      manifest: response.manifest,
      report: artifact.compatibility,
      artifact,
    };
  }

  async export(request: Parameters<NativeDocumentTransport['export']>[0]): ReturnType<NativeDocumentTransport['export']> {
    const metadata = await this.api.saveNativeDocumentArtifact(request.unitId, {
      revision: request.revision,
      fileName: request.fileName,
      format: exportFormat(request.fileName),
    });
    const downloaded = await this.api.getWorkbookSourceArtifact(request.unitId);
    if (downloaded.metadata.revision !== request.revision || downloaded.metadata.checksum !== metadata.checksum
      || downloaded.artifact.size !== metadata.byteLength) {
      throw new Error('NATIVE_DOCUMENT_EXPORT_IDENTITY_MISMATCH');
    }
    const content = await downloaded.artifact.arrayBuffer();
    if (await checksum(content) !== metadata.checksum) throw new Error('NATIVE_DOCUMENT_EXPORT_CHECKSUM_MISMATCH');
    const artifact = buildArtifact(metadata, request.options.compatibilityTarget);
    return {
      unitId: request.unitId,
      revision: request.revision,
      content,
      fileName: metadata.fileName,
      report: artifact.compatibility,
      artifact,
    };
  }
}
