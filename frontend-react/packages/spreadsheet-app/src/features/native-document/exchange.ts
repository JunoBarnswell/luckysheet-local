import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  CompatibilityLevel,
  NativeDocumentArtifact,
  NativeDocumentExportOptions,
  NativeDocumentImportOptions,
  NativeDocumentWorkerPort,
} from '@react-sheets/exchange-excel-ooxml';
import { nativeDocumentCodecRegistry } from '@react-sheets/exchange-excel-ooxml';
import type { AssetStore } from '../persistence/asset-store';

export type { CompatibilityReport, CompatibilityLevel, NativeDocumentExportOptions, NativeDocumentImportOptions };

export const DEFAULT_NATIVE_COMPATIBILITY: CompatibilityLevel = 'B';

export interface NativeDocumentImportParams {
  fileName: string;
  buffer: ArrayBuffer;
  options?: Partial<NativeDocumentImportOptions>;
  workerPort?: NativeDocumentWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
  assetStore?: AssetStore;
}

export interface NativeDocumentExportParams {
  fileName?: string;
  options?: Partial<NativeDocumentExportOptions>;
  artifact?: NativeDocumentArtifact;
  workerPort?: NativeDocumentWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
  assetStore?: AssetStore;
  mode?: 'save' | 'save-as' | 'export';
}

export interface NativeDocumentSaveAsParams extends Omit<NativeDocumentExportParams, 'fileName'> {
  fileName: string;
}

export interface NativeDocumentExchangeResult {
  report: CompatibilityReport;
  snapshot?: WorkbookSnapshot;
  buffer?: ArrayBuffer;
  fileName?: string;
  artifact: NativeDocumentArtifact;
}

export function buildNativeDocumentImportOptions(overrides: Partial<NativeDocumentImportOptions> = {}): NativeDocumentImportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_NATIVE_COMPATIBILITY,
    compatibilityMode: overrides.compatibilityMode,
    dateSystem: overrides.dateSystem,
    preserveMacros: overrides.preserveMacros ?? true,
    limits: overrides.limits,
  };
}

export function buildNativeDocumentExportOptions(overrides: Partial<NativeDocumentExportOptions> = {}): NativeDocumentExportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_NATIVE_COMPATIBILITY,
    dateSystem: overrides.dateSystem,
    includeCachedValues: overrides.includeCachedValues ?? true,
    preserveMacros: overrides.preserveMacros ?? true,
    limits: overrides.limits,
  };
}

export async function exchangeImportDocument(params: NativeDocumentImportParams): Promise<NativeDocumentExchangeResult> {
  const request = {
    fileName: params.fileName,
    buffer: params.buffer,
    options: buildNativeDocumentImportOptions(params.options),
    workerPort: params.workerPort,
    execution: params.execution,
    revision: params.revision,
  };
  const result = await nativeDocumentCodecRegistry.import(request);
  if (result.snapshot && params.assetStore) await materializeImportedAssets(result.snapshot, result.artifact.nativeGraph.kind === 'opc' ? result.artifact.nativeGraph.package.parts : {}, params.assetStore);
  else if (result.snapshot && result.report.issues.some((issue) => issue.feature === 'images') && collectAssetRefs(result.snapshot).length === 0) {
    throw new Error('ASSET_IMPORT_UNSUPPORTED: native image drawing has no canonical AssetRef metadata');
  }
  return {
    report: result.report,
    snapshot: result.snapshot,
    artifact: result.artifact,
  };
}

async function materializeImportedAssets(snapshot: WorkbookSnapshot, parts: Record<string, Uint8Array>, assetStore: AssetStore): Promise<void> {
  const refs = collectAssetRefs(snapshot);
  for (const ref of refs) {
    const media = Object.entries(parts).find(([name]) => name.startsWith(`xl/media/${ref.assetId}.`) || name.startsWith(`xl/embeddings/${ref.assetId}.`))?.[1];
    if (!media) throw new Error(`ASSET_IMPORT_MISSING: ${ref.assetId}`);
    const stored = await assetStore.put({ content: new Blob([Uint8Array.from(media).buffer], { type: ref.mimeType }), mimeType: ref.mimeType, width: ref.width, height: ref.height });
    if (stored.assetId !== ref.assetId || stored.contentHash !== ref.contentHash) throw new Error(`ASSET_IMPORT_MISMATCH: ${ref.assetId}`);
  }
}

export async function exchangeExportDocument(
  snapshot: WorkbookSnapshot,
  params: NativeDocumentExportParams = {},
): Promise<NativeDocumentExchangeResult> {
  const fileName = params.fileName ?? params.artifact?.fileName ?? `${snapshot.name || 'workbook'}.ssjson`;
  const assetRefs = collectAssetRefs(snapshot);
  const assetBytes: Record<string, Uint8Array> = {};
  if (assetRefs.length) {
    if (!params.assetStore) throw new Error('ASSET_EXPORT_REQUIRED: AssetStore is required for workbook images');
    for (const ref of assetRefs) assetBytes[ref.assetId] = new Uint8Array(await (await params.assetStore.get(ref)).arrayBuffer());
  }
  const request = {
    snapshot,
    fileName,
    options: { ...buildNativeDocumentExportOptions(params.options), ...(assetRefs.length ? { assetBytes } : {}) },
    ...(params.artifact ? { artifact: params.artifact } : {}),
    mode: params.mode ?? 'export',
    workerPort: params.workerPort,
    execution: params.execution,
    revision: params.revision,
  };
  const result = await nativeDocumentCodecRegistry.export(request);
  return {
    report: result.report,
    buffer: result.buffer,
    fileName: result.fileName,
    artifact: result.artifact,
  };
}

/** Native Save keeps the current artifact identity and never changes format. */
export function exchangeSaveDocument(
  snapshot: WorkbookSnapshot,
  artifact: NativeDocumentArtifact | undefined,
  params: Omit<NativeDocumentExportParams, 'artifact' | 'fileName'> & { fileName?: string } = {},
): Promise<NativeDocumentExchangeResult> {
  const fileName = params.fileName ?? artifact?.fileName ?? `${snapshot.name || 'workbook'}.ssjson`;
  if (artifact && fileName !== artifact.fileName) return Promise.reject(new Error('NATIVE_DOCUMENT_SAVE_TARGET_MISMATCH: Save must keep the original native file name; use Save As for a new target.'));
  return exchangeExportDocument(snapshot, { ...params, fileName, artifact, mode: 'save' });
}

/** Explicit conversion boundary. Callers must provide a target file name. */
export function exchangeSaveAsDocument(snapshot: WorkbookSnapshot, params: NativeDocumentSaveAsParams): Promise<NativeDocumentExchangeResult> {
  if (!params.fileName.trim()) return Promise.reject(new Error('NATIVE_DOCUMENT_SAVE_AS_TARGET_REQUIRED: Save As requires an explicit target file name'));
  return exchangeExportDocument(snapshot, { ...params, mode: 'save-as' });
}

function collectAssetRefs(value: unknown): import('@react-sheets/core-model').AssetRef[] {
  const refs: import('@react-sheets/core-model').AssetRef[] = [];
  const visit = (entry: unknown): void => {
    if (entry && typeof entry === 'object' && !Array.isArray(entry)
      && (entry as { schema?: unknown }).schema === 'AssetRef') {
      refs.push(structuredClone(entry) as import('@react-sheets/core-model').AssetRef);
      return;
    }
    if (Array.isArray(entry)) for (const child of entry) visit(child);
    else if (entry && typeof entry === 'object') for (const child of Object.values(entry)) visit(child);
  };
  visit(value);
  return [...new Map(refs.map((ref) => [ref.assetId, ref])).values()];
}

export function summarizeCompatibilityReport(report: CompatibilityReport): string {
  const { editableFeatures, preservedOnly, unsupported } = report.summary;
  return `Import compatibility: ${editableFeatures} editable, ${preservedOnly} preserved, ${unsupported} unsupported`;
}
