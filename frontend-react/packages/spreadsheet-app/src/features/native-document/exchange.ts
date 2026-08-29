import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  CompatibilityLevel,
  NativeDocumentArtifact,
  NativeDocumentExportOptions,
  NativeDocumentImportOptions,
  NativeDocumentWorkerPort,
} from '@react-sheets/exchange-excel-ooxml';
import { nativeDocumentCodecRegistry, verifyNativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
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

export type NativeDocumentTransactionState = 'idle' | 'imported' | 'exported' | 'failed';

/**
 * The sole application-facing native I/O transaction.  Import, Save, Save As
 * and export all pass through this serialized boundary; a failed asset or
 * codec step never publishes a half-updated artifact.
 */
export class NativeDocumentTransaction {
  private busy = false;
  private state: NativeDocumentTransactionState = 'idle';
  private currentArtifact: NativeDocumentArtifact | undefined;

  get status(): NativeDocumentTransactionState { return this.state; }
  get artifact(): NativeDocumentArtifact | undefined { return this.currentArtifact ? structuredClone(this.currentArtifact) : undefined; }

  /** Attach a persisted artifact only after its checksum/ownership graph is verified. */
  async attach(artifact: NativeDocumentArtifact): Promise<void> {
    await this.run(async () => {
      await verifyNativeDocumentArtifact(artifact);
      this.currentArtifact = structuredClone(artifact);
      this.state = 'imported';
    });
  }

  async import(params: NativeDocumentImportParams): Promise<NativeDocumentExchangeResult> {
    return this.run(async () => {
      const result = await exchangeImportDocument(params);
      this.currentArtifact = structuredClone(result.artifact);
      this.state = 'imported';
      return result;
    });
  }

  async export(snapshot: WorkbookSnapshot, params: NativeDocumentExportParams = {}): Promise<NativeDocumentExchangeResult> {
    const updatesBaseline = params.mode === 'save' || params.mode === undefined;
    return this.run(async () => {
      const artifact = params.artifact ?? this.currentArtifact;
      const result = await exchangeExportDocument(snapshot, { ...params, ...(artifact ? { artifact } : {}) });
      // Save As / Export produce a copy and never retarget the workbook's
      // shared source identity. Only an explicit Save establishes a baseline.
      if (updatesBaseline) {
        this.currentArtifact = structuredClone(result.artifact);
        this.state = 'exported';
      }
      return result;
    }, updatesBaseline);
  }

  async save(snapshot: WorkbookSnapshot, params: Omit<NativeDocumentExportParams, 'mode' | 'artifact'> & { fileName?: string } = {}): Promise<NativeDocumentExchangeResult> {
    return this.run(async () => {
      if (!this.currentArtifact) throw new Error('NATIVE_DOCUMENT_TRANSACTION_ARTIFACT_REQUIRED: Save requires a successful import or export transaction');
      const result = await exchangeSaveDocument(snapshot, this.currentArtifact, params);
      this.currentArtifact = structuredClone(result.artifact);
      this.state = 'exported';
      return result;
    });
  }

  private async run<T>(work: () => Promise<T>, invalidateBaselineOnFailure = true): Promise<T> {
    if (this.busy) throw new Error('NATIVE_DOCUMENT_TRANSACTION_BUSY: native document transactions are serialized');
    this.busy = true;
    try {
      return await work();
    } catch (error) {
      if (invalidateBaselineOnFailure) {
        this.state = 'failed';
        // A failed baseline-changing transaction cannot prove that the
        // previous source still matches the live workbook.
        this.currentArtifact = undefined;
      }
      throw error;
    } finally {
      this.busy = false;
    }
  }
}

export function createNativeDocumentTransaction(): NativeDocumentTransaction {
  return new NativeDocumentTransaction();
}

/**
 * Instance-scoped transaction ownership.  Catalog and session callers for a
 * unit resolve the same transaction here; there is intentionally no module
 * singleton and no second artifact cache.
 */
export class NativeDocumentTransactionRegistry {
  private readonly transactions = new Map<string, NativeDocumentTransaction>();

  get(unitId: string): NativeDocumentTransaction | undefined {
    assertTransactionUnitId(unitId);
    return this.transactions.get(unitId);
  }

  getOrCreate(unitId: string): NativeDocumentTransaction {
    assertTransactionUnitId(unitId);
    const current = this.transactions.get(unitId);
    if (current) return current;
    const created = createNativeDocumentTransaction();
    this.transactions.set(unitId, created);
    return created;
  }

  delete(unitId: string): boolean {
    assertTransactionUnitId(unitId);
    return this.transactions.delete(unitId);
  }

  clear(): void {
    this.transactions.clear();
  }
}

function assertTransactionUnitId(unitId: string): void {
  if (typeof unitId !== 'string' || !unitId.trim()) throw new Error('NATIVE_DOCUMENT_TRANSACTION_UNIT_REQUIRED: transaction registry requires a unit id');
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
  if (result.snapshot && params.assetStore) {
    const graph = result.artifact.nativeGraph.kind === 'opc' ? result.artifact.nativeGraph.package : undefined;
    await materializeImportedAssets(result.snapshot, graph?.parts ?? {}, graph?.assetPartById ?? {}, params.assetStore);
  }
  else if (result.snapshot && result.report.issues.some((issue) => issue.feature === 'images') && collectAssetRefs(result.snapshot).length === 0) {
    throw new Error('ASSET_IMPORT_UNSUPPORTED: native image drawing has no canonical AssetRef metadata');
  }
  return {
    report: result.report,
    snapshot: result.snapshot,
    artifact: result.artifact,
  };
}

async function materializeImportedAssets(snapshot: WorkbookSnapshot, parts: Record<string, Uint8Array>, assetPartById: Record<string, string>, assetStore: AssetStore): Promise<void> {
  const refs = collectAssetRefs(snapshot);
  for (const ref of refs) {
    const part = assetPartById[ref.assetId];
    const media = part ? parts[part] : undefined;
    if (!part || !media) throw new Error(`ASSET_IMPORT_MISSING: ${ref.assetId}`);
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
