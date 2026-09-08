import type { KernelReplicaManifest, KernelReplicaPagePayload } from '@react-sheets/core-model';
import {
  asNativeDocumentError,
  assertNativeDocumentArtifact,
  cloneNativeDocumentArtifact,
  type CompatibilityReport,
  type CompatibilityLevel,
  type NativeDocumentArtifact,
  type NativeDocumentExportOptions,
  type NativeDocumentExportRequest,
  type NativeDocumentImportOptions,
  type NativeDocumentImportRequest,
  type NativeDocumentTransport,
} from '@react-sheets/exchange-excel-ooxml';

export type { CompatibilityReport, CompatibilityLevel, NativeDocumentExportOptions, NativeDocumentImportOptions, NativeDocumentTransport };

export const DEFAULT_NATIVE_COMPATIBILITY: CompatibilityLevel = 'B';

export interface NativeDocumentImportParams extends Omit<NativeDocumentImportRequest, 'options'> {
  options?: Partial<NativeDocumentImportOptions>;
  transport?: NativeDocumentTransport;
}

export interface NativeDocumentExportParams extends Omit<NativeDocumentExportRequest, 'options'> {
  options?: Partial<NativeDocumentExportOptions>;
  transport?: NativeDocumentTransport;
  mode?: 'save' | 'save-as' | 'export';
}

export interface NativeDocumentExchangeResult {
  manifest?: KernelReplicaManifest;
  pages?: KernelReplicaPagePayload[];
  content?: ArrayBuffer;
  fileName?: string;
  report: CompatibilityReport;
  artifact: NativeDocumentArtifact;
}

export type NativeDocumentTransactionState = 'idle' | 'imported' | 'exported' | 'failed';

/**
 * Serializes native document transport calls and owns only the verified
 * server artifact identity. Native bytes and document parsing never enter the
 * browser transaction.
 */
export class NativeDocumentTransaction {
  private busy = false;
  private state: NativeDocumentTransactionState = 'idle';
  private currentArtifact: NativeDocumentArtifact | undefined;

  constructor(private readonly defaultTransport?: NativeDocumentTransport, private readonly boundUnitId?: string) {}

  get status(): NativeDocumentTransactionState { return this.state; }
  get artifact(): NativeDocumentArtifact | undefined { return this.currentArtifact ? cloneNativeDocumentArtifact(this.currentArtifact) : undefined; }

  async attach(artifact: NativeDocumentArtifact): Promise<void> {
    await this.run(async () => {
      assertNativeDocumentArtifact(artifact);
      if (this.boundUnitId !== undefined && artifact.unitId !== this.boundUnitId) throw new Error('NATIVE_DOCUMENT_ATTACH_UNIT_MISMATCH: artifact belongs to another workbook');
      this.currentArtifact = cloneNativeDocumentArtifact(artifact);
      this.state = 'imported';
    });
  }

  async import(params: NativeDocumentImportParams): Promise<NativeDocumentExchangeResult> {
    return this.run(async () => {
      const result = await resolveTransport(params.transport ?? this.defaultTransport).import({
        fileName: params.fileName,
        content: params.content,
        options: buildNativeDocumentImportOptions(params.options),
      });
      assertNativeDocumentArtifact(result.artifact);
      if (result.unitId !== result.artifact.unitId || result.manifest.unitId !== result.unitId || (this.boundUnitId !== undefined && result.unitId !== this.boundUnitId)) {
        throw new Error('NATIVE_DOCUMENT_IMPORT_IDENTITY_MISMATCH: transport returned mismatched workbook identity');
      }
      this.currentArtifact = cloneNativeDocumentArtifact(result.artifact);
      this.state = 'imported';
      return { manifest: result.manifest, pages: result.pages, report: result.report, artifact: cloneNativeDocumentArtifact(result.artifact) };
    });
  }

  async export(params: NativeDocumentExportParams): Promise<NativeDocumentExchangeResult> {
    const updatesBaseline = params.mode === 'save' || params.mode === undefined;
    return this.run(async () => {
      if (this.boundUnitId !== undefined && params.unitId !== this.boundUnitId) throw new Error('NATIVE_DOCUMENT_EXPORT_UNIT_MISMATCH: transaction is bound to another workbook');
      const result = await resolveTransport(params.transport ?? this.defaultTransport).export({
        unitId: params.unitId,
        revision: params.revision,
        fileName: params.fileName,
        options: buildNativeDocumentExportOptions(params.options),
      });
      assertNativeDocumentArtifact(result.artifact);
      if (result.unitId !== params.unitId || result.revision !== params.revision || result.artifact.unitId !== params.unitId) {
        throw new Error('NATIVE_DOCUMENT_EXPORT_IDENTITY_MISMATCH: transport returned mismatched workbook identity');
      }
      if (updatesBaseline) {
        this.currentArtifact = cloneNativeDocumentArtifact(result.artifact);
        this.state = 'exported';
      }
      return { content: result.content, fileName: result.fileName, report: result.report, artifact: cloneNativeDocumentArtifact(result.artifact) };
    }, updatesBaseline);
  }

  private async run<T>(work: () => Promise<T>, invalidateBaselineOnFailure = true): Promise<T> {
    if (this.busy) throw new Error('NATIVE_DOCUMENT_TRANSACTION_BUSY: native document transactions are serialized');
    this.busy = true;
    try { return await work(); }
    catch (error) { if (invalidateBaselineOnFailure) { this.state = 'failed'; this.currentArtifact = undefined; } throw asNativeDocumentError(error); }
    finally { this.busy = false; }
  }
}

export function createNativeDocumentTransaction(transport?: NativeDocumentTransport, unitId?: string): NativeDocumentTransaction { return new NativeDocumentTransaction(transport, unitId); }

/** Unit-scoped transaction ownership; transport is supplied by the application host. */
export class NativeDocumentTransactionRegistry {
  private readonly transactions = new Map<string, NativeDocumentTransaction>();
  constructor(private readonly transport?: NativeDocumentTransport) {}
  get(unitId: string): NativeDocumentTransaction | undefined { assertUnitId(unitId); return this.transactions.get(unitId); }
  getOrCreate(unitId: string, transport: NativeDocumentTransport | undefined = this.transport): NativeDocumentTransaction {
    assertUnitId(unitId);
    const current = this.transactions.get(unitId);
    if (current) return current;
    const created = createNativeDocumentTransaction(transport, unitId);
    this.transactions.set(unitId, created);
    return created;
  }
  delete(unitId: string): boolean { assertUnitId(unitId); return this.transactions.delete(unitId); }
  clear(): void { this.transactions.clear(); }
}

function assertUnitId(unitId: string): void { if (typeof unitId !== 'string' || !unitId.trim()) throw new Error('NATIVE_DOCUMENT_TRANSACTION_UNIT_REQUIRED: transaction registry requires a unit id'); }

function resolveTransport(transport: NativeDocumentTransport | undefined): NativeDocumentTransport {
  if (!transport) throw new Error('NATIVE_DOCUMENT_TRANSPORT_UNAVAILABLE: a server native document transport is required');
  return transport;
}

export function buildNativeDocumentImportOptions(overrides: Partial<NativeDocumentImportOptions> = {}): NativeDocumentImportOptions {
  return { compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_NATIVE_COMPATIBILITY, compatibilityMode: overrides.compatibilityMode, preserveMacros: overrides.preserveMacros ?? true, dateSystem: overrides.dateSystem };
}

export function buildNativeDocumentExportOptions(overrides: Partial<NativeDocumentExportOptions> = {}): NativeDocumentExportOptions {
  return { compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_NATIVE_COMPATIBILITY, includeCachedValues: overrides.includeCachedValues ?? true, preserveMacros: overrides.preserveMacros ?? true, dateSystem: overrides.dateSystem };
}

export async function exchangeImportDocument(params: NativeDocumentImportParams): Promise<NativeDocumentExchangeResult> {
  return createNativeDocumentTransaction(params.transport).import(params);
}

export async function exchangeExportDocument(params: NativeDocumentExportParams): Promise<NativeDocumentExchangeResult> {
  return createNativeDocumentTransaction(params.transport).export(params);
}

export function summarizeCompatibilityReport(report: CompatibilityReport): string {
  const { editableFeatures, preservedOnly, unsupported } = report.summary;
  return `Import compatibility: ${editableFeatures} editable, ${preservedOnly} preserved, ${unsupported} unsupported`;
}
