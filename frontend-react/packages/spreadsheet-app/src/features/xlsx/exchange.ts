import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  CompatibilityLevel,
  NativePackageState,
  XlsxExportOptions,
  XlsxImportOptions,
  XlsxWorkerPort,
} from '@react-sheets/exchange-excel-ooxml';
import { excelCodecRegistry } from '@react-sheets/exchange-excel-ooxml';
import type { AssetStore } from '../persistence/asset-store';

export type { CompatibilityReport, CompatibilityLevel, XlsxExportOptions, XlsxImportOptions };

export const DEFAULT_XLSX_COMPATIBILITY: CompatibilityLevel = 'B';

export interface XlsxImportParams {
  fileName: string;
  buffer: ArrayBuffer;
  options?: Partial<XlsxImportOptions>;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
  assetStore?: AssetStore;
}

export interface XlsxExportParams {
  fileName?: string;
  options?: Partial<XlsxExportOptions>;
  nativePackage?: NativePackageState;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
  assetStore?: AssetStore;
}

export interface XlsxExchangeResult {
  report: CompatibilityReport;
  snapshot?: WorkbookSnapshot;
  buffer?: ArrayBuffer;
  fileName?: string;
  nativePackage?: NativePackageState;
}

export function buildXlsxImportOptions(overrides: Partial<XlsxImportOptions> = {}): XlsxImportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_XLSX_COMPATIBILITY,
    compatibilityMode: overrides.compatibilityMode,
    dateSystem: overrides.dateSystem,
    preserveMacros: overrides.preserveMacros ?? true,
  };
}

export function buildXlsxExportOptions(overrides: Partial<XlsxExportOptions> = {}): XlsxExportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_XLSX_COMPATIBILITY,
    dateSystem: overrides.dateSystem,
    includeCachedValues: overrides.includeCachedValues ?? true,
  };
}

export async function exchangeImportXlsx(params: XlsxImportParams): Promise<XlsxExchangeResult> {
  const request = {
    fileName: params.fileName,
    buffer: params.buffer,
    options: buildXlsxImportOptions(params.options),
    workerPort: params.workerPort,
    execution: params.execution,
    revision: params.revision,
  };
  const result = await excelCodecRegistry.import(request);
  if (result.snapshot && params.assetStore) await materializeImportedAssets(result.snapshot, result.nativePackage?.packageGraph.parts ?? {}, params.assetStore);
  else if (result.snapshot && result.report.issues.some((issue) => issue.feature === 'images') && collectAssetRefs(result.snapshot).length === 0) {
    throw new Error('ASSET_IMPORT_UNSUPPORTED: XLSX image drawing has no canonical AssetRef metadata');
  }
  return {
    report: result.report,
    snapshot: result.snapshot,
    nativePackage: result.nativePackage,
  };
}

async function materializeImportedAssets(snapshot: WorkbookSnapshot, parts: Record<string, Uint8Array>, assetStore: AssetStore): Promise<void> {
  const refs = collectAssetRefs(snapshot);
  for (const ref of refs) {
    const media = Object.entries(parts).find(([name]) => name.startsWith(`xl/media/${ref.assetId}.`))?.[1];
    if (!media) throw new Error(`ASSET_IMPORT_MISSING: ${ref.assetId}`);
    const stored = await assetStore.put({ content: new Blob([Uint8Array.from(media).buffer], { type: ref.mimeType }), mimeType: ref.mimeType, width: ref.width, height: ref.height });
    if (stored.assetId !== ref.assetId || stored.contentHash !== ref.contentHash) throw new Error(`ASSET_IMPORT_MISMATCH: ${ref.assetId}`);
  }
}

export async function exchangeExportXlsx(
  snapshot: WorkbookSnapshot,
  params: XlsxExportParams = {},
): Promise<XlsxExchangeResult> {
  const fileName = params.fileName ?? `${snapshot.name || 'workbook'}.xlsx`;
  const assetRefs = collectAssetRefs(snapshot);
  const assetBytes: Record<string, Uint8Array> = {};
  if (assetRefs.length) {
    if (!params.assetStore) throw new Error('ASSET_EXPORT_REQUIRED: AssetStore is required for workbook images');
    for (const ref of assetRefs) assetBytes[ref.assetId] = new Uint8Array(await (await params.assetStore.get(ref)).arrayBuffer());
  }
  const request = {
    snapshot,
    fileName,
    options: { ...buildXlsxExportOptions(params.options), ...(assetRefs.length ? { assetBytes } : {}) },
    ...(params.nativePackage ? { nativePackage: params.nativePackage } : {}),
    workerPort: params.workerPort,
    execution: params.execution,
    revision: params.revision,
  };
  const result = await excelCodecRegistry.export(request);
  return {
    report: result.report,
    buffer: result.buffer,
    fileName: result.fileName,
    nativePackage: result.nativePackage,
  };
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
