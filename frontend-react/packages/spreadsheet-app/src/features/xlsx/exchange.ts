import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  CompatibilityLevel,
  XlsxPackage,
  XlsxSourceArtifact,
  XlsxExportOptions,
  XlsxImportOptions,
  XlsxWorkerPort,
} from '@react-sheets/exchange-xlsx';

export type { CompatibilityReport, CompatibilityLevel, XlsxExportOptions, XlsxImportOptions };

export const DEFAULT_XLSX_COMPATIBILITY: CompatibilityLevel = 'B';

export interface XlsxImportParams {
  fileName: string;
  buffer: ArrayBuffer;
  options?: Partial<XlsxImportOptions>;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface XlsxExportParams {
  fileName?: string;
  options?: Partial<XlsxExportOptions>;
  package?: XlsxPackage;
  sourceArtifact?: XlsxSourceArtifact;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface XlsxExchangeResult {
  report: CompatibilityReport;
  snapshot?: WorkbookSnapshot;
  buffer?: ArrayBuffer;
  fileName?: string;
  sourceArtifact?: XlsxSourceArtifact;
  package?: XlsxPackage;
}

export function buildXlsxImportOptions(overrides: Partial<XlsxImportOptions> = {}): XlsxImportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_XLSX_COMPATIBILITY,
    dateSystem: overrides.dateSystem,
    preserveMacros: overrides.preserveMacros ?? false,
  };
}

export function buildXlsxExportOptions(overrides: Partial<XlsxExportOptions> = {}): XlsxExportOptions {
  return {
    compatibilityTarget: overrides.compatibilityTarget ?? DEFAULT_XLSX_COMPATIBILITY,
    dateSystem: overrides.dateSystem,
    includeCachedValues: overrides.includeCachedValues ?? true,
  };
}

async function loadExchangeXlsx() {
  return import('@react-sheets/exchange-xlsx');
}

export async function exchangeImportXlsx(params: XlsxImportParams): Promise<XlsxExchangeResult> {
  const exchange = await loadExchangeXlsx();
  const request = {
    fileName: params.fileName,
    buffer: params.buffer,
    options: buildXlsxImportOptions(params.options),
  };
  const result = params.execution === 'inline-test'
    ? await exchange.importXlsx(request)
    : await exchange.importXlsxWithWorker(request, params.workerPort, params.revision ?? 0);
  return {
    report: result.report,
    snapshot: result.snapshot,
    sourceArtifact: result.sourceArtifact,
    package: result.package,
  };
}

export async function exchangeExportXlsx(
  snapshot: WorkbookSnapshot,
  params: XlsxExportParams = {},
): Promise<XlsxExchangeResult> {
  const exchange = await loadExchangeXlsx();
  const fileName = params.fileName ?? `${snapshot.name || 'workbook'}.xlsx`;
  const request = {
    snapshot,
    fileName,
    options: buildXlsxExportOptions(params.options),
    ...(params.package ? { package: params.package } : {}),
    ...(params.sourceArtifact ? { sourceArtifact: params.sourceArtifact } : {}),
  };
  const result = params.execution === 'inline-test'
    ? await exchange.exportXlsx(request)
    : await exchange.exportXlsxWithWorker(request, params.workerPort, params.revision ?? 0);
  return {
    report: result.report,
    buffer: result.buffer,
    fileName: result.fileName,
    package: result.package,
  };
}

export function summarizeCompatibilityReport(report: CompatibilityReport): string {
  const { editableFeatures, preservedOnly, unsupported } = report.summary;
  return `Import compatibility: ${editableFeatures} editable, ${preservedOnly} preserved, ${unsupported} unsupported`;
}
