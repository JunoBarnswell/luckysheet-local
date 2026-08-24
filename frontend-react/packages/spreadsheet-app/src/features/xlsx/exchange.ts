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
  nativePackage?: NativePackageState;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
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
  return {
    report: result.report,
    snapshot: result.snapshot,
    nativePackage: result.nativePackage,
  };
}

export async function exchangeExportXlsx(
  snapshot: WorkbookSnapshot,
  params: XlsxExportParams = {},
): Promise<XlsxExchangeResult> {
  const fileName = params.fileName ?? `${snapshot.name || 'workbook'}.xlsx`;
  const request = {
    snapshot,
    fileName,
    options: buildXlsxExportOptions(params.options),
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

export function summarizeCompatibilityReport(report: CompatibilityReport): string {
  const { editableFeatures, preservedOnly, unsupported } = report.summary;
  return `Import compatibility: ${editableFeatures} editable, ${preservedOnly} preserved, ${unsupported} unsupported`;
}
