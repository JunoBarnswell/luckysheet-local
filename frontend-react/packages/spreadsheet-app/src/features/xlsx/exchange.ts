import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  CompatibilityLevel,
  XlsxExportOptions,
  XlsxImportOptions,
} from '@react-sheets/exchange-xlsx';

export type { CompatibilityReport, CompatibilityLevel, XlsxExportOptions, XlsxImportOptions };

export const DEFAULT_XLSX_COMPATIBILITY: CompatibilityLevel = 'B';

export interface XlsxImportParams {
  fileName: string;
  base64: string;
  options?: Partial<XlsxImportOptions>;
}

export interface XlsxExportParams {
  fileName?: string;
  options?: Partial<XlsxExportOptions>;
}

export interface XlsxExchangeResult {
  report: CompatibilityReport;
  snapshot?: WorkbookSnapshot;
  base64?: string;
  fileName?: string;
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
  const { importXlsx } = await loadExchangeXlsx();
  const result = await importXlsx({
    fileName: params.fileName,
    base64: params.base64,
    options: buildXlsxImportOptions(params.options),
  });
  return {
    report: result.report,
    snapshot: result.snapshot,
  };
}

export async function exchangeExportXlsx(
  snapshot: WorkbookSnapshot,
  params: XlsxExportParams = {},
): Promise<XlsxExchangeResult> {
  const { exportXlsx } = await loadExchangeXlsx();
  const fileName = params.fileName ?? `${snapshot.name || 'workbook'}.xlsx`;
  const result = await exportXlsx({
    snapshot,
    fileName,
    options: buildXlsxExportOptions(params.options),
  });
  return {
    report: result.report,
    base64: result.base64,
    fileName: result.fileName,
  };
}

export function summarizeCompatibilityReport(report: CompatibilityReport): string {
  const { editableFeatures, preservedOnly, unsupported } = report.summary;
  return `Import compatibility: ${editableFeatures} editable, ${preservedOnly} preserved, ${unsupported} unsupported`;
}
