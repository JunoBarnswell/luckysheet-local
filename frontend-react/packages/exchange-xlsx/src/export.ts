import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBase64 } from './archive';
import { detectPackageFeatures } from './ooxml';
import { createCompatibilityReport } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import type { XlsxExportOptions, XlsxExportResult } from './types';

export interface XlsxExportRequest {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: XlsxExportOptions;
  /** Original package returned by importXlsx. Unsupported parts are preserved. */
  package?: import('./types').XlsxPackage;
}

/** 导出 XLSX 并生成 Compatibility Report */
export async function exportXlsx(request: XlsxExportRequest): Promise<XlsxExportResult> {
  const dateSystem = request.options.dateSystem ?? request.package?.dateSystem ?? '1900';
  const detectedFeatures = [...new Set([
    ...scanSnapshotFeatures(request.snapshot),
    ...(request.package ? detectPackageFeatures(request.package, request.snapshot) : []),
  ])];
  const preservedFeatures = request.package ? new Set(Object.keys(request.package.opaqueParts).flatMap((name) => {
    const lower = name.toLowerCase();
    if (lower.includes('/charts/')) return ['charts'];
    if (lower.includes('/pivot')) return ['pivot'];
    if (lower.includes('vba') || lower.endsWith('.bin')) return ['vba'];
    if (lower.includes('externalconnections') || lower.includes('connections.xml')) return ['external-connection'];
    if (lower.includes('/drawings/')) return ['images'];
    return [];
  })) : new Set<string>();
  const editableFeatures = new Set(['cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names', 'hyperlinks']);
  const unsupportedFeatures = detectedFeatures.filter((feature) => !editableFeatures.has(feature) && !preservedFeatures.has(feature));
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures,
    preservedFeatures,
    editableFeatures,
    unsupportedFeatures,
  });
  report.issues.push(...scanFormulaPreserveIssues(request.snapshot));

  return {
    taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    report,
    base64: exportSnapshotToXlsxBase64(request.snapshot, request.package, request.options),
    fileName: request.fileName.endsWith('.xlsx') ? request.fileName : `${request.fileName}.xlsx`,
  };
}
