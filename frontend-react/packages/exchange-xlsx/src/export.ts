import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { XlsxExportOptions, XlsxExportResult } from './types';
import { createCompatibilityReport } from './compatibility-report';

export interface XlsxExportRequest {
  snapshot: WorkbookSnapshotV1;
  fileName: string;
  options: XlsxExportOptions;
}

/** server-first XLSX export — 走 Task Center */
export async function exportXlsx(request: XlsxExportRequest): Promise<XlsxExportResult> {
  const dateSystem = request.options.dateSystem ?? '1900';
  const sheetFeatures = request.snapshot.sheets.flatMap((s) => {
    const features = ['cells'];
    if (Object.keys(s.cells).length > 0) features.push('formulas');
    if (s.charts.length > 0) features.push('charts');
    if (s.pivots.length > 0) features.push('pivot');
    return features;
  });

  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures: [...new Set(sheetFeatures)],
  });

  return {
    taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    report,
  };
}
