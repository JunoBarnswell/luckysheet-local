import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBase64 } from './archive';
import { createCompatibilityReport } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import type { XlsxExportOptions, XlsxExportResult } from './types';

export interface XlsxExportRequest {
  snapshot: WorkbookSnapshotV1;
  fileName: string;
  options: XlsxExportOptions;
}

/** 导出 XLSX 并生成 Compatibility Report */
export async function exportXlsx(request: XlsxExportRequest): Promise<XlsxExportResult> {
  const dateSystem = request.options.dateSystem ?? '1900';
  const detectedFeatures = scanSnapshotFeatures(request.snapshot);
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures,
  });
  report.issues.push(...scanFormulaPreserveIssues(request.snapshot));

  return {
    taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    report,
    base64: exportSnapshotToXlsxBase64(request.snapshot),
    fileName: request.fileName.endsWith('.xlsx') ? request.fileName : `${request.fileName}.xlsx`,
  };
}
