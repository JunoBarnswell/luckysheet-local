import { parseDateSystem } from './date-system';
import { createCompatibilityReport } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import { parseXlsxXmlToSnapshot, unzipXlsxBase64 } from './archive';
import type { XlsxImportOptions, XlsxImportResult } from './types';

export interface XlsxImportRequest {
  fileName: string;
  base64?: string;
  buffer?: ArrayBuffer;
  options: XlsxImportOptions;
}

function base64FromBuffer(buffer: ArrayBuffer): string {
  return Buffer.from(buffer).toString('base64');
}

/** 解析 XLSX 并生成 Compatibility Report */
export async function importXlsx(request: XlsxImportRequest): Promise<XlsxImportResult> {
  const base64 = request.base64 ?? (request.buffer ? base64FromBuffer(request.buffer) : undefined);
  if (!base64) throw new Error('XLSX import requires base64 or buffer payload');

  const files = unzipXlsxBase64(base64);
  if (!files['xl/workbook.xml']) {
    throw new Error('Not a valid XLSX package');
  }
  const snapshot = parseXlsxXmlToSnapshot(files);
  const dateSystem = request.options.dateSystem ?? parseDateSystem(files['xl/workbook.xml'] ?? '');
  const detectedFeatures = scanSnapshotFeatures(snapshot);
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures,
  });
  report.issues.push(...scanFormulaPreserveIssues(snapshot));

  return {
    payload: {
      name: request.fileName.replace(/\.xlsx$/i, ''),
      sheetCount: snapshot.sheets.length,
      dateSystem,
      compatibilityLevel: request.options.compatibilityTarget,
    },
    report,
    snapshot,
    taskId: `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  };
}

export { parseDateSystem };
