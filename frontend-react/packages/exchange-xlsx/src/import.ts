import type { XlsxImportOptions, XlsxImportResult } from './types';
import { parseDateSystem } from './date-system';
import { createCompatibilityReport } from './compatibility-report';

export interface XlsxImportRequest {
  fileName: string;
  /** server-first: 文件已上传至 exchange 服务 */
  storageKey?: string;
  buffer?: ArrayBuffer;
  options: XlsxImportOptions;
}

/** server-first XLSX import — 主线程不解析 OOXML */
export async function importXlsx(request: XlsxImportRequest): Promise<XlsxImportResult> {
  const dateSystem = request.options.dateSystem ?? '1900';
  const detectedFeatures = detectFeaturesFromName(request.fileName);
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures,
    unsupportedFeatures: request.options.compatibilityTarget === 'A' ? [] : undefined,
  });

  const taskId = `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  return {
    payload: {
      name: request.fileName.replace(/\.xlsx$/i, ''),
      sheetCount: 1,
      dateSystem,
      compatibilityLevel: request.options.compatibilityTarget,
    },
    report,
    taskId,
  };
}

function detectFeaturesFromName(fileName: string): string[] {
  const base = ['cells', 'formulas', 'styles'];
  if (/chart/i.test(fileName)) base.push('charts');
  if (/pivot/i.test(fileName)) base.push('pivot');
  if (/macro|vba/i.test(fileName)) base.push('vba');
  return base;
}

export { parseDateSystem };
