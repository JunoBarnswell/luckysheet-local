import { parseDateSystem } from './date-system';
import { createCompatibilityReport } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import { detectPackageFeatures, loadXlsxPackage, parseLoadedXlsx } from './ooxml';
import type { XlsxImportOptions, XlsxImportResult } from './types';

export interface XlsxImportRequest {
  fileName: string;
  base64?: string;
  buffer?: ArrayBuffer;
  options: XlsxImportOptions;
}

function base64FromBuffer(buffer: ArrayBuffer): string {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

/** 解析 XLSX 并生成 Compatibility Report */
export async function importXlsx(request: XlsxImportRequest): Promise<XlsxImportResult> {
  const base64 = request.base64 ?? (request.buffer ? base64FromBuffer(request.buffer) : undefined);
  if (!base64) throw new Error('XLSX import requires base64 or buffer payload');

  const loaded = loadXlsxPackage(base64, request.options.limits);
  const parsed = parseLoadedXlsx(loaded);
  const snapshot = parsed.snapshot;
  const dateSystem = request.options.dateSystem ?? parsed.package.dateSystem ?? parseDateSystem('');
  const snapshotFeatures = scanSnapshotFeatures(snapshot);
  const packageFeatures = detectPackageFeatures(parsed.package);
  const detectedFeatures = [...new Set([...packageFeatures, ...snapshotFeatures])];
  const editableFeatures = new Set(['cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names', 'hyperlinks', 'comments']);
  const preservedFeatures = new Set(packageFeatures.filter((feature) => !editableFeatures.has(feature)));
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures,
    preservedFeatures,
    editableFeatures,
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
    package: parsed.package,
    taskId: `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  };
}

export { parseDateSystem };
