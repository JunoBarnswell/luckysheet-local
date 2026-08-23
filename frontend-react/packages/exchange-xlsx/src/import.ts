import { parseDateSystem } from './date-system';
import { createCompatibilityReport, refreshCompatibilitySummary } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import { detectPackageFeatures, loadXlsxPackage, parseLoadedXlsx } from './ooxml';
import { createXlsxSourceArtifact } from './source-artifact';
import type { XlsxImportOptions, XlsxImportResult } from './types';

export interface XlsxImportRequest {
  fileName: string;
  buffer: ArrayBuffer | Uint8Array;
  options: XlsxImportOptions;
}

/** 解析 XLSX 并生成 Compatibility Report */
export async function importXlsx(request: XlsxImportRequest): Promise<XlsxImportResult> {
  const bytes = request.buffer instanceof Uint8Array ? request.buffer.slice() : new Uint8Array(request.buffer);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const loaded = loadXlsxPackage(buffer, request.options.limits);
  const parsed = parseLoadedXlsx(loaded);
  const snapshot = parsed.snapshot;
  const dateSystem = request.options.dateSystem ?? parsed.package.dateSystem ?? parseDateSystem('');
  const snapshotFeatures = scanSnapshotFeatures(snapshot);
  const packageFeatures = detectPackageFeatures(parsed.package);
  const detectedFeatures = [...new Set([...packageFeatures, ...snapshotFeatures])];
  const editableFeatures = new Set(['cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names', 'hyperlinks', 'tables']);
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
  const completedReport = refreshCompatibilitySummary({ ...report, issues: [...report.issues, ...scanFormulaPreserveIssues(snapshot)] });
  const sourceArtifact = await createXlsxSourceArtifact({
    fileName: request.fileName,
    buffer,
    dateSystem,
    detectedFeatures,
  });

  return {
    payload: {
      name: request.fileName.replace(/\.xlsx$/i, ''),
      sheetCount: snapshot.sheets.length,
      dateSystem,
      compatibilityLevel: request.options.compatibilityTarget,
    },
    report: completedReport,
    snapshot,
    package: parsed.package,
    sourceArtifact,
    taskId: `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  };
}

export { parseDateSystem };
