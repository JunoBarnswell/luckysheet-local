import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBuffer, loadXlsxPackage } from './archive';
import { detectPackageFeatures } from './ooxml';
import { nativePivotFeatureStatus } from './native-pivot';
import { createCompatibilityReport, refreshCompatibilitySummary } from './compatibility-report';
import { verifyXlsxSourceArtifact } from './source-artifact';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import type { XlsxExportOptions, XlsxExportResult, XlsxSourceArtifact } from './types';

export interface XlsxExportRequest {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: XlsxExportOptions;
  /** Original package returned by importXlsx. Unsupported parts are preserved. */
  package?: import('./types').XlsxPackage;
  /** Original bytes are sufficient to reconstruct a package when `package` is not retained. */
  sourceArtifact?: XlsxSourceArtifact;
}

/** 导出 XLSX 并生成 Compatibility Report */
export async function exportXlsx(request: XlsxExportRequest): Promise<XlsxExportResult> {
  if (request.sourceArtifact) await verifyXlsxSourceArtifact(request.sourceArtifact);
  const sourcePackage = request.package ?? (request.sourceArtifact ? loadXlsxPackage(request.sourceArtifact.buffer).package : undefined);
  const dateSystem = request.options.dateSystem ?? sourcePackage?.dateSystem ?? request.sourceArtifact?.dateSystem ?? '1900';
  const buffer = exportSnapshotToXlsxBuffer(request.snapshot, sourcePackage, request.options);
  // Report the package that was actually emitted. This prevents a deleted
  // native Pivot/Slicer/Timeline from being reported as preserved merely
  // because its source package contained the old opaque part.
  const emittedPackage = loadXlsxPackage(buffer).package;
  const snapshotFeatures = scanSnapshotFeatures(request.snapshot);
  const packageFeatures = detectPackageFeatures(emittedPackage);
  const detectedFeatures = [...new Set([...packageFeatures, ...snapshotFeatures])];
  const nativeStatus = nativePivotFeatureStatus(request.snapshot, emittedPackage.nativePivotGraph);
  const preservedFeatures = sourcePackage ? new Set(Object.keys(sourcePackage.opaqueParts).flatMap((name) => {
    const lower = name.toLowerCase();
    if (lower.includes('/charts/')) return ['charts'];
    if (lower.includes('/pivot')) return ['pivot'];
    if (lower.includes('/slicer')) return ['slicer'];
    if (lower.includes('/timeline')) return ['timeline'];
    if (lower.includes('vba') || lower.endsWith('.bin')) return request.options.preserveMacros === false ? [] : ['vba'];
    if (lower.includes('externalconnections') || lower.includes('connections.xml')) return ['external-connection'];
    if (lower.includes('/drawings/')) return ['images'];
    return [];
  })) : new Set<string>();
  // Only parts that survived the writer can be preserved-only. The source
  // package is not authoritative after native graph synchronization.
  for (const feature of [...preservedFeatures]) if (!packageFeatures.includes(feature)) preservedFeatures.delete(feature);
  const editableFeatures = new Set(['cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names', 'hyperlinks', 'tables']);
  if (nativeStatus.pivot) editableFeatures.add('pivot');
  if (nativeStatus.slicer) editableFeatures.add('slicer');
  if (nativeStatus.timeline) editableFeatures.add('timeline');
  for (const feature of ['slicer', 'timeline'] as const) if (detectedFeatures.includes(feature) && !editableFeatures.has(feature)) preservedFeatures.add(feature);
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
  const completedReport = refreshCompatibilitySummary({ ...report, issues: [...report.issues, ...scanFormulaPreserveIssues(request.snapshot)] });

  return {
    taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    report: completedReport,
    buffer,
    fileName: request.fileName.endsWith('.xlsx') ? request.fileName : `${request.fileName}.xlsx`,
    package: emittedPackage,
  };
}
