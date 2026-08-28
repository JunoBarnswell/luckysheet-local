import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { exportSnapshotToXlsxBuffer, loadOpcPackageGraph } from './archive';
import { detectPackageFeatures } from './ooxml';
import { nativePivotFeatureStatus } from './native-pivot';
import { createCompatibilityReport, refreshCompatibilitySummary } from './compatibility-report';
import { createNativePackageState, verifyNativePackageState } from './native-package-state';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import type { XlsxExportOptions, XlsxExportResult, NativePackageState } from './types';
import { capabilityFor, detectWorksheetCapabilities } from './capability-manifest';

export interface XlsxExportRequest {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: XlsxExportOptions;
  /** The one native package state returned by the import transaction. */
  nativePackage?: NativePackageState;
}

/** 导出 XLSX 并生成 Compatibility Report */
export async function exportXlsx(request: XlsxExportRequest): Promise<XlsxExportResult> {
  if (request.nativePackage) await verifyNativePackageState(request.nativePackage);
  const sourcePackage = request.nativePackage?.packageGraph;
  const dateSystem = request.options.dateSystem ?? sourcePackage?.dateSystem ?? request.nativePackage?.dateSystem ?? '1900';
  const buffer = exportSnapshotToXlsxBuffer(request.snapshot, sourcePackage, request.options);
  // Report the package that was actually emitted. This prevents a deleted
  // native Pivot/Slicer/Timeline from being reported as preserved merely
  // because its source package contained the old opaque part.
  const emittedPackage = loadOpcPackageGraph(buffer, {}, request.fileName).packageGraph;
  const emittedFileName = fileNameForFormat(request.fileName, emittedPackage.format.variant);
  const snapshotFeatureSet = new Set(scanSnapshotFeatures(request.snapshot));
  const packageFeatureSet = new Set(detectPackageFeatures(emittedPackage));
  const preservedNativeChartDetections = emittedPackage.nativeChartGraph?.charts.filter((chart) => !chart.editable).map((chart) => ({ feature: 'preserved-native-chart', location: chart.chartPart, reason: chart.reason })) ?? [];
  const opaqueChartParts = Object.keys(emittedPackage.opaqueParts).filter((part) => part.toLowerCase().includes('/charts/') && !emittedPackage.nativeChartGraph?.charts.some((chart) => chart.chartPart === part));
  for (const detection of preservedNativeChartDetections) packageFeatureSet.add(detection.feature);
  const snapshotFeatures = [...snapshotFeatureSet];
  const packageFeatures = [...packageFeatureSet];
  const emittedWorksheetDetections = detectWorksheetCapabilities(emittedPackage.parts, emittedPackage);
  const sourceWorksheetDetections = sourcePackage ? detectWorksheetCapabilities(sourcePackage.parts, sourcePackage) : [];
  const detectedFeatures = [...new Set([...packageFeatures, ...snapshotFeatures, ...emittedWorksheetDetections.map((entry) => entry.feature), ...sourceWorksheetDetections.map((entry) => entry.feature), ...preservedNativeChartDetections.map((entry) => entry.feature)])];
  const nativeStatus = nativePivotFeatureStatus(request.snapshot, emittedPackage.nativePivotGraph);
  const preservedFeatures = sourcePackage ? new Set(Object.keys(sourcePackage.opaqueParts).flatMap((name) => {
    const lower = name.toLowerCase();
    if (lower.includes('/charts/')) return ['charts'];
    if (lower.includes('/pivot')) return ['pivot'];
    if (lower.includes('/slicer')) return ['slicer'];
    if (lower.includes('/timeline')) return ['timeline'];
    if (isVbaPartForReport(name, sourcePackage)) return request.options.preserveMacros === false ? [] : ['vba'];
    if (lower.includes('externalconnections') || lower.includes('connections.xml')) return ['external-connection'];
    if (lower.includes('/drawings/')) return ['images'];
    return [];
  })) : new Set<string>();
  if (preservedNativeChartDetections.length) preservedFeatures.add('preserved-native-chart');
  // Only parts that survived the writer can be preserved-only. The source
  // package is not authoritative after native graph synchronization.
  for (const feature of [...preservedFeatures]) if (!packageFeatures.includes(feature)) preservedFeatures.delete(feature);
  const editableFeatures = new Set(detectedFeatures.filter((feature) => capabilityFor(feature).read !== 'none' && capabilityFor(feature).write !== 'none'));
  editableFeatures.add('defined-names');
  if (opaqueChartParts.length && !snapshotFeatureSet.has('charts') && !emittedPackage.nativeChartGraph?.charts.some((chart) => chart.editable)) editableFeatures.delete('charts');
  if (nativeStatus.pivot) editableFeatures.add('pivot');
  if (nativeStatus.slicer) editableFeatures.add('slicer');
  if (nativeStatus.timeline) editableFeatures.add('timeline');
  for (const feature of ['slicer', 'timeline'] as const) if (detectedFeatures.includes(feature) && !editableFeatures.has(feature)) preservedFeatures.add(feature);
  const unsupportedFeatures = detectedFeatures.filter((feature) => !editableFeatures.has(feature) && !preservedFeatures.has(feature));
  const projectedFeatures = new Set(['table-sheet', 'gantt-sheet', 'report-sheet', 'barcode', 'camera', 'form-control'].filter((feature) => detectedFeatures.includes(feature)));
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures: [...detectedFeatures, ...emittedWorksheetDetections, ...sourceWorksheetDetections, ...preservedNativeChartDetections],
    preservedFeatures,
    editableFeatures,
    unsupportedFeatures,
    projectedFeatures,
  });
  const completedReport = refreshCompatibilitySummary({ ...report, issues: [...report.issues, ...scanFormulaPreserveIssues(request.snapshot)] });

  return {
    taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    report: completedReport,
    buffer,
    fileName: emittedFileName,
    nativePackage: await createNativePackageState({
      fileName: request.fileName,
      buffer,
      dateSystem,
      packageGraph: emittedPackage,
      format: emittedPackage.format,
      detectedFeatures,
      compatibility: completedReport,
    }),
  };
}

function isVbaPartForReport(name: string, packageGraph: NonNullable<XlsxExportRequest['nativePackage']>['packageGraph']): boolean {
  const lower = name.toLowerCase();
  if (lower.endsWith('vbaproject.bin') || lower.endsWith('vbaprojectsignature.bin')) return true;
  for (const [source, relationships] of Object.entries(packageGraph.relationships)) {
    for (const relationship of relationships) {
      if (relationship.type.toLowerCase().includes('vbaproject')) {
        const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
        const target = relationship.target.startsWith('/') ? relationship.target.slice(1) : `${base}${relationship.target}`;
        if (target.replace(/\\/g, '/').replace(/^\/+/, '') === name.replace(/^\/+/, '')) return true;
      }
    }
  }
  return false;
}

function fileNameForFormat(input: string, variant: string): string {
  const expected = `.${variant}`;
  const known = /\.(xlsx|xlsm|xltx|xltm|xlam|xlsb|xls|csv|txt|ods)$/i;
  if (input.toLowerCase().endsWith(expected)) return input;
  return `${input.replace(known, '')}${expected}`;
}
