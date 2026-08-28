import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { exportSnapshotToOoxmlBuffer, loadOpcPackageGraph } from './archive';
import { detectPackageFeatures } from './ooxml';
import { nativePivotFeatureStatus } from './native-pivot';
import { createCompatibilityReport, refreshCompatibilitySummary } from './compatibility-report';
import { createNativeDocumentArtifact, nativeSnapshotHash, verifyNativeDocumentArtifact } from './native-document-artifact';
import { NativeDocumentError } from './native-document-error';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import type { NativeDocumentExportOptions, NativeDocumentExportResult, NativeDocumentArtifact, OpcPackageGraph } from './types';
import { capabilityFor, detectWorksheetCapabilities } from './capability-manifest';

export interface NativeDocumentExportRequest {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: NativeDocumentExportOptions;
  /** The one native document artifact returned by the import transaction. */
  artifact?: NativeDocumentArtifact;
  mode?: 'save' | 'save-as' | 'export';
}

/** Export an OOXML document and generate its Compatibility Report. */
export async function exportOoxmlDocument(request: NativeDocumentExportRequest): Promise<NativeDocumentExportResult> {
  if (request.artifact) await verifyNativeDocumentArtifact(request.artifact);
  if (request.artifact
    && request.artifact.nativeGraph.kind === 'opc'
    && !request.artifact.nativeGraph.package.nativePivotGraph
    && request.artifact.fileName === request.fileName
    && request.artifact.sourceSnapshotHash === nativeSnapshotHash(request.snapshot)) {
    return {
      taskId: `export-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      report: structuredClone(request.artifact.compatibility),
      buffer: request.artifact.sourceBytes.slice(0),
      fileName: request.fileName,
      artifact: request.artifact,
    };
  }
  const sourcePackage = request.artifact?.nativeGraph.kind === 'opc' ? request.artifact.nativeGraph.package : undefined;
  const targetFormat = ooxmlTargetFormat(request.fileName, sourcePackage);
  if (sourcePackage && targetFormat && targetFormat.variant !== sourcePackage.format.variant && hasMacroParts(sourcePackage) && !macroVariant(targetFormat.variant)) {
    throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_UNSUPPORTED', message: `Save As ${targetFormat.variant} would discard the source macro project`, format: targetFormat, recovery: 'Choose a macro-enabled target or explicitly remove the macro project in a dedicated conversion workflow.' });
  }
  const dateSystem = request.options.dateSystem ?? sourcePackage?.dateSystem ?? (request.artifact?.dateSystem ?? '1900');
  const buffer = exportSnapshotToOoxmlBuffer(request.snapshot, sourcePackage, { ...request.options, targetFormat });
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
  const projectedFeatures = new Set(['table-sheet', 'gantt-sheet', 'report-sheet', 'barcode', 'camera', 'screenshot', 'form-control', 'icons', 'models3d', 'smartart', 'wordart', 'signature-line', 'embedded-object', 'equation'].filter((feature) => detectedFeatures.includes(feature)));
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
    artifact: await createNativeDocumentArtifact({
      fileName: request.fileName,
      buffer,
      dateSystem,
      nativeGraph: { kind: 'opc', package: emittedPackage },
      format: emittedPackage.format,
      snapshot: request.snapshot,
      detectedFeatures,
      compatibility: completedReport,
    }),
  };
}

function isVbaPartForReport(name: string, packageGraph: OpcPackageGraph): boolean {
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

function ooxmlTargetFormat(fileName: string, source?: OpcPackageGraph): Extract<import('./types').NativeDocumentFormat, { family: 'ooxml' }> | undefined {
  const extension = fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  if (!extension || !['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'].includes(extension)) return source?.format;
  return { family: 'ooxml', profile: source?.profile ?? 'transitional', variant: extension as Extract<import('./types').NativeDocumentFormat, { family: 'ooxml' }>['variant'] };
}

function macroVariant(variant: Extract<import('./types').NativeDocumentFormat, { family: 'ooxml' }>['variant']): boolean {
  return variant === 'xlsm' || variant === 'xltm' || variant === 'xlam';
}

function hasMacroParts(packageGraph: OpcPackageGraph): boolean {
  return Object.keys(packageGraph.parts).some((name) => isVbaPartForReport(name, packageGraph));
}

function fileNameForFormat(input: string, variant: string): string {
  const expected = `.${variant}`;
  const known = /\.(xlsx|xlsm|xltx|xltm|xlam|xlsb|xls|csv|txt|ods)$/i;
  if (input.toLowerCase().endsWith(expected)) return input;
  return `${input.replace(known, '')}${expected}`;
}
