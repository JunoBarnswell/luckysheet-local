import { parseDateSystem } from './date-system';
import { createCompatibilityReport, refreshCompatibilitySummary } from './compatibility-report';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import { detectPackageFeatures, loadOpcPackageGraph, parseLoadedXlsx } from './ooxml';
import { nativePivotFeatureStatus } from './native-pivot';
import { createNativePackageState } from './native-package-state';
import type { XlsxImportOptions, XlsxImportResult } from './types';
import { sanitizeImportedWorkbookName } from './ooxml-metrics';
import { capabilityFor, detectWorksheetCapabilities } from './capability-manifest';

export interface XlsxImportRequest {
  fileName: string;
  buffer: ArrayBuffer | Uint8Array;
  options: XlsxImportOptions;
}

/** 解析 XLSX 并生成 Compatibility Report */
export async function importXlsx(request: XlsxImportRequest): Promise<XlsxImportResult> {
  const bytes = request.buffer instanceof Uint8Array ? request.buffer.slice() : new Uint8Array(request.buffer);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const loaded = loadOpcPackageGraph(buffer, request.options.limits, request.fileName);
  const importedName = sanitizeImportedWorkbookName(request.fileName);
  const parsed = parseLoadedXlsx(loaded, { workbookName: importedName });
  const snapshot = parsed.snapshot;
  snapshot.name = importedName;
  const dateSystem = request.options.dateSystem ?? parsed.packageGraph.dateSystem ?? parseDateSystem('');
  const snapshotFeatures = scanSnapshotFeatures(snapshot);
  const packageFeatures = detectPackageFeatures(parsed.packageGraph);
  const worksheetDetections = detectWorksheetCapabilities(loaded.files, parsed.packageGraph);
  const mode = request.options.compatibilityMode ?? (request.options.compatibilityTarget === 'A' ? 'strict' : request.options.compatibilityTarget === 'C' ? 'best-effort' : 'balanced');
  const capabilityDetections = worksheetDetections.map((detection) => {
    const capability = capabilityFor(detection.feature);
    const reason = detection.reason ?? (capability.read === 'partial' || capability.write === 'partial'
      ? mode === 'best-effort'
        ? capability.approximation ?? 'Imported through an explicitly bounded approximate conversion'
        : 'The editable canonical subset is imported and the original OOXML package is retained'
      : undefined);
    return reason ? { ...detection, reason } : detection;
  });
  const detectedFeatures = [...new Set([...packageFeatures, ...snapshotFeatures, ...worksheetDetections.map((entry) => entry.feature)])];
  const nativeStatus = nativePivotFeatureStatus(snapshot, parsed.packageGraph.nativePivotGraph);
  const editableFeatures = new Set(detectedFeatures.filter((feature) => capabilityFor(feature).read !== 'none' && capabilityFor(feature).write !== 'none'));
  editableFeatures.add('defined-names');
  if (nativeStatus.pivot) editableFeatures.add('pivot');
  if (nativeStatus.slicer) editableFeatures.add('slicer');
  if (nativeStatus.timeline) editableFeatures.add('timeline');
  const preservedFeatures = new Set(detectedFeatures.filter((feature) => !editableFeatures.has(feature) && capabilityFor(feature).preserve !== 'none'));
  for (const feature of ['slicer', 'timeline'] as const) if (snapshotFeatures.includes(feature) && !editableFeatures.has(feature)) preservedFeatures.add(feature);
  const report = createCompatibilityReport({
    fileName: request.fileName,
    importLevel: request.options.compatibilityTarget,
    exportLevel: request.options.compatibilityTarget,
    dateSystem,
    detectedFeatures: [...detectedFeatures, ...capabilityDetections],
    preservedFeatures,
    editableFeatures,
    unsupportedFeatures: detectedFeatures.filter((feature) => !editableFeatures.has(feature) && !preservedFeatures.has(feature)),
    projectedFeatures: ['table-sheet', 'gantt-sheet', 'report-sheet', 'barcode', 'camera', 'form-control'].filter((feature) => detectedFeatures.includes(feature)),
  });
  const completedReport = refreshCompatibilitySummary({ ...report, issues: [...report.issues, ...scanFormulaPreserveIssues(snapshot)] });
  if (mode === 'strict') {
    const unsafe = completedReport.issues.filter((issue) => issue.status === 'unsupported');
    if (unsafe.length) throw new Error(`Strict XLSX import rejected unsafe capabilities: ${unsafe.map((issue) => `${issue.feature}${issue.location ? ` at ${issue.location}` : ''}`).join(', ')}`);
  }
  const nativePackage = await createNativePackageState({
    fileName: request.fileName,
    buffer,
    dateSystem,
    packageGraph: parsed.packageGraph,
    format: parsed.packageGraph.format,
    detectedFeatures,
    compatibility: completedReport,
  });

  return {
    payload: {
      name: importedName,
      sheetCount: snapshot.sheets.length,
      dateSystem,
      compatibilityLevel: request.options.compatibilityTarget,
    },
    report: completedReport,
    snapshot,
    nativePackage,
    taskId: `import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  };
}

export { parseDateSystem };
