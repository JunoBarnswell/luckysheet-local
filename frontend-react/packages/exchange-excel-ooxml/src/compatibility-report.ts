import type { CompatibilityIssue, CompatibilityLevel, CompatibilityReport, DateSystem } from './types';

const LEVEL_A_FEATURES = new Set([
  'cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names',
  'hyperlinks', 'comments', 'filters', 'validation', 'conditional-format', 'tables', 'images',
  'rich-text', 'split', 'outline',
]);

const LEVEL_B_FEATURES = new Set([
  'charts', 'sparklines', 'pivot', 'slicer', 'timeline', 'print-setup', 'theme', 'protection', 'extended-validation', 'extended-conditional-format', 'cell-style-template',
]);

const LEVEL_C_FEATURES = new Set([
  'vba', 'unknown-extension', 'external-connection', 'macro', 'external-workbook',
  'cube', 'indirect', 'getpivotdata', 'unknown-worksheet-node',
]);

export function classifyFeature(feature: string): CompatibilityLevel {
  if (LEVEL_A_FEATURES.has(feature)) return 'A';
  if (LEVEL_B_FEATURES.has(feature)) return 'B';
  if (LEVEL_C_FEATURES.has(feature)) return 'C';
  // Unknown extensions are never treated as editable by default.
  return 'C';
}

export type CompatibilityFeatureDetection = {
  feature: string;
  location?: string;
  reason?: string;
};

/**
 * Assign every detected capability to exactly one honest state.  The report
 * is deliberately based on the reader/writer capability sets, not just the
 * compatibility level requested by a caller.
 */
export function createCompatibilityReport(input: {
  fileName: string;
  importLevel: CompatibilityLevel;
  exportLevel: CompatibilityLevel;
  dateSystem: DateSystem;
  detectedFeatures: Iterable<string | CompatibilityFeatureDetection>;
  unsupportedFeatures?: Iterable<string>;
  /** Features copied byte-for-byte from the source package. */
  preservedFeatures?: Iterable<string>;
  /** Features serialized by the editable snapshot writer. */
  editableFeatures?: Iterable<string>;
}): CompatibilityReport {
  const preserved = new Set(input.preservedFeatures ?? []);
  const editable = new Set(input.editableFeatures ?? LEVEL_A_FEATURES);
  const unsupported = new Set(input.unsupportedFeatures ?? []);
  const detections = new Map<string, CompatibilityFeatureDetection>();

  for (const value of input.detectedFeatures) {
    const detection = typeof value === 'string' ? { feature: value } : value;
    if (!detection.feature) continue;
    const key = `${detection.feature}\u0000${detection.location ?? ''}`;
    const existing = detections.get(key);
    if (!existing || (!existing.reason && detection.reason)) detections.set(key, detection);
  }
  for (const feature of unsupported) {
    const key = `${feature}\u0000`;
    if (!detections.has(key)) detections.set(key, { feature });
  }

  const issues: CompatibilityIssue[] = [];
  for (const detection of detections.values()) {
    const feature = detection.feature;
    const level = classifyFeature(feature);
    const status: CompatibilityIssue['status'] = unsupported.has(feature)
      ? 'unsupported'
      : editable.has(feature)
        ? 'editable'
        : preserved.has(feature)
          ? 'preserved-only'
          : 'unsupported';
    const reason = detection.reason ?? reasonFor(status, feature);
    const severity = status === 'unsupported'
      ? 'error'
      : status === 'preserved-only'
        ? (level === 'C' ? 'info' : 'warning')
        : 'info';
    issues.push({
      level,
      severity,
      feature,
      ...(detection.location ? { location: detection.location } : {}),
      message: messageFor(status, feature, reason),
      preserved: status === 'preserved-only',
      status,
      reason,
    });
  }

  return {
    schema: 'CompatibilityReport',
    fileName: input.fileName,
    importLevel: input.importLevel,
    exportLevel: input.exportLevel,
    dateSystem: input.dateSystem,
    issues,
    summary: {
      editableFeatures: issues.filter((issue) => issue.status === 'editable').length,
      preservedOnly: issues.filter((issue) => issue.status === 'preserved-only').length,
      unsupported: issues.filter((issue) => issue.status === 'unsupported').length,
    },
  };
}

export function mergeReports(...reports: CompatibilityReport[]): CompatibilityReport {
  if (reports.length === 0) throw new Error('No reports to merge');
  const first = reports[0]!;
  const issues = new Map<string, CompatibilityIssue>();
  for (const report of reports) {
    for (const issue of report.issues) {
      const key = `${issue.feature}\u0000${issue.location ?? ''}\u0000${issue.reason}`;
      if (!issues.has(key)) issues.set(key, issue);
    }
  }
  const merged = [...issues.values()];
  return {
    ...first,
    issues: merged,
    summary: {
      editableFeatures: merged.filter((issue) => issue.status === 'editable').length,
      preservedOnly: merged.filter((issue) => issue.status === 'preserved-only').length,
      unsupported: merged.filter((issue) => issue.status === 'unsupported').length,
    },
  };
}

export function refreshCompatibilitySummary(report: CompatibilityReport): CompatibilityReport {
  const unique = new Map<string, CompatibilityIssue>();
  for (const issue of report.issues) {
    const key = `${issue.feature}\u0000${issue.location ?? ''}\u0000${issue.reason}`;
    if (!unique.has(key)) unique.set(key, issue);
  }
  const issues = [...unique.values()];
  return {
    ...report,
    issues,
    summary: {
      editableFeatures: issues.filter((issue) => issue.status === 'editable').length,
      preservedOnly: issues.filter((issue) => issue.status === 'preserved-only').length,
      unsupported: issues.filter((issue) => issue.status === 'unsupported').length,
    },
  };
}

function reasonFor(status: CompatibilityIssue['status'], feature: string): string {
  switch (status) {
    case 'editable': return 'serialized by the canonical workbook writer';
    case 'preserved-only': return 'copied from the source OOXML package without editable model support';
    case 'unsupported': return `no validated reader/writer contract exists for ${feature}`;
  }
}

function messageFor(status: CompatibilityIssue['status'], feature: string, reason: string): string {
  switch (status) {
    case 'editable': return `${feature} is editable: ${reason}`;
    case 'preserved-only': return `${feature} is preserved-only: ${reason}`;
    case 'unsupported': return `${feature} is unsupported: ${reason}`;
  }
}
