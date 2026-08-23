import type { CompatibilityIssue, CompatibilityLevel, CompatibilityReport, DateSystem } from './types';

const LEVEL_A_FEATURES = new Set([
  'cells', 'formulas', 'styles', 'merges', 'freeze', 'defined-names',
  'hyperlinks', 'comments', 'filters', 'validation', 'conditional-format', 'tables', 'images',
]);

const LEVEL_B_FEATURES = new Set([
  'charts', 'sparklines', 'pivot', 'slicer', 'print-setup', 'theme',
]);

const LEVEL_C_FEATURES = new Set([
  'vba', 'unknown-extension', 'external-connection', 'macro',
]);

export function classifyFeature(feature: string): CompatibilityLevel {
  if (LEVEL_A_FEATURES.has(feature)) return 'A';
  if (LEVEL_B_FEATURES.has(feature)) return 'B';
  return 'C';
}

export function createCompatibilityReport(input: {
  fileName: string;
  importLevel: CompatibilityLevel;
  exportLevel: CompatibilityLevel;
  dateSystem: DateSystem;
  detectedFeatures: string[];
  unsupportedFeatures?: string[];
  /** Features copied byte-for-byte from the source package. */
  preservedFeatures?: Iterable<string>;
  /** Features serialized by the editable snapshot writer. */
  editableFeatures?: Iterable<string>;
}): CompatibilityReport {
  const issues: CompatibilityIssue[] = [];
  const preserved = new Set(input.preservedFeatures ?? []);
  const editable = new Set(input.editableFeatures ?? LEVEL_A_FEATURES);
  const unsupported = new Set(input.unsupportedFeatures ?? []);

  for (const feature of input.detectedFeatures) {
    const level = classifyFeature(feature);
    if (unsupported.has(feature)) {
      issues.push({
        level,
        severity: 'error',
        feature,
        message: `${feature} is not supported by this export path`,
        preserved: false,
      });
    } else if (preserved.has(feature)) {
      issues.push({
        level,
        severity: level === 'C' ? 'info' : 'warning',
        feature,
        message: `${feature} is preserved as original OOXML and is not edited by the snapshot writer`,
        preserved: true,
      });
    } else if (level === 'B' && input.importLevel === 'A') {
      issues.push({
        level: 'B',
        severity: 'warning',
        feature,
        message: `${feature} imported with limited edit support`,
        preserved: true,
      });
    } else if (level === 'A' && !editable.has(feature)) {
      issues.push({
        level: 'A',
        severity: 'error',
        feature,
        message: `${feature} is not supported by this export path`,
        preserved: false,
      });
    }
  }

  for (const feature of input.unsupportedFeatures ?? []) {
    issues.push({
      level: 'C',
      severity: 'error',
      feature,
      message: `${feature} is not supported`,
      preserved: false,
    });
  }

  return {
    schema: 'CompatibilityReportV1',
    fileName: input.fileName,
    importLevel: input.importLevel,
    exportLevel: input.exportLevel,
    dateSystem: input.dateSystem,
    issues,
    summary: {
      editableFeatures: issues.filter((i) => i.level === 'A' && i.preserved).length,
      preservedOnly: issues.filter((i) => i.level === 'C' && i.preserved).length,
      unsupported: issues.filter((i) => !i.preserved).length,
    },
  };
}

export function mergeReports(...reports: CompatibilityReport[]): CompatibilityReport {
  if (reports.length === 0) {
    throw new Error('No reports to merge');
  }
  const first = reports[0]!;
  return {
    ...first,
    issues: reports.flatMap((r) => r.issues),
    summary: {
      editableFeatures: reports.reduce((n, r) => n + r.summary.editableFeatures, 0),
      preservedOnly: reports.reduce((n, r) => n + r.summary.preservedOnly, 0),
      unsupported: reports.reduce((n, r) => n + r.summary.unsupported, 0),
    },
  };
}
