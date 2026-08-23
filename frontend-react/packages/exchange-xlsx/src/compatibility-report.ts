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
}): CompatibilityReport {
  const issues: CompatibilityIssue[] = [];

  for (const feature of input.detectedFeatures) {
    const level = classifyFeature(feature);
    if (level === 'C') {
      issues.push({
        level: 'C',
        severity: 'info',
        feature,
        message: `${feature} will be preserved but not editable`,
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
