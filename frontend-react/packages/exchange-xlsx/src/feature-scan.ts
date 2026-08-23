import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { classifyFeature } from './compatibility-report';
import type { CompatibilityIssue, CompatibilityLevel } from './types';

const PRESERVE_FORMULA_PATTERNS: Array<{ feature: string; pattern: RegExp }> = [
  { feature: 'external-workbook', pattern: /\[[^\]]+\]/ },
  { feature: 'indirect', pattern: /\bINDIRECT\s*\(/i },
  { feature: 'getpivotdata', pattern: /\bGETPIVOTDATA\s*\(/i },
  { feature: 'cube', pattern: /\bCUBE(?:VALUE|SET|MEMBER)\s*\(/i },
];

export function scanSnapshotFeatures(snapshot: WorkbookSnapshot): string[] {
  const features = new Set<string>(['cells', 'styles']);
  let hasFormula = false;

  for (const sheet of snapshot.sheets) {
    for (const rowKey of Object.keys(sheet.cells)) {
      const row = sheet.cells[rowKey] ?? {};
      for (const colKey of Object.keys(row)) {
        const cell = row[colKey];
        if (!cell) continue;
        if (cell.formula) hasFormula = true;
        if (cell.numberFormat) features.add('styles');
      }
    }
    if (sheet.merges.length > 0) features.add('merges');
    if (sheet.pane?.kind === 'frozen') features.add('freeze');
    if (sheet.pane?.kind === 'split') features.add('split');
    if (sheet.pivots.length > 0) features.add('pivot');
    // Floating objects are represented by one canonical collection. Scan the
    // payloads instead of consulting removed per-kind arrays so XLSX reports
    // cannot silently miss a chart/image created through the command runtime.
    for (const payload of Object.values(sheet.drawingPayloads)) {
      if (payload.kind === 'chart') features.add('charts');
      else if (payload.kind === 'slicer') features.add('slicer');
      else if (payload.kind === 'timeline') features.add('timeline');
      else if (payload.kind === 'image' || payload.kind === 'shape' || payload.kind === 'textbox') features.add('images');
    }
    if (sheet.sparklines.length > 0 || sheet.sparklineGroups?.length) features.add('sparklines');
    if ((sheet.conditionalFormats?.length ?? 0) > 0) features.add('conditional-format');
    if ((sheet.dataValidations?.length ?? 0) > 0) features.add('validation');
    if ((sheet.sheetTables?.length ?? 0) > 0) features.add('tables');
    if (sheet.filter) features.add('filters');
    if (sheet.protectionRules?.length) features.add('protection');
    if ((sheet.commentThreads?.length ?? 0) > 0) features.add('comments');
    if (sheet.notes?.length) features.add('comments');
  }

  if (hasFormula) features.add('formulas');
  if (snapshot.definedNames && Object.keys(snapshot.definedNames).length > 0) {
    features.add('defined-names');
  }

  return [...features];
}

export function scanFormulaPreserveIssues(snapshot: WorkbookSnapshot): CompatibilityIssue[] {
  const issues: CompatibilityIssue[] = [];
  const seen = new Set<string>();

  for (const sheet of snapshot.sheets) {
    for (const rowKey of Object.keys(sheet.cells)) {
      const row = sheet.cells[rowKey] ?? {};
      for (const colKey of Object.keys(row)) {
        const formula = row[colKey]?.formula;
        if (!formula) continue;
        for (const rule of PRESERVE_FORMULA_PATTERNS) {
          if (!rule.pattern.test(formula) || seen.has(rule.feature)) continue;
          seen.add(rule.feature);
          issues.push({
            level: classifyFeature(rule.feature) === 'A' ? 'B' : 'C',
            severity: 'info',
            feature: rule.feature,
            location: `${sheet.name}!${colKey}${Number(rowKey) + 1}`,
            message: `${rule.feature} preserved on import; recalculation may differ from Excel`,
            preserved: true,
            status: 'preserved-only',
            reason: 'formula depends on an external or Excel-only calculation contract',
          });
        }
      }
    }
  }

  return issues;
}

export function summarizeCompatibilityLevel(features: string[]): CompatibilityLevel {
  let level: CompatibilityLevel = 'A';
  for (const feature of features) {
    const classified = classifyFeature(feature);
    if (classified === 'C') return 'C';
    if (classified === 'B') level = 'B';
  }
  return level;
}
