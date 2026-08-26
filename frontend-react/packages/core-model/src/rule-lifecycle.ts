import type {
  CellAddress,
  ConditionalFormatRule,
  ConditionalFormatTopBottom,
  DataValidationRule,
  RangeRef,
  WorksheetModel,
} from './index';

export type SheetRule = ConditionalFormatRule | DataValidationRule;
export type SheetRuleKind = 'conditional-format' | 'data-validation';

export interface RuleTransform {
  mapRange: (range: RangeRef) => readonly RangeRef[];
  mapAddress: (address: CellAddress) => CellAddress;
}

export interface RulePasteTransform {
  source: RangeRef;
  target: RangeRef;
  transpose: boolean;
  id: (rule: SheetRule) => string;
}

function isConditionalFormat(rule: SheetRule): rule is ConditionalFormatRule {
  return 'operator' in rule && 'stopIfTrue' in rule;
}

function normalizeRange(range: RangeRef): RangeRef {
  if (!range.sheetId.trim() || !Number.isSafeInteger(range.startRow) || !Number.isSafeInteger(range.endRow)
    || !Number.isSafeInteger(range.startColumn) || !Number.isSafeInteger(range.endColumn)
    || range.startRow < 0 || range.endRow < 0 || range.startColumn < 0 || range.endColumn < 0) {
    throw new Error('Sheet rule range is invalid');
  }
  return {
    ...range,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function normalizeRanges(sheetId: string, ranges: readonly RangeRef[]): RangeRef[] {
  if (ranges.length === 0) throw new Error('Sheet rule requires at least one range');
  const normalized = ranges.map(normalizeRange);
  if (normalized.some((range) => range.sheetId !== sheetId)) throw new Error('Sheet rule ranges must target the rule sheet');
  return normalized;
}

function normalizeAnchor(sheetId: string, ranges: readonly RangeRef[], anchor?: CellAddress): CellAddress {
  if (anchor === undefined) return { sheetId, row: ranges[0]!.startRow, column: ranges[0]!.startColumn };
  if (anchor.sheetId !== sheetId || !Number.isSafeInteger(anchor.row) || !Number.isSafeInteger(anchor.column) || anchor.row < 0 || anchor.column < 0) {
    throw new Error('Sheet rule formula anchor is invalid');
  }
  return structuredClone(anchor);
}

function subtractRange(source: RangeRef, clear: RangeRef): RangeRef[] {
  if (source.sheetId !== clear.sheetId
    || source.startRow > clear.endRow || clear.startRow > source.endRow
    || source.startColumn > clear.endColumn || clear.startColumn > source.endColumn) {
    return [structuredClone(source)];
  }
  const top = Math.max(source.startRow, clear.startRow);
  const bottom = Math.min(source.endRow, clear.endRow);
  const left = Math.max(source.startColumn, clear.startColumn);
  const right = Math.min(source.endColumn, clear.endColumn);
  const result: RangeRef[] = [];
  if (source.startRow < top) result.push({ ...source, endRow: top - 1 });
  if (bottom < source.endRow) result.push({ ...source, startRow: bottom + 1 });
  if (source.startColumn < left) result.push({ ...source, startRow: top, endRow: bottom, endColumn: left - 1 });
  if (right < source.endColumn) result.push({ ...source, startRow: top, endRow: bottom, startColumn: right + 1 });
  return result;
}

function intersectRange(source: RangeRef, clip: RangeRef): RangeRef | undefined {
  if (source.sheetId !== clip.sheetId
    || source.startRow > clip.endRow || clip.startRow > source.endRow
    || source.startColumn > clip.endColumn || clip.startColumn > source.endColumn) return undefined;
  return {
    sheetId: source.sheetId,
    startRow: Math.max(source.startRow, clip.startRow),
    endRow: Math.min(source.endRow, clip.endRow),
    startColumn: Math.max(source.startColumn, clip.startColumn),
    endColumn: Math.min(source.endColumn, clip.endColumn),
  };
}

function remapPasteRange(range: RangeRef, transform: RulePasteTransform): RangeRef {
  if (transform.transpose) {
    return {
      sheetId: transform.target.sheetId,
      startRow: transform.target.startRow + (range.startColumn - transform.source.startColumn),
      endRow: transform.target.startRow + (range.endColumn - transform.source.startColumn),
      startColumn: transform.target.startColumn + (range.startRow - transform.source.startRow),
      endColumn: transform.target.startColumn + (range.endRow - transform.source.startRow),
    };
  }
  return {
    sheetId: transform.target.sheetId,
    startRow: transform.target.startRow + (range.startRow - transform.source.startRow),
    endRow: transform.target.startRow + (range.endRow - transform.source.startRow),
    startColumn: transform.target.startColumn + (range.startColumn - transform.source.startColumn),
    endColumn: transform.target.startColumn + (range.endColumn - transform.source.startColumn),
  };
}

function exactlyOneRange(ranges: readonly RangeRef[], label: string): RangeRef {
  if (ranges.length !== 1) throw new Error(`${label} cannot be represented as one range after transform`);
  return ranges[0]!;
}

/**
 * Canonical lifecycle owner for worksheet rules. Every caller uses this
 * registry for range normalization, structural transforms, Clear cropping,
 * and Paste cloning; rule-specific evaluation remains in the feature layer.
 */
export class SheetRuleRegistry {
  normalizeConditionalFormat(
    rule: ConditionalFormatRule,
    fallbackPriority = 1,
    validateFormula?: (formula: string) => void,
  ): ConditionalFormatRule {
    const ranges = normalizeRanges(rule.sheetId, rule.ranges);
    if (!rule.id.trim()) throw new Error('Conditional format id is required');
    if (rule.priority !== undefined && (!Number.isInteger(rule.priority) || rule.priority <= 0)) throw new Error('Conditional format priority must be positive');
    if (rule.stopIfTrue !== undefined && typeof rule.stopIfTrue !== 'boolean') throw new Error('Conditional format stopIfTrue must be boolean');
    if (rule.operator === 'formula') {
      const formula = String(rule.value1 ?? '').trim();
      if (!formula) throw new Error(`Conditional format ${rule.id} requires a formula predicate`);
      validateFormula?.(formula.startsWith('=') ? formula : `=${formula}`);
    }
    const topBottom: ConditionalFormatTopBottom | undefined = rule.type === 'topBottom'
      ? rule.topBottom ?? { direction: rule.operator === 'bottom' ? 'bottom' : 'top', rank: Number(rule.value1 ?? 10) }
      : undefined;
    if (topBottom && (!Number.isFinite(topBottom.rank) || topBottom.rank <= 0 || (topBottom.percent && topBottom.rank > 100))) {
      throw new Error('Conditional format top/bottom rank is invalid');
    }
    return {
      ...structuredClone(rule),
      ranges,
      formulaAnchor: normalizeAnchor(rule.sheetId, ranges, rule.formulaAnchor),
      priority: Number.isInteger(rule.priority) && (rule.priority ?? 0) > 0 ? rule.priority : fallbackPriority,
      stopIfTrue: rule.stopIfTrue ?? false,
      ...(topBottom ? { topBottom: structuredClone(topBottom) } : {}),
    };
  }

  normalizeDataValidation(
    rule: DataValidationRule,
    validateFormula?: (formula: string) => void,
  ): DataValidationRule {
    const ranges = normalizeRanges(rule.sheetId, rule.ranges);
    if (!rule.id.trim()) throw new Error('Data validation id is required');
    if (rule.type === 'list' && !rule.formula1 && !rule.listSource) throw new Error(`List validation ${rule.id} requires a list source`);
    if ((rule.type === 'whole' || rule.type === 'decimal' || rule.type === 'textLength') && rule.formula1 === undefined) throw new Error(`Data validation ${rule.id} requires a lower bound`);
    if (rule.type === 'checkbox' && rule.operator !== undefined) throw new Error('Checkbox validation does not accept a comparison operator');
    if (rule.alertStyle !== undefined && !['stop', 'warning', 'information'].includes(rule.alertStyle)) throw new Error('Data validation alert style is invalid');
    if (rule.listSource?.kind === 'range' && rule.listSource.range.sheetId !== rule.sheetId) throw new Error('Validation list range must target the validation sheet');
    const formula = rule.type === 'custom' ? rule.formula1 : rule.listSource?.kind === 'formula' ? rule.listSource.formula : undefined;
    if (formula?.trim().startsWith('=')) validateFormula?.(formula.trim());
    return {
      ...structuredClone(rule),
      ranges,
      formulaAnchor: normalizeAnchor(rule.sheetId, ranges, rule.formulaAnchor),
      alertStyle: rule.alertStyle ?? 'stop',
      showErrorMessage: rule.showErrorMessage ?? true,
      showInputMessage: rule.showInputMessage ?? false,
      allowBlank: rule.allowBlank ?? true,
    };
  }

  affectedRanges(rule: SheetRule): RangeRef[] {
    return structuredClone(rule.ranges);
  }

  affectedColumnEnd(sheet: Pick<WorksheetModel, 'columnCount' | 'conditionalFormats' | 'dataValidations' | 'protectionRules'>, baseline: number): number {
    let end = Math.max(baseline, sheet.columnCount - 1);
    for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
      for (const range of rule.ranges) end = Math.max(end, range.endColumn);
    }
    for (const rule of sheet.protectionRules) if (rule.range) end = Math.max(end, rule.range.endColumn);
    return end;
  }

  transform<T extends SheetRule>(rule: T, transform: RuleTransform): T {
    const next = structuredClone(rule) as T;
    next.ranges = rule.ranges.flatMap((range) => transform.mapRange(range).map((mapped) => structuredClone(mapped)));
    if (next.ranges.length === 0) throw new Error(`Rule ${rule.id} has no range after structural transform`);
    if (rule.formulaAnchor) next.formulaAnchor = transform.mapAddress(rule.formulaAnchor);
    if (!isConditionalFormat(rule) && rule.listSource?.kind === 'range') {
      const validation = next as unknown as DataValidationRule;
      const mapped = transform.mapRange(rule.listSource.range);
      validation.listSource = { ...rule.listSource, range: exactlyOneRange(mapped, `Validation ${rule.id} list source`) };
    }
    return next;
  }

  crop<T extends SheetRule>(rule: T, clear: RangeRef): T | undefined {
    const ranges = rule.ranges.flatMap((range) => subtractRange(range, clear));
    if (ranges.length === 0) return undefined;
    return { ...structuredClone(rule), ranges } as T;
  }

  cloneForPaste<T extends SheetRule>(rule: T, transform: RulePasteTransform): T {
    const next = structuredClone(rule) as T;
    next.id = transform.id(rule);
    next.sheetId = transform.target.sheetId;
    next.ranges = rule.ranges.map((range) => remapPasteRange(range, transform));
    if (rule.formulaAnchor) {
      const formulaAnchor = {
        sheetId: transform.target.sheetId,
        row: transform.target.startRow + (transform.transpose ? rule.formulaAnchor.column - transform.source.startColumn : rule.formulaAnchor.row - transform.source.startRow),
        column: transform.target.startColumn + (transform.transpose ? rule.formulaAnchor.row - transform.source.startRow : rule.formulaAnchor.column - transform.source.startColumn),
      };
      if (formulaAnchor.row < 0 || formulaAnchor.column < 0) throw new Error(`Rule ${rule.id} formula anchor cannot be represented at paste target`);
      next.formulaAnchor = formulaAnchor;
    }
    if (!isConditionalFormat(rule) && rule.listSource?.kind === 'range') {
      const validation = next as unknown as DataValidationRule;
      validation.listSource = { ...rule.listSource, range: remapPasteRange(rule.listSource.range, transform) };
    }
    return next;
  }

  cropRules<T extends SheetRule>(rules: readonly T[], clear: RangeRef): T[] {
    return rules.flatMap((rule) => {
      const cropped = this.crop(rule, clear);
      return cropped ? [cropped] : [];
    });
  }

  cloneRulesForPaste<T extends SheetRule>(rules: readonly T[], transform: RulePasteTransform): T[] {
    return rules.flatMap((rule) => {
      const ranges = rule.ranges.flatMap((range) => {
        const clipped = intersectRange(range, transform.source);
        return clipped ? [clipped] : [];
      });
      return ranges.length === 0 ? [] : [this.cloneForPaste({ ...structuredClone(rule), ranges } as T, transform)];
    });
  }
}

export const sheetRuleRegistry = new SheetRuleRegistry();

export function ruleRangesIntersect(rule: SheetRule, range: RangeRef): boolean {
  return rule.ranges.some((candidate) => candidate.sheetId === range.sheetId
    && candidate.startRow <= range.endRow && range.startRow <= candidate.endRow
    && candidate.startColumn <= range.endColumn && range.startColumn <= candidate.endColumn);
}
