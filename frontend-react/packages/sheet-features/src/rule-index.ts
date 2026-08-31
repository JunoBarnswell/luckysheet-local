import type { DataValidationRule, RangeRef, WorksheetModel } from '@react-sheets/core-model';

export interface RuleIntervalEntry<T> {
  readonly range: RangeRef;
  readonly value: T;
}

/** Compact interval index used by all rule lookups; query cost is bounded by matching intervals. */
export class RuleIntervalIndex<T> {
  private readonly entries: RuleIntervalEntry<T>[] = [];
  private sorted = true;

  constructor(entries: readonly RuleIntervalEntry<T>[] = []) {
    for (const entry of entries) this.add(entry.range, entry.value);
  }

  add(range: RangeRef, value: T): void {
    this.entries.push({ range: structuredClone(range), value });
    this.sorted = false;
  }

  query(sheetId: string, row: number, column: number): readonly T[] {
    if (!this.sorted) {
      this.entries.sort((left, right) => left.range.startRow - right.range.startRow || left.range.startColumn - right.range.startColumn);
      this.sorted = true;
    }
    let low = 0;
    let high = this.entries.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (this.entries[middle]!.range.startRow <= row) low = middle + 1;
      else high = middle;
    }
    const result: T[] = [];
    for (let index = low - 1; index >= 0; index -= 1) {
      const entry = this.entries[index]!;
      if (entry.range.startRow > row) continue;
      if (entry.range.endRow < row) continue;
      if (entry.range.sheetId === sheetId && entry.range.startColumn <= column && column <= entry.range.endColumn) result.push(entry.value);
    }
    return result;
  }

  get size(): number { return this.entries.length; }
}

const validationIndexCache = new WeakMap<WorksheetModel, RuleIntervalIndex<DataValidationRule>>();
const validationIndexSignatures = new WeakMap<WorksheetModel, string>();

export function validationRuleIndex(sheet: WorksheetModel): RuleIntervalIndex<DataValidationRule> {
  const cached = validationIndexCache.get(sheet);
  const signature = JSON.stringify(sheet.dataValidations);
  if (cached && validationIndexSignatures.get(sheet) === signature) return cached;
  const index = new RuleIntervalIndex<DataValidationRule>();
  for (const rule of sheet.dataValidations) for (const range of rule.ranges) index.add(range, rule);
  validationIndexCache.set(sheet, index);
  validationIndexSignatures.set(sheet, signature);
  return index;
}

export function resolveValidationRule(sheet: WorksheetModel, row: number, column: number): DataValidationRule | undefined {
  const matches = validationRuleIndex(sheet).query(sheet.id, row, column);
  if (matches.length <= 1) return matches[0];
  const unique = new Map(matches.map((rule) => [rule.id, rule]));
  if (unique.size !== 1) throw new Error(`RULE_OWNER_AMBIGUOUS: multiple data-validation rules own ${sheet.id}!${row}:${column}`);
  return matches[0];
}
