import { isFormulaError, type FormulaValue } from '../values';
import { compareWorkbookValues, type WorkbookCollationContext } from '../collation';

export type LookupMatchMode = 0 | -1 | 1 | 2;

function scalar(value: FormulaValue): number | string | boolean | null {
  return isFormulaError(value) || Array.isArray(value) ? null : value;
}

export function lookupCompare(left: FormulaValue, right: FormulaValue, context?: WorkbookCollationContext): number | null {
  const a = scalar(left);
  const b = scalar(right);
  if (a === null || b === null) return null;
  return compareWorkbookValues(a, b, context);
}

function wildcardRegex(pattern: string): RegExp {
  let expression = '^';
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]!;
    if (char === '~' && index + 1 < pattern.length) expression += pattern[++index]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (char === '*') expression += '.*';
    else if (char === '?') expression += '.';
    else expression += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`${expression}$`, 'i');
}

export function findLookupIndex(value: FormulaValue | undefined, vector: readonly FormulaValue[], mode: LookupMatchMode = 0, searchMode = 1, context?: WorkbookCollationContext): number {
  if (vector.length === 0) return -1;
  const order = searchMode === -1 || searchMode === -2 ? [...vector.keys()].reverse() : [...vector.keys()];
  if (mode === 0 || mode === 2) {
    const expected = String(value ?? '');
    return order.find((index) => mode === 2
      ? wildcardRegex(expected).test(String(vector[index] ?? ''))
      : lookupCompare(value ?? null, vector[index]!, context) === 0) ?? -1;
  }
  let best = -1;
  for (const index of order) {
    const comparison = lookupCompare(vector[index]!, value ?? null, context);
    if (comparison === null) continue;
    if (mode === -1 && comparison <= 0 && (best < 0 || (lookupCompare(vector[index]!, vector[best]!, context) ?? 1) > 0)) best = index;
    if (mode === 1 && comparison >= 0 && (best < 0 || (lookupCompare(vector[index]!, vector[best]!, context) ?? -1) < 0)) best = index;
  }
  return best;
}
