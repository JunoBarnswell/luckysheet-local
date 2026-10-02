import { isFormulaError, isReferenceValue, type FormulaValue } from '../values';
import { compareWorkbookValues, type WorkbookCollationContext } from '../collation';

export type LookupMatchMode = 0 | -1 | 1 | 2;

export function lookupCompare(left: FormulaValue, right: FormulaValue, context?: WorkbookCollationContext): number | null {
  if (isFormulaError(left) || isFormulaError(right) || isReferenceValue(left) || isReferenceValue(right) || Array.isArray(left) || Array.isArray(right)) return null;
  return compareWorkbookValues(left, right, context);
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

export interface LookupVector {
  readonly length: number;
  at(index: number): FormulaValue;
}

export function findLookupIndex(value: FormulaValue | undefined, vector: readonly FormulaValue[] | LookupVector, mode: LookupMatchMode = 0, searchMode = 1, context?: WorkbookCollationContext): number {
  if (vector.length === 0) return -1;
  const read = (index: number) => Array.isArray(vector) ? vector[index]! : (vector as LookupVector).at(index);
  if (searchMode === 2 || searchMode === -2) {
    if (mode === 2) return -1;
    const direction = searchMode === 2 ? 1 : -1;
    let low = 0;
    let high = vector.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const comparison = lookupCompare(read(middle), value ?? null, context);
      if (comparison === null) return -1;
      if (comparison * direction < 0) low = middle + 1;
      else high = middle;
    }
    if (low < vector.length && lookupCompare(read(low), value ?? null, context) === 0) return low;
    if (mode === 0) return -1;
    const index = (mode === -1) === (direction === 1) ? low - 1 : low;
    return index >= 0 && index < vector.length ? index : -1;
  }
  const reverse = searchMode === -1;
  const pattern = mode === 2 ? wildcardRegex(String(value ?? '')) : undefined;
  let best = -1;
  for (let index = reverse ? vector.length - 1 : 0; reverse ? index >= 0 : index < vector.length; index += reverse ? -1 : 1) {
    const candidate = read(index);
    if (mode === 2) {
      if (typeof candidate === 'string' && pattern!.test(candidate)) return index;
      continue;
    }
    const comparison = lookupCompare(candidate, value ?? null, context);
    if (comparison === null) continue;
    if (comparison === 0) return index;
    if (mode === -1 && comparison < 0 && (best < 0 || (lookupCompare(candidate, read(best), context) ?? 1) > 0)) best = index;
    if (mode === 1 && comparison > 0 && (best < 0 || (lookupCompare(candidate, read(best), context) ?? -1) < 0)) best = index;
  }
  return best;
}
