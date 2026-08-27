import type { CellData } from '@react-sheets/core-model';
export interface ColumnAutocompleteSource {
  key: string;
  revision: number;
  entries: Iterable<{ row: number; cell: CellData }>;
  excludeRow: number;
  cultureId: string;
}

interface NormalizedValue {
  key: string;
  value: string;
}

interface ColumnAutocompleteEntry {
  revision: number;
  values: readonly NormalizedValue[];
  locale: string;
}

function compare(left: NormalizedValue, right: NormalizedValue): number {
  return left.key < right.key ? -1 : left.key > right.key ? 1 : 0;
}

function normalize(value: string, locale: string): string {
  return value.normalize('NFKC').toLocaleLowerCase(locale);
}

function mergeUnique(left: readonly NormalizedValue[], right: readonly NormalizedValue[]): NormalizedValue[] {
  const merged: NormalizedValue[] = [];
  let leftIndex = 0;
  let rightIndex = 0;
  while (leftIndex < left.length || rightIndex < right.length) {
    const leftValue = left[leftIndex];
    const rightValue = right[rightIndex];
    let value: NormalizedValue;
    if (rightValue === undefined || leftValue !== undefined && compare(leftValue, rightValue) <= 0) {
      value = leftValue!;
      leftIndex += 1;
      if (rightValue !== undefined && value.key === rightValue.key) rightIndex += 1;
    } else {
      value = rightValue;
      rightIndex += 1;
    }
    if (merged.at(-1)?.key !== value.key) merged.push(value);
  }
  return merged;
}

function lowerBound(values: readonly NormalizedValue[], prefix: string): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (values[middle]!.key < prefix) low = middle + 1;
    else high = middle;
  }
  return low;
}

export class ColumnValueAutocompleteIndex {
  private static readonly MAX_CACHED_COLUMNS = 8;
  private readonly cache = new Map<string, ColumnAutocompleteEntry>();

  has(source: Pick<ColumnAutocompleteSource, 'key' | 'revision'>): boolean {
    return this.cache.get(source.key)?.revision === source.revision;
  }

  async rebuild(source: ColumnAutocompleteSource, signal: AbortSignal): Promise<void> {
    const locale = source.cultureId === 'invariant' ? 'en-US' : source.cultureId;
    const chunks: NormalizedValue[][] = [];
    let chunk: NormalizedValue[] = [];
    for (const entry of source.entries) {
      if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
      if (entry.row === source.excludeRow || entry.cell.formula || typeof entry.cell.value !== 'string' || entry.cell.value.length === 0) continue;
      chunk.push({ key: normalize(entry.cell.value, locale), value: entry.cell.value });
      if (chunk.length >= 4096) {
        chunk.sort(compare);
        chunks.push(chunk);
        chunk = [];
        await Promise.resolve();
      }
    }
    if (chunk.length > 0) {
      chunk.sort(compare);
      chunks.push(chunk);
    }
    while (chunks.length > 1) {
      const merged: NormalizedValue[][] = [];
      for (let index = 0; index < chunks.length; index += 2) {
        if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
        merged.push(chunks[index + 1] ? mergeUnique(chunks[index]!, chunks[index + 1]!) : chunks[index]!);
        await Promise.resolve();
      }
      chunks.splice(0, chunks.length, ...merged);
    }
    if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
    this.cache.delete(source.key);
    this.cache.set(source.key, { revision: source.revision, values: chunks[0] ? mergeUnique(chunks[0], []) : [], locale });
    while (this.cache.size > ColumnValueAutocompleteIndex.MAX_CACHED_COLUMNS) {
      const oldest = this.cache.keys().next().value;
      if (oldest === undefined) break;
      this.cache.delete(oldest);
    }
  }

  query(key: string, revision: number, prefix: string): string | null {
    if (!prefix) return null;
    const entry = this.cache.get(key);
    if (!entry || entry.revision !== revision) return null;
    const normalizedPrefix = normalize(prefix, entry.locale);
    const index = lowerBound(entry.values, normalizedPrefix);
    for (let cursor = index; cursor < entry.values.length; cursor += 1) {
      const candidate = entry.values[cursor]!;
      if (!candidate.key.startsWith(normalizedPrefix)) return null;
      if (candidate.value.length > prefix.length) return candidate.value;
    }
    return null;
  }

  invalidate(key: string): void {
    this.cache.delete(key);
  }
}
