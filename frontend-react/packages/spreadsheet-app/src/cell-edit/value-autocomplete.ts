import type { CellData } from '@react-sheets/core-model';

export interface ColumnAutocompleteSource {
  key: string;
  revision: number;
  entries: Iterable<{ row: number; cell: CellData }>;
  excludeRow: number;
  cultureId: string;
}

interface ColumnAutocompleteEntry {
  revision: number;
  values: readonly string[];
  locale: string;
}

function mergeUnique(left: readonly string[], right: readonly string[], compare: (left: string, right: string) => number): string[] {
  const merged: string[] = [];
  let leftIndex = 0;
  let rightIndex = 0;
  while (leftIndex < left.length || rightIndex < right.length) {
    const leftValue = left[leftIndex];
    const rightValue = right[rightIndex];
    let value: string;
    if (rightValue === undefined || leftValue !== undefined && compare(leftValue, rightValue) <= 0) {
      value = leftValue!;
      leftIndex += 1;
      if (rightValue !== undefined && compare(value, rightValue) === 0) rightIndex += 1;
    } else {
      value = rightValue;
      rightIndex += 1;
    }
    if (!merged.at(-1) || compare(merged.at(-1)!, value) !== 0) merged.push(value);
  }
  return merged;
}

function lowerBound(values: readonly string[], prefix: string, compare: (left: string, right: string) => number): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (compare(values[middle]!, prefix) < 0) low = middle + 1;
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
    const compare = (left: string, right: string) => left.localeCompare(right, locale, { sensitivity: 'base' });
    const chunks: string[][] = [];
    let chunk: string[] = [];
    for (const entry of source.entries) {
      if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
      if (entry.row === source.excludeRow || entry.cell.formula || typeof entry.cell.value !== 'string' || entry.cell.value.length === 0) continue;
      chunk.push(entry.cell.value);
      if (chunk.length >= 2048) {
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
      const merged: string[][] = [];
      for (let index = 0; index < chunks.length; index += 2) {
        if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
        merged.push(chunks[index + 1] ? mergeUnique(chunks[index]!, chunks[index + 1]!, compare) : chunks[index]!);
        await Promise.resolve();
      }
      chunks.splice(0, chunks.length, ...merged);
    }
    if (signal.aborted) throw new DOMException('Column autocomplete rebuild was aborted', 'AbortError');
    this.cache.delete(source.key);
    this.cache.set(source.key, { revision: source.revision, values: chunks[0] ? mergeUnique(chunks[0], [], compare) : [], locale });
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
    const compare = (left: string, right: string) => left.localeCompare(right, entry.locale, { sensitivity: 'base' });
    const index = lowerBound(entry.values, prefix, compare);
    const normalizedPrefix = prefix.toLocaleLowerCase(entry.locale);
    for (let cursor = index; cursor < entry.values.length; cursor += 1) {
      const candidate = entry.values[cursor]!;
      if (!candidate.toLocaleLowerCase(entry.locale).startsWith(normalizedPrefix)) return null;
      if (candidate.length > prefix.length) return candidate;
    }
    return null;
  }

  invalidate(key: string): void {
    this.cache.delete(key);
  }
}
