import type { CellRange } from './types';

export function normalizeCellRange(range: CellRange): CellRange | null {
  const values = [range.startRow, range.endRow, range.startColumn, range.endColumn];
  if (values.some((value) => !Number.isFinite(value))) return null;

  const rowStart = Math.trunc(Math.min(range.startRow, range.endRow));
  const rowEnd = Math.trunc(Math.max(range.startRow, range.endRow));
  const columnStart = Math.trunc(Math.min(range.startColumn, range.endColumn));
  const columnEnd = Math.trunc(Math.max(range.startColumn, range.endColumn));
  if (rowEnd < 0 || columnEnd < 0) return null;

  return {
    startRow: Math.max(0, rowStart),
    endRow: Math.max(0, rowEnd),
    startColumn: Math.max(0, columnStart),
    endColumn: Math.max(0, columnEnd),
  };
}

function rangesTouch(first: CellRange, second: CellRange): boolean {
  return first.startRow <= second.endRow + 1
    && second.startRow <= first.endRow + 1
    && first.startColumn <= second.endColumn + 1
    && second.startColumn <= first.endColumn + 1;
}

function mergeRange(first: CellRange, second: CellRange): CellRange {
  return {
    startRow: Math.min(first.startRow, second.startRow),
    endRow: Math.max(first.endRow, second.endRow),
    startColumn: Math.min(first.startColumn, second.startColumn),
    endColumn: Math.max(first.endColumn, second.endColumn),
  };
}

export function mergeCellRanges(ranges: readonly CellRange[]): CellRange[] {
  const pending = ranges
    .map(normalizeCellRange)
    .filter((range): range is CellRange => range !== null)
    .sort((first, second) => first.startRow - second.startRow || first.startColumn - second.startColumn);

  const merged: CellRange[] = [];
  for (const range of pending) {
    let candidate = range;
    let changed = true;
    while (changed) {
      changed = false;
      for (let index = merged.length - 1; index >= 0; index -= 1) {
        const existing = merged[index];
        if (!existing || !rangesTouch(existing, candidate)) continue;
        candidate = mergeRange(existing, candidate);
        merged.splice(index, 1);
        changed = true;
      }
    }
    merged.push(candidate);
  }

  return merged.sort((first, second) => first.startRow - second.startRow || first.startColumn - second.startColumn);
}

export function cellRangesIntersect(first: CellRange, second: CellRange): boolean {
  return first.startRow <= second.endRow
    && second.startRow <= first.endRow
    && first.startColumn <= second.endColumn
    && second.startColumn <= first.endColumn;
}

export function intersectCellRange(first: CellRange, second: CellRange): CellRange | null {
  if (!cellRangesIntersect(first, second)) return null;
  return {
    startRow: Math.max(first.startRow, second.startRow),
    endRow: Math.min(first.endRow, second.endRow),
    startColumn: Math.max(first.startColumn, second.startColumn),
    endColumn: Math.min(first.endColumn, second.endColumn),
  };
}

export class DirtyRangeSet {
  private ranges: CellRange[] = [];

  add(range: CellRange): void {
    const normalized = normalizeCellRange(range);
    if (!normalized) return;
    this.ranges = mergeCellRanges([...this.ranges, normalized]);
  }

  addMany(ranges: readonly CellRange[]): void {
    this.ranges = mergeCellRanges([...this.ranges, ...ranges]);
  }

  clear(): void {
    this.ranges = [];
  }

  get size(): number {
    return this.ranges.length;
  }

  toArray(): CellRange[] {
    return this.ranges.map((range) => ({ ...range }));
  }

  consume(): CellRange[] {
    const result = this.toArray();
    this.clear();
    return result;
  }
}
