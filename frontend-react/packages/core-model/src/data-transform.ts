import type { RangeRef, Row } from './index';

/** Pure sort intent. The kernel computes and commits row/reference permutations. */
export interface RowPermutationPlan {
  readonly range: RangeRef;
  readonly sourceRows: readonly Row[];
}

function normalizeRange(range: RangeRef): RangeRef {
  if (![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isInteger)) {
    throw new Error('Row permutation range must contain integer coordinates');
  }
  if (range.startRow < 0 || range.startColumn < 0) throw new Error('Row permutation range is outside worksheet bounds');
  return {
    ...range,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

/** Build the immutable command payload consumed by the kernel sort mutation. */
export function createRowPermutationPlan(range: RangeRef, sourceRows: readonly Row[]): RowPermutationPlan {
  const normalized = normalizeRange(range);
  const expectedCount = normalized.endRow - normalized.startRow + 1;
  if (sourceRows.length !== expectedCount) throw new Error('Row permutation length does not match the range');
  const expected = new Set<number>();
  for (let row = normalized.startRow; row <= normalized.endRow; row += 1) expected.add(row);
  const seen = new Set<number>();
  for (const sourceRow of sourceRows) {
    if (!Number.isInteger(sourceRow) || !expected.has(sourceRow) || seen.has(sourceRow)) {
      throw new Error('Row permutation must contain every selected row exactly once');
    }
    seen.add(sourceRow);
  }
  return Object.freeze({ range: Object.freeze(normalized), sourceRows: Object.freeze([...sourceRows]) });
}
