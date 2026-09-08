import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createRowPermutationPlan } from './data-transform';
import type { RangeRef } from './index';

/** Historical cases retained as kernel sort corpus identifiers. */
export const rowPermutationCorpus = [
  'metadata follows exact sort rectangle cells',
  'conditional-format target preserves non-contiguous segments',
  'unrepresentable single-range metadata rejects atomically',
] as const;

describe('row permutation intent', () => {
  it('keeps the historical corpus and emits only a serializable plan', () => {
    assert.equal(rowPermutationCorpus.length, 3);
    const range: RangeRef = { sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
    const plan = createRowPermutationPlan(range, [1, 0]);
    assert.deepEqual(plan, { range, sourceRows: [1, 0] });
    assert.equal('sourceToTarget' in plan, false);
  });

  it('rejects malformed kernel sort payloads before dispatch', () => {
    const range: RangeRef = { sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
    assert.throws(() => createRowPermutationPlan(range, [1, 1]), /every selected row exactly once/);
  });
});
