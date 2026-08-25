import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { buildPivotSlicerItems } from './PivotSlicer';
import { pivotMemberKey } from '@react-sheets/core-model';

describe('PivotSlicer typed member items', () => {
  it('keeps colliding labels independently addressable and formats blanks consistently', () => {
    const items = buildPivotSlicerItems([1, '1', true, 'true', false, 'false', null, '']);
    assert.equal(items.length, 7);
    assert.deepEqual(items.map((item) => item.label), ['1', '1', 'true', 'true', 'false', 'false', '(blank)']);
    assert.equal(new Set(items.map((item) => pivotMemberKey(item.key))).size, 7);
    assert.equal(items[0]?.key.type, 'number');
    assert.equal(items[1]?.key.type, 'text');
    assert.equal(items[2]?.key.type, 'boolean');
    assert.equal(items[3]?.key.type, 'text');
    assert.equal(items[6]?.key.type, 'blank');
  });

  it('deduplicates only canonical blank identity, never display labels', () => {
    const items = buildPivotSlicerItems([1, '1', 1, '1', null, '']);
    assert.deepEqual(items.map((item) => item.key.type), ['number', 'text', 'blank']);
  });
});
