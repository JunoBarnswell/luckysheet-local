import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { buildPivotSlicerItems } from './PivotSlicer';
import { applyPivotManualMemberDelta, pivotManualMemberSelected } from './pivot-member-filter';
import { createPivotMemberKey, pivotMemberKey } from '@react-sheets/core-model';

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

  it('represents all to one unchecked member as a compact exclusion delta', () => {
    const state = applyPivotManualMemberDelta({ mode: 'all', memberKeys: [] }, [createPivotMemberKey('member-1')], false);
    assert.deepEqual(state, { mode: 'exclude', memberKeys: [createPivotMemberKey('member-1')] });
    assert.equal(pivotManualMemberSelected(state, createPivotMemberKey('member-10001')), true);
    assert.equal(pivotManualMemberSelected(state, createPivotMemberKey('member-1')), false);
  });

  it('updates include and exclude deltas without rebuilding unseen members', () => {
    const unseen = createPivotMemberKey('member-10001');
    const visible = createPivotMemberKey('member-2');
    const included = applyPivotManualMemberDelta({ mode: 'include', memberKeys: [unseen, visible] }, [visible], false);
    assert.deepEqual(included.memberKeys, [unseen]);
    assert.equal(pivotManualMemberSelected(included, unseen), true);
    const excluded = applyPivotManualMemberDelta({ mode: 'exclude', memberKeys: [unseen] }, [visible], false);
    assert.deepEqual(excluded.memberKeys, [unseen, visible]);
    assert.equal(pivotManualMemberSelected(excluded, unseen), false);
  });

  it('keeps typed collisions and group-all members distinct', () => {
    const numeric = createPivotMemberKey(1);
    const text = createPivotMemberKey('1');
    const state = applyPivotManualMemberDelta({ mode: 'all', memberKeys: [] }, [numeric], false);
    assert.equal(pivotManualMemberSelected(state, text), true);
    assert.equal(pivotManualMemberSelected(state, numeric), false);
    const groupAll = applyPivotManualMemberDelta({ mode: 'all', memberKeys: [] }, [numeric, text], false);
    assert.equal(groupAll.memberKeys.length, 2);
    assert.equal(pivotManualMemberSelected(groupAll, createPivotMemberKey(true)), true);
  });
});
