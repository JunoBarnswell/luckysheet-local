import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { defaultPivotFieldArea, PIVOT_FIELD_PANE_LAYOUTS } from './pivot-contract';

describe('Pivot Field List pane layout contract', () => {
  it('exposes six presentation modes without changing Pivot area ownership', () => {
    assert.deepEqual([...PIVOT_FIELD_PANE_LAYOUTS], ['stacked', 'side-by-side', 'areas-2x2', 'areas-1x4', 'fields-only', 'areas-only']);
    assert.equal(new Set(PIVOT_FIELD_PANE_LAYOUTS).size, 6);
  });

  it('keeps the report layout modes separate from Field List pane modes', () => {
    assert.equal(PIVOT_FIELD_PANE_LAYOUTS.includes('compact' as never), false);
    assert.equal(PIVOT_FIELD_PANE_LAYOUTS.includes('outline' as never), false);
    assert.equal(PIVOT_FIELD_PANE_LAYOUTS.includes('tabular' as never), false);
  });

  it('uses the same type-aware default area for keyboard and pointer assignment', () => {
    assert.equal(defaultPivotFieldArea({ dataType: 'number' }), 'values');
    assert.equal(defaultPivotFieldArea({ dataType: 'date' }), 'columns');
    assert.equal(defaultPivotFieldArea({ dataType: 'text' }), 'rows');
  });
});
