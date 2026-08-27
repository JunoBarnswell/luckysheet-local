import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { DEFAULT_PIVOT_FIELD_PANE_LAYOUT, defaultPivotFieldArea, PIVOT_FIELD_AREAS, PIVOT_FIELD_PANE_LAYOUTS, shouldDeferPivotLayoutUpdates } from './pivot-contract';

describe('Pivot Field List pane layout contract', () => {
  it('exposes six presentation modes without changing Pivot area ownership', () => {
    assert.deepEqual([...PIVOT_FIELD_PANE_LAYOUTS], ['stacked', 'side-by-side', 'areas-2x2', 'areas-1x4', 'fields-only', 'areas-only']);
    assert.equal(new Set(PIVOT_FIELD_PANE_LAYOUTS).size, 6);
  });

  it('defaults to a stacked, task-first workspace and keeps one assignment-area order', () => {
    assert.equal(DEFAULT_PIVOT_FIELD_PANE_LAYOUT, 'stacked');
    assert.deepEqual(PIVOT_FIELD_AREAS, ['filters', 'columns', 'rows', 'values']);
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

  it('uses official deferred layout updates for large worksheet and block sources', () => {
    assert.equal(shouldDeferPivotLayoutUpdates({ source: { kind: 'worksheet-range', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4058, startColumn: 0, endColumn: 22 } } }), true);
    assert.equal(shouldDeferPivotLayoutUpdates({ source: { kind: 'worksheet-range', range: { sheetId: 'sheet-1', startRow: 0, endRow: 99, startColumn: 0, endColumn: 9 } } }), false);
    assert.equal(shouldDeferPivotLayoutUpdates({ source: { kind: 'data-source', dataSourceId: 'source-1' } }), true);
  });
});
