import test from 'node:test';
import assert from 'node:assert/strict';
import { planSheetExtentGrowth, resolveAutoScrollExtentGrowth } from './sheet-extent-growth';

test('extent growth emits one canonical request until the model reaches the pending target', () => {
  const current = { sheetId: 'sheet-1', rowCount: 4_000, columnCount: 100 };
  const first = planSheetExtentGrowth(current, current, { rows: true });
  assert.deepEqual(first, { sheetId: 'sheet-1', rowCount: 5_000, columnCount: 100 });
  assert.equal(planSheetExtentGrowth(current, first!, { rows: true }), null);
  assert.deepEqual(planSheetExtentGrowth(
    { ...current, rowCount: 5_000 },
    first!,
    { rows: true },
  ), { sheetId: 'sheet-1', rowCount: 6_000, columnCount: 100 });
});

test('selection auto-scroll grows only an axis that is moving at its true extent boundary', () => {
  const base = {
    viewport: { width: 1_000, height: 600, scrollX: 200, scrollY: 400 },
    content: { width: 10_000, height: 20_000 },
    defaultRowHeight: 20,
    defaultColumnWidth: 80,
  };
  assert.deepEqual(resolveAutoScrollExtentGrowth({ ...base, right: true, bottom: true }), { rows: false, columns: false });
  assert.deepEqual(resolveAutoScrollExtentGrowth({
    ...base,
    right: true,
    bottom: true,
    viewport: { ...base.viewport, scrollX: 8_900, scrollY: 19_000 },
  }), { rows: true, columns: true });
});
