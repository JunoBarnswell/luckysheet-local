import assert from 'node:assert/strict';
import test from 'node:test';
import { applyHeaderSelection, headerContextMenuCatalog, headerTargetSelected, selectedHeaderIndices } from './header-interaction-domain';
import { createInitialSelection } from './selection-service';

const bounds = { rowCount: 100, columnCount: 26 };

test('header selection supports Shift extension and Ctrl/Meta non-adjacent ranges without collapsing existing ranges', () => {
  const initial = createInitialSelection('sheet-1');
  const first = applyHeaderSelection(initial, { kind: 'column', index: 2 }, 'sheet-1', bounds, { additive: false, extend: false });
  const extended = applyHeaderSelection(first, { kind: 'column', index: 5 }, 'sheet-1', bounds, { additive: false, extend: true });
  assert.deepEqual(extended.ranges[0], { sheetId: 'sheet-1', startRow: 0, endRow: 99, startColumn: 2, endColumn: 5 });
  const added = applyHeaderSelection(extended, { kind: 'column', index: 9 }, 'sheet-1', bounds, { additive: true, extend: false });
  assert.equal(added.ranges.length, 2);
  assert.deepEqual(selectedHeaderIndices(added, 'column', bounds), [2, 3, 4, 5, 9]);
  assert.equal(headerTargetSelected(added, { kind: 'column', index: 4 }, bounds), true);
  assert.equal(headerTargetSelected(added, { kind: 'column', index: 8 }, bounds), false);
});

test('row and column context catalogs are structurally identical and omit non-Excel AutoFit menu shortcuts', () => {
  const columns = headerContextMenuCatalog('column');
  const rows = headerContextMenuCatalog('row');
  assert.deepEqual(columns.map((entry) => entry.action), rows.map((entry) => entry.action));
  assert.equal(columns.some((entry) => /AutoFit/i.test(entry.label)), false);
  assert.equal(rows.some((entry) => /AutoFit/i.test(entry.label)), false);
  assert.ok(columns.some((entry) => entry.label === 'Insert Columns'));
  assert.ok(rows.some((entry) => entry.label === 'Insert Rows'));
  assert.ok(columns.some((entry) => entry.label === 'Format Cells…'));
});
