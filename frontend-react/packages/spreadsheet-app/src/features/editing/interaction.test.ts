import test from 'node:test';
import assert from 'node:assert/strict';
import { createInitialSelection } from '../../selection-service';
import { applyHeaderSelection } from '../../header-interaction-domain';
import { planRangeDrag, rangeDragMode } from './index';

test('header selection has one row/column ownership path with additive and extend modes', () => {
  const initial = createInitialSelection('sheet-1');
  const bounds = { rowCount: 10, columnCount: 8 };
  const column = applyHeaderSelection(initial, { kind: 'column', index: 2 }, 'sheet-1', bounds, { additive: false, extend: false });
  assert.deepEqual(column.ranges, [{ sheetId: 'sheet-1', startRow: 0, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const extended = applyHeaderSelection(column, { kind: 'column', index: 5 }, 'sheet-1', bounds, { additive: false, extend: true });
  assert.deepEqual(extended.ranges, [{ sheetId: 'sheet-1', startRow: 0, endRow: 9, startColumn: 2, endColumn: 5 }]);
  const added = applyHeaderSelection(extended, { kind: 'column', index: 7 }, 'sheet-1', bounds, { additive: true, extend: false });
  assert.equal(added.ranges.length, 2);
  assert.equal(added.selectionKind, 'columns');
});

test('range drag planner resolves copy/move and insert modes before execution', () => {
  assert.equal(rangeDragMode({ copy: false, insert: false }), 'move-replace');
  assert.equal(rangeDragMode({ copy: true, insert: true }), 'copy-insert');
  const plan = planRangeDrag({ sheetId: 'sheet-1', startRow: 1, endRow: 2, startColumn: 1, endColumn: 3 }, { row: 5, column: 1 }, 'move-replace', { rowCount: 10, columnCount: 8 });
  assert.deepEqual(plan.targetRange, { sheetId: 'sheet-1', startRow: 5, endRow: 6, startColumn: 1, endColumn: 3 });
  assert.throws(() => planRangeDrag(plan.sourceRange, { row: 9, column: 1 }, plan.mode, { rowCount: 10, columnCount: 8 }), /OUT_OF_BOUNDS/);
  assert.throws(() => planRangeDrag(plan.sourceRange, { row: 1, column: 1 }, plan.mode, { rowCount: 10, columnCount: 8 }), /NOOP/);
});
