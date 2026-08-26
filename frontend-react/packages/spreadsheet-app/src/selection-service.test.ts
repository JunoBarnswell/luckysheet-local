import assert from 'node:assert/strict';
import test from 'node:test';
import { moveSelection, selectionFromGesture } from './selection-interaction-machine';
import { resolveSelectionTarget } from './selection-target-resolver';
import { SelectionService, createInitialSelection } from './selection-service';

function createService() {
  let activeSheetId = 'sheet-1';
  return new SelectionService(
    'unit-1',
    () => activeSheetId,
    () => ({ rowCount: 1000, columnCount: 26 }),
    createInitialSelection(activeSheetId),
  );
}

test('canonical state keeps release cell active for a forward drag', () => {
  const service = createService();
  service.applyState({
    ranges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }],
    primaryRangeIndex: 0,
    activeCell: { row: 8, column: 2 },
    anchorCell: { row: 1, column: 0 },
    selectionKind: 'cells',
    mode: 'normal',
  });

  assert.deepEqual(service.getState(), {
    ranges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }],
    primaryRangeIndex: 0,
    activeCell: { row: 8, column: 2 },
    anchorCell: { row: 1, column: 0 },
    selectionKind: 'cells',
    mode: 'normal',
  });
  assert.equal(service.activeCell, 'C9');
});

test('reverse drag keeps the release cell active and normalizes the range', () => {
  const service = createService();
  service.applyState({
    ranges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }],
    primaryRangeIndex: 0,
    activeCell: { row: 1, column: 0 },
    anchorCell: { row: 8, column: 2 },
    selectionKind: 'cells',
    mode: 'normal',
  });

  assert.equal(service.activeCell, 'A2');
  assert.deepEqual(service.getState().anchorCell, { row: 8, column: 2 });
  assert.deepEqual(service.primaryRangeOrDefault(), { sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 });
});

test('extend and add transitions update all canonical fields together', () => {
  const service = createService();
  service.selectCellAt(1, 0);
  service.selectRange({ startRow: 8, endRow: 8, startColumn: 2, endColumn: 2 }, 'extend');
  assert.deepEqual(service.getState(), {
    ranges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }],
    primaryRangeIndex: 0,
    activeCell: { row: 8, column: 2 },
    anchorCell: { row: 1, column: 0 },
    selectionKind: 'cells',
    mode: 'normal',
  });

  service.selectRange({ startRow: 10, endRow: 11, startColumn: 3, endColumn: 4 }, 'add');
  const state = service.getState();
  assert.equal(state.primaryRangeIndex, 1);
  assert.deepEqual(state.activeCell, { row: 10, column: 3 });
  assert.deepEqual(state.anchorCell, { row: 10, column: 3 });
  assert.equal(state.ranges.length, 2);
});

test('selection snapshots and states are defensive copies', () => {
  const service = createService();
  const state = service.getState();
  state.ranges[0]!.startRow = 9;
  state.activeCell.row = 9;
  const snapshot = service.getSnapshot();
  snapshot.ranges[0]!.startColumn = 9;
  snapshot.anchorCell.column = 9;

  assert.deepEqual(service.getState(), {
    ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    primaryRangeIndex: 0,
    activeCell: { row: 0, column: 0 },
    anchorCell: { row: 0, column: 0 },
    selectionKind: 'cells',
    mode: 'normal',
  });
});

test('row, column, and all selection use explicit bounds instead of sentinels', () => {
  const service = createService();
  service.selectRow(4, 7);
  assert.deepEqual(service.primaryRangeOrDefault(), { sheetId: 'sheet-1', startRow: 4, endRow: 4, startColumn: 0, endColumn: 6 });
  service.selectColumn(3, 9);
  assert.deepEqual(service.primaryRangeOrDefault(), { sheetId: 'sheet-1', startRow: 0, endRow: 8, startColumn: 3, endColumn: 3 });
  service.selectAll(9, 7);
  assert.deepEqual(service.primaryRangeOrDefault(), { sheetId: 'sheet-1', startRow: 0, endRow: 8, startColumn: 0, endColumn: 6 });
});

test('selection machine preserves release target and fixed drag anchor', () => {
  const initial = createInitialSelection('sheet-1');
  const forward = selectionFromGesture(initial, {
    origin: { row: 1, column: 0 },
    target: { row: 8, column: 2 },
    expandedRange: { sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 },
  }, 'sheet-1');
  const reverse = selectionFromGesture(initial, {
    origin: { row: 8, column: 2 },
    target: { row: 1, column: 0 },
    expandedRange: { sheetId: 'sheet-1', startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 },
  }, 'sheet-1');
  assert.deepEqual(forward.activeCell, { row: 8, column: 2 });
  assert.deepEqual(forward.anchorCell, { row: 1, column: 0 });
  assert.deepEqual(reverse.activeCell, { row: 1, column: 0 });
  assert.deepEqual(reverse.anchorCell, { row: 8, column: 2 });
  assert.deepEqual(reverse.ranges, forward.ranges);
});

test('selection machine rejects hidden keyboard targets by moving to the next visible coordinate', () => {
  const initial = createInitialSelection('sheet-1');
  const next = moveSelection({ ...initial, activeCell: { row: 1, column: 0 } }, 1, 0, false, { rowCount: 10, columnCount: 4, hiddenRows: [2] });
  assert.deepEqual(next.activeCell, { row: 3, column: 0 });
});

test('selection target resolver returns merge anchor and exact merge range', () => {
  const target = resolveSelectionTarget({
    rowCount: 20,
    columnCount: 10,
    merges: [{ range: { sheetId: 'sheet-1', startRow: 2, endRow: 4, startColumn: 3, endColumn: 5 }, anchor: { row: 2, column: 3 } }],
  }, { row: 3, column: 4 }, 'cells', 'sheet-1');
  assert.deepEqual(target.cell, { row: 2, column: 3 });
  assert.deepEqual(target.range, { sheetId: 'sheet-1', startRow: 2, endRow: 4, startColumn: 3, endColumn: 5 });
});
