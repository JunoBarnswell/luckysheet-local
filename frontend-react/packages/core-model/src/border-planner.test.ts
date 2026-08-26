import test from 'node:test';
import assert from 'node:assert/strict';
import { planBorderChange, type BorderLine, type RangeRef } from './index';

const range: RangeRef = { sheetId: 'sheet-1', startRow: 1, endRow: 3, startColumn: 2, endColumn: 4 };
const line: BorderLine = { style: 'thin', color: '#334155' };

test('border planner emits only the requested side for directional placements', () => {
  assert.deepEqual(planBorderChange(range, 'bottom', line, { rowCount: 10, columnCount: 10 }).cells[0]?.sides, { bottom: line });
  assert.deepEqual(planBorderChange(range, 'top', line, { rowCount: 10, columnCount: 10 }).cells[0]?.sides, { top: line });
  assert.deepEqual(planBorderChange(range, 'left', line, { rowCount: 10, columnCount: 10 }).cells[0]?.sides, { left: line });
  assert.deepEqual(planBorderChange(range, 'right', line, { rowCount: 10, columnCount: 10 }).cells[0]?.sides, { right: line });
});

test('outside and thick-outside have no internal edges', () => {
  const outside = planBorderChange(range, 'outside', line, { rowCount: 10, columnCount: 10 });
  const center = outside.cells.find((cell) => cell.row === 2 && cell.column === 3);
  const corner = outside.cells.find((cell) => cell.row === 1 && cell.column === 2);
  assert.deepEqual(center?.sides, {});
  assert.deepEqual(corner?.sides, { top: line, left: line });
  assert.equal(planBorderChange(range, 'thick-outside', line, { rowCount: 10, columnCount: 10 }).cells[0]?.sides.top?.style, 'thick');
});

test('all and inside placements are deterministic and side-aware', () => {
  const all = planBorderChange(range, 'all', line, { rowCount: 10, columnCount: 10 });
  assert.equal(all.cells.length, 9);
  assert.deepEqual(all.cells[4]?.sides, { top: line, right: line, bottom: line, left: line });
  const horizontal = planBorderChange(range, 'inside-horizontal', line, { rowCount: 10, columnCount: 10 });
  assert.deepEqual(horizontal.cells.find((cell) => cell.row === 1)?.sides, {});
  assert.deepEqual(horizontal.cells.find((cell) => cell.row === 2)?.sides, { top: line });
  const vertical = planBorderChange(range, 'inside-vertical', line, { rowCount: 10, columnCount: 10 });
  assert.deepEqual(vertical.cells.find((cell) => cell.column === 2)?.sides, {});
  assert.deepEqual(vertical.cells.find((cell) => cell.column === 3)?.sides, { left: line });
});

test('none clears only the border shape and invalid ranges fail closed', () => {
  assert.deepEqual(planBorderChange(range, 'none', undefined, { rowCount: 10, columnCount: 10 }).cells[0]?.sides, {});
  assert.throws(() => planBorderChange({ ...range, endRow: 10 }, 'all', line, { rowCount: 10, columnCount: 10 }), /outside worksheet bounds/);
  assert.throws(() => planBorderChange(range, 'all', undefined, { rowCount: 10, columnCount: 10 }), /Border line/);
});
