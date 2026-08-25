import assert from 'node:assert/strict';
import test from 'node:test';
import type { MergeSpan } from '@react-sheets/core-model';
import { createMergeSpatialIndex } from './merge-spatial-index';

function merge(startRow: number, endRow: number, startColumn: number, endColumn: number): MergeSpan {
  return { range: { sheetId: 'sheet-1', startRow, endRow, startColumn, endColumn }, anchor: { row: startRow, column: startColumn } };
}

test('finds blank merged cells without scanning unrelated offscreen merge ranges', () => {
  const merges = Array.from({ length: 20_000 }, (_, index) => merge(10_000 + index, 10_000 + index, 50, 51));
  merges.push(merge(1, 2, 1, 2));
  const findMerge = createMergeSpatialIndex(merges);

  assert.equal(findMerge(0, 0), undefined);
  assert.equal(findMerge(2, 2), merges.at(-1));
});

test('preserves first-match behavior for overlapping imported merge ranges', () => {
  const first = merge(0, 5, 0, 5);
  const second = merge(2, 3, 2, 3);
  const findMerge = createMergeSpatialIndex([first, second]);

  assert.equal(findMerge(2, 2), first);
});
