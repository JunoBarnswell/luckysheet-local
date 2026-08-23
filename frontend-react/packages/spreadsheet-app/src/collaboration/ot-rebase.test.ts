import test from 'node:test';
import assert from 'node:assert/strict';
import { classifyMutation } from './operation-types';
import { rebaseMutation } from './ot-rebase';

test('rebaseMutation shifts cell reference after insert-rows', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 's1', [{
    sheetId: 's1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0,
  }]);
  const pending = classifyMutation('cell.set', { row: 9, column: 0 }, 's1', [{
    sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
  }]);

  const { rebased, transformed } = rebaseMutation(pending, committed);
  assert.equal(transformed, true);
  assert.equal(rebased.affectedRanges[0]?.startRow, 10);
  assert.equal((rebased.params as { row: number }).row, 10);
});
