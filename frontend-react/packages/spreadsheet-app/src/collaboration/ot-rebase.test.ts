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

test('fails closed when a structural revision would move an unclassified mutation', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 's1', [{
    sheetId: 's1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0,
  }]);
  const pending = classifyMutation('custom.known-but-unclassified', { row: 9 }, 's1', [{
    sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
  }]);
  assert.throws(() => rebaseMutation(pending, committed), /Cannot rebase unknown mutation/);
});

test('fails closed when a committed mutation has no registered transform', () => {
  const committed = classifyMutation('extension.structural.change', { row: 2 }, 's1', []);
  const pending = classifyMutation('cell.set', { row: 9, column: 0 }, 's1', [{
    sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
  }]);

  assert.throws(() => rebaseMutation(pending, committed), /committed extension\.structural\.change has no registered structural transform/);
});

test('classifies worksheet identity changes and fails closed without a canonical patch', () => {
  for (const mutationId of ['sheet.add', 'sheet.remove', 'sheet.rename', 'sheet.duplicated', 'sheet.restore', 'sheet.reordered']) {
    assert.equal(classifyMutation(mutationId, {}, 's1', []).kind, 'sheet-identity');
  }
  const committed = classifyMutation('sheet.rename', { sheetId: 's1', name: 'Renamed' }, 's1', []);
  const pending = classifyMutation('cell.set', { row: 0, column: 0, formula: '=Sheet1!A1' }, 's1', [{
    sheetId: 's1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0,
  }]);

  assert.throws(() => rebaseMutation(pending, committed), /no canonical structural patch/);
  const rowInsert = classifyMutation('row.insert', { at: 2, count: 1 }, 's1', []);
  const pendingRename = classifyMutation('sheet.rename', { sheetId: 's1', name: 'Local name' }, 's1', []);
  assert.throws(() => rebaseMutation(pendingRename, rowInsert), /pending sheet\.rename has no canonical structural patch/);
});

test('does not infer pending move, sort, or table-resize transforms from coordinate field names', () => {
  const committed = classifyMutation('row.insert', { at: 2, count: 1 }, 's1', []);
  const unsupported = [
    ['range.move', { sourceRange: { sheetId: 's1', startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 }, targetOrigin: { row: 8, column: 2 } }, 'move-range'],
    ['sort.apply', { range: { sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 }, sourceRows: [5, 4, 6, 7, 8] }, 'sort'],
    ['table.resize', { tableId: 'table-1', range: { sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 } }, 'table-resize'],
  ] as const;
  for (const [mutationId, params, kind] of unsupported) {
    const pending = classifyMutation(mutationId, params, 's1', [{ sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 }]);
    assert.equal(pending.kind, kind);
    assert.throws(() => rebaseMutation(pending, committed), new RegExp(`pending ${mutationId.replace('.', '\\.')} has no canonical structural patch`));
  }
});

test('rebases explicitly qualified formulas using canonical worksheet identities', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 'source-id', []);
  const pending = classifyMutation('cell.set', { row: 0, column: 0, formula: '=Source!A10' }, 'owner-id', [{
    sheetId: 'owner-id', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0,
  }]);

  assert.throws(() => rebaseMutation(pending, committed), /worksheet identity is missing/);

  const { rebased } = rebaseMutation(pending, committed, {
    sheetOrder: [
      { id: 'source-id', name: 'Source' },
      { id: 'owner-id', name: 'Owner' },
    ],
  });

  assert.equal((rebased.params as { formula: string }).formula, '=Source!A11');
});

test('rebases formula-rule parameters from their formula anchor sheet', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 'source-id', []);
  const pending = classifyMutation('range.paste', {
    transfer: 'copy',
    clearSource: false,
    snapshot: {
      cells: [],
      validations: [{
        sheetId: 'owner-id',
        formulaAnchor: { sheetId: 'source-id', row: 0, column: 0 },
        operator: 'formula',
        value1: 'A10',
      }],
    },
  }, 'owner-id', []);

  const { rebased } = rebaseMutation(pending, committed, {
    sheetOrder: [
      { id: 'source-id', name: 'Source' },
      { id: 'owner-id', name: 'Owner' },
    ],
  });

  const validations = (rebased.params as { snapshot: { validations: Array<{ value1: string }> } }).snapshot.validations;
  assert.equal(validations[0]?.value1, 'A11');
});

test('classifies cut paste as structural and fails closed across a committed cut', () => {
  const committed = classifyMutation('range.paste', { transfer: 'move', clearSource: true }, 's1', []);
  const pending = classifyMutation('cell.set', { row: 9, column: 0 }, 's1', [{
    sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
  }]);

  assert.equal(committed.kind, 'move-range');
  assert.throws(() => rebaseMutation(pending, committed), /STRUCTURAL_REBASE_CONFLICT/);
});

test('does not classify malformed paste payloads as plain cell writes', () => {
  assert.equal(classifyMutation('range.paste', { transfer: 'copy' }, 's1', []).kind, 'unknown');
  assert.equal(classifyMutation('range.paste', { transfer: 'copy', clearSource: false }, 's1', []).kind, 'cell-value');
  assert.equal(classifyMutation('range.paste', {
    transfer: 'copy', clearSource: false, sourceRange: undefined, sourceSnapshot: undefined,
  }, 's1', []).kind, 'cell-value');
});

test('rebases every absolute coordinate in a pending paste snapshot', () => {
  const pending = classifyMutation('range.paste', {
    transfer: 'move',
    clearSource: true,
    sourceRange: { sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
    targetOrigin: { row: 9, column: 2 },
    snapshot: {
      clearRanges: [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }],
      cells: [{ row: 9, column: 2, value: { value: 'moved' } }],
      notes: [{ key: '9:2' }],
      hyperlinks: [{ key: '9:2' }],
      commentCells: ['9:2'],
      comments: [{ sheetId: 's1', row: 9, column: 2 }],
      columnWidths: [{ column: 2, widthPx: 80 }],
      validations: [{ formulaAnchor: { sheetId: 's1', row: 9, column: 2 }, ranges: [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }] }],
    },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const rowInsert = classifyMutation('row.insert', { at: 5 }, 's1', []);
  const columnInsert = classifyMutation('column.insert', { at: 0 }, 's1', []);

  const afterRows = rebaseMutation(pending, rowInsert).rebased;
  const { rebased } = rebaseMutation(afterRows, columnInsert);
  const params = rebased.params as {
    sourceRange: { startRow: number; startColumn: number };
    targetOrigin: { row: number; column: number };
    snapshot: {
      cells: Array<{ row: number; column: number }>;
      notes: Array<{ key: string }>;
      hyperlinks: Array<{ key: string }>;
      commentCells: string[];
      comments: Array<{ row: number; column: number }>;
      columnWidths: Array<{ column: number }>;
      validations: Array<{ formulaAnchor: { row: number; column: number } }>;
    };
  };

  assert.deepEqual(params.targetOrigin, { row: 10, column: 3 });
  assert.deepEqual(params.sourceRange, { sheetId: 's1', startRow: 10, endRow: 10, startColumn: 3, endColumn: 3 });
  assert.deepEqual(params.snapshot.cells[0], { row: 10, column: 3, value: { value: 'moved' } });
  assert.equal(params.snapshot.notes[0]?.key, '10:3');
  assert.equal(params.snapshot.hyperlinks[0]?.key, '10:3');
  assert.equal(params.snapshot.commentCells[0], '10:3');
  assert.deepEqual(params.snapshot.comments[0], { sheetId: 's1', row: 10, column: 3 });
  assert.equal(params.snapshot.columnWidths[0]?.column, 3);
  assert.deepEqual(params.snapshot.validations[0]?.formulaAnchor, { sheetId: 's1', row: 10, column: 3 });
});

test('rebases cross-sheet source snapshots in their own coordinate space', () => {
  const pending = classifyMutation('range.paste', {
    transfer: 'move',
    clearSource: true,
    sourceRange: { sheetId: 'source', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
    targetOrigin: { row: 9, column: 2 },
    snapshot: { cells: [{ row: 9, column: 2 }] },
    sourceSnapshot: {
      cells: [{ row: 9, column: 2 }],
      notes: [{ key: '9:2' }],
      comments: [{ sheetId: 'source', row: 9, column: 2 }],
    },
  }, 'target', [
    { sheetId: 'target', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
    { sheetId: 'source', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
  ]);
  const committed = classifyMutation('row.insert', { at: 5 }, 'source', []);
  const { rebased } = rebaseMutation(pending, committed);
  const sourceSnapshot = (rebased.params as { sourceSnapshot: { cells: Array<{ row: number; column: number }>; notes: Array<{ key: string }>; comments: Array<{ row: number; column: number }> } }).sourceSnapshot;

  assert.deepEqual(sourceSnapshot.cells[0], { row: 10, column: 2 });
  assert.equal(sourceSnapshot.notes[0]?.key, '10:2');
  assert.deepEqual(sourceSnapshot.comments[0], { sheetId: 'source', row: 10, column: 2 });
});
