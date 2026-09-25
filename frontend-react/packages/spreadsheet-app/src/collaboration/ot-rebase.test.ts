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

test('rebases cell.set write-authority target with its canonical cell address', () => {
  const formula = { value: 1, formula: '=A10', formulaMetadata: { kind: 'normal', sourceFormula: '=A10' } };
  const pending = classifyMutation('cell.set', {
    sheetId: 's1',
    row: 9,
    column: 2,
    value: formula,
    writeAuthority: {
      kind: 'direct-entry',
      target: { sheetId: 's1', row: 9, column: 2 },
      candidate: formula,
      validationDecision: { status: 'accepted' },
    },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);

  const { rebased } = rebaseMutation(pending, committed, { sheetOrder: [{ id: 's1', name: 'S1' }] });
  const params = rebased.params as {
    row: number;
    value: { value: number; formula: string; formulaMetadata: { sourceFormula: string } };
    writeAuthority: { target: { row: number }; candidate: { value: number; formula: string; formulaMetadata: { sourceFormula: string } } };
  };
  assert.equal(params.row, 10);
  assert.equal(params.writeAuthority.target.row, 10);
  assert.deepEqual(params.writeAuthority.candidate, params.value);
  assert.equal(params.writeAuthority.candidate.formula, '=A11');
  assert.equal(params.writeAuthority.candidate.formulaMetadata.sourceFormula, '=A11');
});

test('keeps local-only data-region materialization outside collaboration rebase kinds', () => {
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);
  for (const mutationId of ['dataRegion.materialize.commit', 'dataRegion.materialize.restore']) {
    const localOnly = classifyMutation(mutationId, {}, 's1', []);
    assert.equal(localOnly.kind, 'unknown');
    assert.throws(() => rebaseMutation(localOnly, committed), /Cannot rebase unknown mutation/);
  }
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

test('fails closed for pending move, sort, table-resize, and pivot transforms without canonical patches', () => {
  const committed = classifyMutation('row.insert', { at: 2, count: 1 }, 's1', []);
  const unsupported = [
    ['range.move', { sourceRange: { sheetId: 's1', startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 }, targetOrigin: { row: 8, column: 2 } }, 'move-range'],
    ['sort.apply', { range: { sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 }, sourceRows: [5, 4, 6, 7, 8] }, 'sort'],
    ['table.resize', { tableId: 'table-1', range: { sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 } }, 'table-resize'],
    ['pivot.layout.set', { pivotId: 'pivot-1', sourceRange: { sheetId: 's1', startRow: 4, endRow: 8, startColumn: 0, endColumn: 2 } }, 'pivot-config'],
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
    transfer: 'copy',
    clearSource: false,
    targetOrigin: { row: 9, column: 2 },
    clipboard: {
      range: { sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
      occupiedCells: [{ rowOffset: 0, columnOffset: 0, value: { value: 1, formula: '=A10', formulaMetadata: { kind: 'normal', sourceFormula: '=A10' } } }],
      rangeMetadata: { validations: [{ formulaAnchor: { sheetId: 's1', row: 9, column: 2 } }] },
    },
    snapshot: {
      clearRanges: [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }],
      cells: [{ row: 9, column: 2, value: { value: 1, formula: '=A10', formulaMetadata: { kind: 'normal', sourceFormula: '=A10' } } }],
      notes: [{ key: '9:2' }],
      hyperlinks: [{ key: '9:2', value: { id: 'link-1', target: { kind: 'sheet', sheetId: 's1', address: 'A10' } } }],
      commentCells: ['9:2'],
      comments: [{ sheetId: 's1', row: 9, column: 2 }],
      columnWidths: [{ column: 2, widthPx: 80 }],
      validations: [{ formulaAnchor: { sheetId: 's1', row: 9, column: 2 }, ranges: [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }] }],
    },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const rowInsert = classifyMutation('row.insert', { at: 5 }, 's1', []);
  const columnInsert = classifyMutation('column.insert', { at: 0 }, 's1', []);
  const context = { sheetOrder: [{ id: 's1', name: 'S1' }] };

  const afterRows = rebaseMutation(pending, rowInsert, context).rebased;
  const { rebased } = rebaseMutation(afterRows, columnInsert, context);
  const params = rebased.params as {
    targetOrigin: { row: number; column: number };
    clipboard: {
      range: { startRow: number; startColumn: number };
      occupiedCells: Array<{ value: { formula: string; formulaMetadata: { sourceFormula: string } } }>;
      rangeMetadata: { validations: Array<{ formulaAnchor: { row: number; column: number } }> };
    };
    snapshot: {
      cells: Array<{ row: number; column: number; value: { formula: string; formulaMetadata: { sourceFormula: string } } }>;
      notes: Array<{ key: string }>;
      hyperlinks: Array<{ key: string; value: { target: { address: string } } }>;
      commentCells: string[];
      comments: Array<{ row: number; column: number }>;
      columnWidths: Array<{ column: number }>;
      validations: Array<{ formulaAnchor: { row: number; column: number } }>;
    };
  };

  assert.deepEqual(params.targetOrigin, { row: 10, column: 3 });
  assert.deepEqual(params.clipboard.range, { sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 });
  assert.equal(params.clipboard.occupiedCells[0]?.value.formula, '=A10');
  assert.equal(params.clipboard.occupiedCells[0]?.value.formulaMetadata.sourceFormula, '=A10');
  assert.deepEqual(params.clipboard.rangeMetadata.validations[0]?.formulaAnchor, { sheetId: 's1', row: 9, column: 2 });
  assert.equal(params.snapshot.cells[0]?.row, 10);
  assert.equal(params.snapshot.cells[0]?.column, 3);
  assert.equal(params.snapshot.cells[0]?.value.formula, '=B11');
  assert.equal(params.snapshot.cells[0]?.value.formulaMetadata.sourceFormula, '=B11');
  assert.equal(params.snapshot.notes[0]?.key, '10:3');
  assert.equal(params.snapshot.hyperlinks[0]?.key, '10:3');
  assert.equal(params.snapshot.hyperlinks[0]?.value.target.address, 'B11');
  assert.equal(params.snapshot.commentCells[0], '10:3');
  assert.deepEqual(params.snapshot.comments[0], { sheetId: 's1', row: 10, column: 3 });
  assert.equal(params.snapshot.columnWidths[0]?.column, 3);
  assert.deepEqual(params.snapshot.validations[0]?.formulaAnchor, { sheetId: 's1', row: 10, column: 3 });
});

test('rejects a pasted worksheet hyperlink when the committed deletion removes its target', () => {
  const pending = classifyMutation('range.paste', {
    transfer: 'copy',
    clearSource: false,
    targetOrigin: { row: 0, column: 0 },
    snapshot: {
      cells: [],
      hyperlinks: [{ key: '0:0', value: { id: 'link-1', target: { kind: 'sheet', sheetId: 's1', address: 'A6' } } }],
    },
  }, 's1', []);
  const committed = classifyMutation('row.delete', { at: 5, count: 1 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /coordinate 5 was deleted/);
});

test('rejects deletion through a pending paste footprint even when its snapshot is sparse', () => {
  const pending = classifyMutation('range.paste', {
    transfer: 'copy',
    clearSource: false,
    targetOrigin: { row: 8, column: 0 },
    sourceExtent: { rows: 3, columns: 1 },
    spec: { transpose: false },
    snapshot: { cells: [] },
  }, 's1', [{ sheetId: 's1', startRow: 8, endRow: 10, startColumn: 0, endColumn: 0 }]);
  const committed = classifyMutation('row.delete', { at: 9, count: 1 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /committed deletion intersects pending range\.paste target/);
});

test('fails closed for pending cross-sheet moves without a canonical structural patch', () => {
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
  assert.throws(() => rebaseMutation(pending, committed), /pending range\.paste has no canonical structural patch/);
});

test('rebases range.set origins together with their declared worksheet footprint', () => {
  const pending = classifyMutation('range.set', {
    sheetId: 's1',
    startRow: 9,
    startColumn: 2,
    values: [[{ value: 'value' }]],
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);

  const { rebased } = rebaseMutation(pending, committed);
  assert.deepEqual(rebased.affectedRanges[0], { sheetId: 's1', startRow: 10, endRow: 10, startColumn: 2, endColumn: 2 });
  assert.equal((rebased.params as { startRow: number }).startRow, 10);
});

test('rebases exact fill writes with their source and target ranges', () => {
  for (const mutationId of ['fill.applied', 'fill.restored']) {
    const pending = classifyMutation(mutationId, {
      sheetId: 's1',
      sourceRange: { sheetId: 's1', startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 },
      targetRange: { sheetId: 's1', startRow: 3, endRow: 3, startColumn: 1, endColumn: 1 },
      direction: 'down',
      mode: 'copy',
      writes: [{ row: 3, column: 1, before: { value: null }, after: { value: 'copied' } }],
    }, 's1', [{ sheetId: 's1', startRow: 3, endRow: 3, startColumn: 1, endColumn: 1 }]);
    const committed = classifyMutation('row.insert', { at: 0 }, 's1', []);

    const { rebased } = rebaseMutation(pending, committed);
    const params = rebased.params as { sourceRange: { startRow: number }; targetRange: { startRow: number }; writes: Array<{ row: number }> };
    assert.equal(params.sourceRange.startRow, 3);
    assert.equal(params.targetRange.startRow, 4);
    assert.equal(params.writes[0]?.row, 4);
  }
});

test('rejects rebasing fill operations when a committed deletion intersects their source or target', () => {
  const pending = classifyMutation('fill.applied', {
    sheetId: 's1',
    sourceRange: { sheetId: 's1', startRow: 2, endRow: 4, startColumn: 1, endColumn: 1 },
    targetRange: { sheetId: 's1', startRow: 5, endRow: 6, startColumn: 1, endColumn: 1 },
    direction: 'down',
    mode: 'copy',
    writes: [{ row: 5, column: 1, after: { value: 'copied' } }],
  }, 's1', [{ sheetId: 's1', startRow: 5, endRow: 6, startColumn: 1, endColumn: 1 }]);
  const committed = classifyMutation('row.delete', { at: 3, count: 1 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /committed deletion intersects pending fill\.applied source range/);
});

test('rebases find-replacement match addresses and their canonical keys', () => {
  const pending = classifyMutation('find.replaced', {
    direction: 'forward',
    patches: [{
      kind: 'cell',
      match: {
        key: 's1!9:2:values:', sheetId: 's1', row: 9, column: 2, target: 'values',
        range: { sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 },
        text: 'old',
      },
      previous: { value: 'old' },
      next: { value: 'new' },
    }],
    affectedRanges: [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }],
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);

  const { rebased } = rebaseMutation(pending, committed);
  const match = (rebased.params as { patches: Array<{ match: { row: number; column: number; key: string } }> }).patches[0]!.match;
  assert.equal(match.row, 10);
  assert.equal(match.column, 2);
  assert.equal(match.key, 's1!10:2:values:');
});

test('rebases each workbook find formula by its own sheet identity and preserves source formula mapping', () => {
  const pending = classifyMutation('find.replaced', {
    direction: 'forward',
    patches: [{
      kind: 'cell',
      match: {
        key: 'target!9:0:formulas:', sheetId: 'target', row: 9, column: 0, target: 'formulas',
        range: { sheetId: 'target', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 },
        text: '=A10',
      },
      previous: { value: 1, formula: '=A10', formulaMetadata: { kind: 'normal', sourceFormula: '=A10' } },
      next: { value: 2, formula: '=A11', formulaMetadata: { kind: 'normal', sourceFormula: '=A11' } },
    }],
    affectedRanges: [{ sheetId: 'target', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 }],
  }, 'owner', [{ sheetId: 'target', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 'target', []);
  const context = { sheetOrder: [{ id: 'owner', name: 'Owner' }, { id: 'target', name: 'Target' }] };

  const { rebased } = rebaseMutation(pending, committed, context);
  const patch = (rebased.params as { patches: Array<{ previous: { formula: string; formulaMetadata: { sourceFormula: string } }; next: { formula: string; formulaMetadata: { sourceFormula: string } } }> }).patches[0]!;
  assert.equal(patch.previous.formula, '=A11');
  assert.equal(patch.previous.formulaMetadata.sourceFormula, '=A11');
  assert.equal(patch.next.formula, '=A12');
  assert.equal(patch.next.formulaMetadata.sourceFormula, '=A12');
});

test('rebases comment.add envelope and embedded thread to the same cell', () => {
  const pending = classifyMutation('comment.add', {
    sheetId: 's1', row: 9, column: 2,
    thread: { id: 'comment-1', sheetId: 's1', row: 9, column: 2, text: 'comment', replies: [] },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);

  const { rebased } = rebaseMutation(pending, committed);
  const params = rebased.params as { row: number; column: number; thread: { row: number; column: number } };
  assert.deepEqual({ row: params.row, column: params.column }, { row: 10, column: 2 });
  assert.deepEqual({ row: params.thread.row, column: params.thread.column }, { row: 10, column: 2 });
});

test('fails closed when a pending cell restore carries unsupported formula-group metadata', () => {
  const pending = classifyMutation('cell.restore', {
    sheetId: 's1', row: 9, column: 0,
    previous: { value: 1, formula: '=A1', formulaMetadata: { kind: 'array', range: 'A1:A3' } },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 }]);
  const committed = classifyMutation('row.insert', { at: 0 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /formula-group metadata without a canonical rebase transform/);
});

test('rejects deleting part of a pending range.set grid instead of compressing its values', () => {
  const pending = classifyMutation('range.set', {
    sheetId: 's1', startRow: 8, startColumn: 0,
    values: [[{ value: 'a' }], [{ value: 'b' }], [{ value: 'c' }]],
  }, 's1', [{ sheetId: 's1', startRow: 8, endRow: 10, startColumn: 0, endColumn: 0 }]);
  const committed = classifyMutation('row.delete', { at: 9, count: 1 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /committed deletion intersects pending range\.set writes/);
});

test('rebases review mutations with cell-owned note and hyperlink coordinates', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 's1', []);
  const mutations = [
    ['note.set', { sheetId: 's1', row: 9, column: 2, note: { id: 'n1', text: 'note' } }],
    ['note.remove', { sheetId: 's1', row: 9, column: 2 }],
    ['note.visibility', { sheetId: 's1', row: 9, column: 2, visible: false }],
    ['hyperlink.set', {
      sheetId: 's1', row: 9, column: 2,
      hyperlink: { id: 'h1', target: { kind: 'sheet', sheetId: 's1', address: 'C10' } },
    }],
    ['hyperlink.remove', { sheetId: 's1', row: 9, column: 2 }],
  ] as const;

  for (const [mutationId, params] of mutations) {
    const pending = classifyMutation(mutationId, params, 's1', [{
      sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2,
    }]);
    const { rebased } = rebaseMutation(pending, committed);
    const rebasedParams = rebased.params as { row: number; column: number; hyperlink?: { target: { address?: string } } };
    assert.equal(rebased.kind, 'comment');
    assert.equal(rebasedParams.row, 10);
    assert.equal(rebasedParams.column, 2);
    if (mutationId === 'hyperlink.set') assert.equal(rebasedParams.hyperlink?.target.address, 'C11');
  }
});

test('rebases stable-id comment operations while moving their conflict ranges', () => {
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 's1', []);
  const mutations = ['comment.reply', 'comment.reply.remove', 'comment.resolve', 'comment.remove'] as const;
  for (const mutationId of mutations) {
    const params = mutationId === 'comment.reply'
      ? { sheetId: 's1', threadId: 'thread-1', reply: { id: 'reply-1', text: 'reply' } }
      : mutationId === 'comment.reply.remove'
        ? { sheetId: 's1', threadId: 'thread-1', replyId: 'reply-1' }
        : mutationId === 'comment.resolve'
          ? { sheetId: 's1', threadId: 'thread-1', resolved: true, resolvedAt: '2026-01-01T00:00:00Z' }
          : { sheetId: 's1', threadId: 'thread-1' };
    const originalParams = structuredClone(params);
    const pending = classifyMutation(mutationId, params, 's1', [{
      sheetId: 's1', startRow: 9, endRow: 9, startColumn: 2, endColumn: 2,
    }]);
    const { rebased } = rebaseMutation(pending, committed);
    assert.equal(rebased.kind, 'comment');
    assert.deepEqual(rebased.params, originalParams);
    assert.deepEqual(rebased.affectedRanges[0], {
      sheetId: 's1', startRow: 10, endRow: 10, startColumn: 2, endColumn: 2,
    });
  }
});

test('rejects deleting a pending internal hyperlink destination', () => {
  const pending = classifyMutation('hyperlink.set', {
    sheetId: 's2', row: 2, column: 3,
    hyperlink: { id: 'h1', target: { kind: 'sheet', sheetId: 's1', row: 8, column: 4 } },
  }, 's2', [{ sheetId: 's2', startRow: 2, endRow: 2, startColumn: 3, endColumn: 3 }]);
  const committed = classifyMutation('row.delete', { at: 8, count: 1 }, 's1', []);

  assert.throws(() => rebaseMutation(pending, committed), /coordinate 8 was deleted/);
});

test('rebases a hyperlink destination by target sheet without moving its source cell', () => {
  const pending = classifyMutation('hyperlink.set', {
    sheetId: 'source', row: 2, column: 3,
    hyperlink: { id: 'h1', target: { kind: 'sheet', sheetId: 'target', address: 'C10' } },
  }, 'source', [{ sheetId: 'source', startRow: 2, endRow: 2, startColumn: 3, endColumn: 3 }]);
  const committed = classifyMutation('row.insert', { at: 5, count: 1 }, 'target', []);

  const { rebased } = rebaseMutation(pending, committed);
  const params = rebased.params as {
    row: number;
    column: number;
    hyperlink: { target: { address: string } };
  };
  assert.deepEqual({ row: params.row, column: params.column }, { row: 2, column: 3 });
  assert.equal(params.hyperlink.target.address, 'C11');
});

test('rebases hidden row and column state indices by their own axis', () => {
  const cases = [
    {
      mutationId: 'rows.visibility',
      params: { sheetId: 's1', states: [{ row: 9, hidden: true }, { row: 12, hidden: false }] },
      delta: classifyMutation('row.insert', { at: 5 }, 's1', []),
      affectedRanges: [
        { sheetId: 's1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 },
        { sheetId: 's1', startRow: 12, endRow: 12, startColumn: 0, endColumn: 0 },
      ],
      axis: 'row',
    },
    {
      mutationId: 'columns.visibility',
      params: { sheetId: 's1', states: [{ column: 9, hidden: true }, { column: 12, hidden: false }] },
      delta: classifyMutation('column.insert', { at: 5 }, 's1', []),
      affectedRanges: [
        { sheetId: 's1', startRow: 0, endRow: 0, startColumn: 9, endColumn: 9 },
        { sheetId: 's1', startRow: 0, endRow: 0, startColumn: 12, endColumn: 12 },
      ],
      axis: 'column',
    },
  ] as const;

  for (const entry of cases) {
    const pending = classifyMutation(entry.mutationId, entry.params, 's1', [...entry.affectedRanges]);
    const { rebased } = rebaseMutation(pending, entry.delta);
    const states = (rebased.params as { states: Array<Record<string, unknown>> }).states;
    assert.equal(rebased.kind, 'visibility');
    assert.deepEqual(states.map((state) => state[entry.axis]), [10, 13]);
  }
});

test('rebases cell-editor mutations through their declared ranges', () => {
  const pending = classifyMutation('cell.editor.set', {
    sheetId: 's1',
    ranges: [
      { sheetId: 's1', startRow: 9, endRow: 10, startColumn: 2, endColumn: 4 },
    ],
    editor: { kind: 'number', min: 0, max: 100 },
  }, 's1', [{ sheetId: 's1', startRow: 9, endRow: 10, startColumn: 2, endColumn: 4 }]);
  const committed = classifyMutation('row.insert', { at: 5 }, 's1', []);

  const { rebased } = rebaseMutation(pending, committed);
  const params = rebased.params as { ranges: Array<{ startRow: number; endRow: number }> };
  assert.equal(rebased.kind, 'cell-style');
  assert.deepEqual(params.ranges, [{ sheetId: 's1', startRow: 10, endRow: 11, startColumn: 2, endColumn: 4 }]);
});

test('rejects rebasing visibility state whose row or column was deleted', () => {
  const cases = [
    ['rows.visibility', { sheetId: 's1', states: [{ row: 9, hidden: true }] }, 'row.delete'],
    ['columns.visibility', { sheetId: 's1', states: [{ column: 9, hidden: true }] }, 'column.delete'],
  ] as const;
  for (const [mutationId, params, structuralId] of cases) {
    const pending = classifyMutation(mutationId, params, 's1', []);
    const committed = classifyMutation(structuralId, { at: 9, count: 1 }, 's1', []);
    assert.throws(() => rebaseMutation(pending, committed), /coordinate 9 was deleted/);
  }
});
