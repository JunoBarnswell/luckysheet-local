import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { createCellSetMutationParams, createPasteSpecialSpec } from '@react-sheets/sheet-features';
import { registerSpreadsheetFeatures } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { WorkbookSession } from './workbook-session';
import { CollaborationSession } from './collaboration/collaboration-session';
import { classifyMutation } from './collaboration/operation-types';
import { rebaseMutation } from './collaboration/ot-rebase';
import { createSpreadsheetRuntime } from './runtime';

describe('WorkbookSession collaboration integration', () => {
  it('exposes collaboration snapshot defaults when session is offline', () => {
    const app = new WorkbookSession();
    const snapshot = app.getCollaborationSnapshot();
    assert.equal(snapshot.pendingCount, 0);
    assert.equal(snapshot.offlineQueueState, 'offline');
    assert.equal(app.getUiSnapshot().pendingChangeSetCount, 0);
  });

  it('publishes edit target/status lifecycle without broadcasting draft characters', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    runtime.collaboration = new CollaborationSession(runtime.commands);
    const broadcasts: unknown[] = [];
    runtime.broadcastPresence = (state) => { broadcasts.push(structuredClone(state)); return true; };
    app.cellEdit.dispatch({ type: 'begin.request', source: 'direct-typing', initialText: '=' });
    app.cellEdit.dispatch({ type: 'text.insert', text: 'SENSITIVE-DRAFT' });
    app.cellEdit.dispatch({ type: 'reference.begin' });
    const active = runtime.collaboration.presence.snapshot().editSessions[0];
    assert.equal(active?.status, 'point');
    assert.equal('draftPreview' in (active ?? {}), false);
    assert.equal(JSON.stringify(broadcasts).includes('SENSITIVE-DRAFT'), false);
    app.cellEdit.dispatch({ type: 'cancel' });
    assert.equal(runtime.collaboration.presence.snapshot().editSessions.length, 0);
    assert.deepEqual((broadcasts.at(-1) as { edit?: unknown }).edit, null);
  });

  it('rebaseMutation shifts cell references after structural row inserts', () => {
    const committed = classifyMutation('rows.inserted', { at: 5, count: 1 }, 'sheet-1', [{
      sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0,
    }]);
    const pending = classifyMutation('cell.set', { row: 9, column: 0 }, 'sheet-1', [{
      sheetId: 'sheet-1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
    }]);
    const { rebased, transformed } = rebaseMutation(pending, committed);
    assert.equal(transformed, true);
    assert.equal(rebased.affectedRanges[0]?.startRow, 10);
    assert.equal((rebased.params as { row: number }).row, 10);
  });

  it('applies remote changesets through CollaborationSession without local undo pollution', () => {
    const workbook = new WorkbookModel('wb-collab', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session',
      operationId: 'remote-op',
      unitId: 'wb-collab',
      actorId: 'actor-2',
      origin: 'client',
      clientSequence: 1,
      baseRevision: 0,
      revision: 1,
      committedAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set',
        sheetId: 'sheet-1',
        params: createCellSetMutationParams(
          workbook.getSheet('sheet-1'),
          { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'remote' } },
          'external-sync',
        ),
        affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      }],
      createdAt: new Date().toISOString(),
    });
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.value, 'remote');
    assert.equal(runtime.undo(), false);
  });

  it('invalidates overlapping local undo after a committed remote cell write', () => {
    const workbook = new WorkbookModel('wb-collab-history-overlap', 'Collaboration history overlap');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 2, endRow: 2, startColumn: 3, endColumn: 3 };
    runtime.execute('sheet.cell.set', { sheetId, row: 2, column: 3, value: { value: 'local' } });
    runtime.execute('sheet.cell.set', { sheetId, row: 8, column: 8, value: { value: 'unrelated-local' } });

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-overlap',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set', sheetId,
        params: createCellSetMutationParams(workbook.getSheet(sheetId), {
          sheetId, row: 2, column: 3, value: { value: 'remote' },
        }, 'external-sync'),
        affectedRanges: [range],
      }],
    });

    assert.equal(workbook.getSheet(sheetId).cells.get(2, 3)?.value, 'remote');
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(sheetId).cells.get(2, 3)?.value, 'remote');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('invalidates local undo after a committed range move without a canonical history transform', () => {
    const workbook = new WorkbookModel('wb-collab-history-move', 'Collaboration history move');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const sourceRange = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
    const targetRange = { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 };
    sheet.cells.set(0, 0, { value: 'source' });
    sheet.cells.set(1, 1, { value: 'previous-target' });
    runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 1, value: { value: 'local' } });

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-range-move',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'range.move', sheetId: sheet.id,
        params: { sheetId: sheet.id, sourceRange, targetOrigin: { row: 1, column: 1 } },
        affectedRanges: [sourceRange, targetRange],
      }],
    });

    assert.equal(sheet.cells.get(1, 1)?.value, 'source');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('rejects structural mutation envelopes whose sheet differs from their target', () => {
    const workbook = new WorkbookModel('wb-structural-scope', 'Structural scope');
    const target = workbook.addSheet('target-sheet', 'Target');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    target.cells.set(4, 0, { value: 'stays-at-row-five' });

    assert.throws(() => runtime.applyRemoteMutations([{
      id: 'rows.inserted',
      unitId: workbook.unitId,
      sheetId: workbook.primarySheetId,
      params: { sheetId: target.id, at: 2, count: 1 },
      affectedRanges: [{ sheetId: target.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }],
    }]), /envelope sheetId differs from its target/);
    assert.equal(target.cells.get(4, 0)?.value, 'stays-at-row-five');
    assert.equal(target.cells.get(5, 0), undefined);
  });

  it('invalidates unscoped history when another unscoped mutation arrives', () => {
    const workbook = new WorkbookModel('wb-defined-name-history', 'Defined name history');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheetId = workbook.primarySheetId;
    runtime.execute('workbook.name.set', { name: 'Rate', formula: '=Sheet1!A1', scope: 'workbook' });
    const session = new CollaborationSession(runtime);

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-name-set',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'name.set', sheetId,
        params: { model: { name: 'Rate', formula: '=Sheet1!A2', scope: 'workbook' } },
        affectedRanges: [],
      }],
    });

    assert.equal(workbook.getDefinedName('Rate')?.formula, '=Sheet1!A2');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('rejects range-paste history with omitted or out-of-footprint affected cells', () => {
    const workbook = new WorkbookModel('wb-paste-footprint', 'Paste footprint');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.range.paste', {
      sheetId,
      targetOrigin: { row: 2, column: 2 },
      clipboard: {
        schema: 'SparseClipboardPayload',
        range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        sourceExtent: { rows: 1, columns: 1 },
        occupiedCells: [{ rowOffset: 0, columnOffset: 0, value: { value: 'paste' } }],
        transfer: 'copy',
        rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
      },
      transfer: 'copy',
      spec: createPasteSpecialSpec({
        formatting: 'none',
        metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false },
      }),
    });
    const mutation = runtime.getUndoEntries().at(-1)?.forwardMutations[0];
    assert.ok(mutation);

    assert.ok(runtime.registry.validateMutationInfo({ ...mutation, affectedRanges: [] })
      .some((issue) => issue.code === 'invalid-affected-ranges'));
    const params = mutation.params as {
      snapshot: { cells: Array<{ row: number; column: number; value?: unknown }> };
      clipboard: { range: { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number }; [key: string]: unknown };
    };
    const escapedSnapshot = {
      ...mutation,
      params: {
        ...params,
        snapshot: { ...params.snapshot, cells: [...params.snapshot.cells, { row: 9, column: 9, value: { value: 'outside' } }] },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(escapedSnapshot)
      .some((issue) => issue.code === 'invalid-params'));
    const outOfBoundsSource = {
      ...mutation,
      params: {
        ...params,
        clipboard: {
          ...params.clipboard,
          range: { sheetId, startRow: 1_048_576, endRow: 1_048_576, startColumn: 0, endColumn: 0 },
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(outOfBoundsSource)
      .some((issue) => issue.code === 'invalid-params'));
    const invalidWidth = {
      ...mutation,
      params: { ...params, snapshot: { ...params.snapshot, columnWidths: [{ column: 2, widthPx: 0 }] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(invalidWidth)
      .some((issue) => issue.code === 'invalid-params'));
    const unexpectedMetadata = {
      ...mutation,
      params: { ...params, snapshot: { ...params.snapshot, validations: [] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(unexpectedMetadata)
      .some((issue) => issue.code === 'invalid-params'));
    const validationSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: true, columnWidths: false, conditionalFormats: false, hyperlinks: false },
    });
    const targetRange = { sheetId, startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 };
    const foreignRule = {
      ...mutation,
      params: {
        ...params,
        spec: validationSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          validations: [{ id: 'foreign-rule', sheetId: 'another-sheet', ranges: [{ sheetId: 'another-sheet', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], type: 'whole', formula1: '1' }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(foreignRule)
      .some((issue) => issue.code === 'invalid-params'));
    const malformedRule = {
      ...mutation,
      params: {
        ...params,
        spec: validationSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          validations: [{ id: 'unknown-rule', sheetId, ranges: [targetRange], type: 'unsupported', formula1: '1' }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(malformedRule)
      .some((issue) => issue.code === 'invalid-params'));

    const widthsSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: false, columnWidths: true, conditionalFormats: false, hyperlinks: false },
    });
    const outOfRangeWidth = {
      ...mutation,
      params: { ...params, spec: widthsSpec, snapshot: { ...params.snapshot, clearMetadataRanges: [targetRange], columnWidths: [{ column: 7, widthPx: 120 }] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(outOfRangeWidth)
      .some((issue) => issue.code === 'invalid-params'));

    const hyperlinksSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: true },
    });
    const nonCanonicalCellKey = {
      ...mutation,
      params: {
        ...params,
        spec: hyperlinksSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          hyperlinks: [{ key: '02:2', value: { id: 'link', target: { kind: 'url', url: 'https://example.com' } } }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(nonCanonicalCellKey)
      .some((issue) => issue.code === 'invalid-params'));

    const overlapRange = { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
    const overlappingMove = {
      ...mutation,
      affectedRanges: [overlapRange, overlapRange],
      params: {
        ...params,
        targetOrigin: { row: 0, column: 0 },
        transfer: 'move',
        clearSource: true,
        sourceRange: overlapRange,
        clipboard: { ...params.clipboard, transfer: 'move' },
        snapshot: { ...params.snapshot, clearRanges: [overlapRange, overlapRange], cells: [] },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(overlappingMove)
      .some((issue) => issue.code === 'invalid-params'));
  });

  it('replays a canonical bulk row visibility mutation without splitting history semantics', () => {
    const workbook = new WorkbookModel('wb-rows-visibility-replay', 'Rows visibility replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheetId = workbook.primarySheetId;
    const affectedRanges = [1, 2].map((row) => ({ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: 0 }));

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-rows-visibility', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'rows.visibility', sheetId, params: { sheetId, states: [{ row: 1, hidden: true }, { row: 2, hidden: true }] }, affectedRanges }],
    });
    assert.deepEqual([...workbook.getSheet(sheetId).hiddenRows].sort((left, right) => left - right), [1, 2]);
    assert.equal(runtime.undo(), false);
  });

  it('replays canonical font-family mutations and rejects an empty family atomically', () => {
    const workbook = new WorkbookModel('wb-font-replay', 'Font replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheet = workbook.getSheet('sheet-1');
    const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-font', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'style.set', sheetId: sheet.id, params: { sheetId: sheet.id, range, style: { fontFamily: '  aRiAl  ' } }, affectedRanges: [range] }],
    });
    assert.equal(sheet.cells.get(0, 0)?.style?.fontFamily, 'Arial');
    assert.equal(runtime.undo(), false);

    assert.throws(() => session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-empty-font', unitId: workbook.unitId, actorId: 'actor-3', origin: 'client',
      clientSequence: 2, baseRevision: 1, revision: 2, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'style.set', sheetId: sheet.id, params: { sheetId: sheet.id, range, style: { fontFamily: '   ' } }, affectedRanges: [range] }],
    }), /must not be empty/);
    assert.equal(sheet.cells.get(0, 0)?.style?.fontFamily, 'Arial');
  });

  it('replays canonical clear families and conditional-format cropping remotely', () => {
    const workbook = new WorkbookModel('wb-clear-replay', 'Clear Replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheet = workbook.getSheet('sheet-1');
    sheet.conditionalFormats.push({
      id: 'cf-replay',
      sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 }],
      type: 'highlight',
    });
    const session = new CollaborationSession(runtime);
    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-clear', unitId: 'wb-clear-replay', actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'range.clear', sheetId: sheet.id,
        params: { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 }, family: 'formats' },
        affectedRanges: [{ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 }],
      }],
    });
    assert.equal(sheet.conditionalFormats[0]?.ranges.length, 4);
    assert.equal(runtime.undo(), false);
  });

  it('turns local undo into a durable compensating operation', () => {
    const runtime = createSpreadsheetRuntime();
    const sheetId = runtime.model.primarySheetId;
    runtime.commands.execute('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'local' },
    });
    assert.equal(runtime.collaboration?.offlineQueue.getPendingCount(), 1);
    assert.equal(runtime.commands.undo(), true);
    const pending = runtime.collaboration?.offlineQueue.getPending() ?? [];
    assert.equal(pending.length, 2);
    assert.notEqual(pending[0]?.operation.operationId, pending[1]?.operation.operationId);
    assert.equal(runtime.model.getSheet(sheetId).cells.get(0, 0), undefined);
  });
});
