import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { createCellSetMutationParams } from '@react-sheets/sheet-features';
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
      schema: 'OperationEnvelope',
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

  it('replays a canonical bulk row visibility mutation without splitting history semantics', () => {
    const workbook = new WorkbookModel('wb-rows-visibility-replay', 'Rows visibility replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheetId = workbook.primarySheetId;
    const affectedRanges = [1, 2].map((row) => ({ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: 0 }));

    session.applyRemote({
      schema: 'OperationEnvelope', operationId: 'remote-rows-visibility', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
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
      schema: 'OperationEnvelope', operationId: 'remote-font', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'style.set', sheetId: sheet.id, params: { sheetId: sheet.id, range, style: { fontFamily: '  aRiAl  ' } }, affectedRanges: [range] }],
    });
    assert.equal(sheet.cells.get(0, 0)?.style?.fontFamily, 'Arial');
    assert.equal(runtime.undo(), false);

    assert.throws(() => session.applyRemote({
      schema: 'OperationEnvelope', operationId: 'remote-empty-font', unitId: workbook.unitId, actorId: 'actor-3', origin: 'client',
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
      schema: 'OperationEnvelope', operationId: 'remote-clear', unitId: 'wb-clear-replay', actorId: 'actor-2', origin: 'client',
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
