import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
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
      clientSequence: 1,
      baseRevision: 0,
      revision: 1,
      committedAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set',
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'remote' } },
        affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      }],
      createdAt: new Date().toISOString(),
    });
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.value, 'remote');
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
