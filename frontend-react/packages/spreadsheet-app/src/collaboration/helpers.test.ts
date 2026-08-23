import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import type { OperationEnvelopeV2 } from '@react-sheets/protocol';
import { registerSpreadsheetFeatures } from '../feature-registry';
import { DrawingRuntime } from '../features/drawing';
import { CollaborationSession } from './collaboration-session';
import {
  buildCollaborationSnapshot,
  mapPeerCursor,
  updatePresenceFromPeer,
} from './helpers';

describe('collaboration helpers', () => {
  it('builds collaboration snapshot from V2 session state', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    session.setRevision(4);
    const peer = mapPeerCursor('peer-1', { row: 2, column: 3, sheetId: 'sheet-1', name: 'Alice' }, 'sheet-1');
    updatePresenceFromPeer(session, peer);
    const snapshot = buildCollaborationSnapshot(session, [peer]);
    assert.equal(snapshot.revision, 4);
    assert.equal(snapshot.peerCount, 1);
    assert.equal(snapshot.presence.users[0]?.displayName, 'Alice');
  });

  it('acknowledges V2 operations and drains the offline queue', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    const session = new CollaborationSession(runtime);
    session.enqueueLocalMutations([{
      id: 'cell.set',
      unitId: 'wb-1',
      sheetId: 'sheet-1',
      params: { row: 1, column: 1, value: { value: 2 } },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    }], 'wb-1', 'op-ack');
    assert.equal(session.offlineQueue.getPendingCount(), 1);
    session.acknowledge('op-ack', 5);
    assert.equal(session.getRevision(), 5);
    assert.equal(session.offlineQueue.getPendingCount(), 0);
  });

  it('rewrites durable pending intent after a remote structural revision', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const persisted: OperationEnvelopeV2[] = [];
    const session = new CollaborationSession(runtime, {
      persistPending: (operations) => {
        persisted.splice(0, persisted.length, ...operations);
      },
    });
    session.setRevision(1);
    session.enqueueLocalMutations([{
      id: 'cell.set',
      unitId: 'wb-1',
      sheetId: 'sheet-1',
      params: { row: 9, column: 0, value: { value: 2 } },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0 }],
    }], 'wb-1', 'op-pending');
    session.recordCommittedMutations([{
      id: 'rows.inserted',
      sheetId: 'sheet-1',
      params: { sheetId: 'sheet-1', at: 5, count: 1 },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 }],
    }]);
    // The public rebase call proves the transform itself; the queued rewrite
    // is exercised by applying the same committed operation through the wire.
    session.applyRemote({
      schema: 'OperationEnvelopeV2',
      operationId: 'remote-insert',
      unitId: 'wb-1',
      actorId: 'peer',
      clientSequence: 1,
      baseRevision: 1,
      revision: 2,
      committedAt: '2026-08-23T00:00:00.000Z',
      mutations: [{
        id: 'rows.inserted',
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', at: 5, count: 1 },
        affectedRanges: [{ sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 }],
      }],
      createdAt: '2026-08-23T00:00:00.000Z',
    });
    assert.equal(persisted[0]?.baseRevision, 2);
    assert.equal((persisted[0]?.mutations[0]?.params as { row: number }).row, 10);
  });
});
