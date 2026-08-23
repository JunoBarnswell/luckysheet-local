import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { CollaborationSession } from './collaboration/collaboration-session';
import {
  buildCollaborationSnapshot,
  mapPeerCursor,
  updatePresenceFromPeer,
} from './collaboration-bridge';

describe('collaboration-bridge', () => {
  it('builds collaboration snapshot from V2 session state', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
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
});
