import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { CollaborationSession } from './collaboration/collaboration-session';
import {
  acknowledgeChangeSet,
  buildChangeSet,
  buildCollaborationSnapshot,
  mapPeerCursor,
  updatePresenceFromPeer,
} from './collaboration-bridge';

describe('collaboration-bridge', () => {
  it('builds collaboration changesets and snapshots', () => {
    const changeSet = buildChangeSet('op-1', 'wb-1', 'actor-1', 1, 3, [{
      id: 'cell.set',
      sheetId: 'sheet-1',
      params: { row: 0, column: 0, value: { value: 1 } },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    }]);
    assert.equal(changeSet.baseRevision, 3);
    assert.equal(changeSet.mutations.length, 1);

    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    const session = new CollaborationSession(runtime, { actorId: 'actor-1' });
    session.setRevision(4);
    const peer = mapPeerCursor('peer-1', { row: 2, column: 3, sheetId: 'sheet-1', name: 'Alice' }, 'sheet-1');
    updatePresenceFromPeer(session, peer);
    const snapshot = buildCollaborationSnapshot(session, [peer]);
    assert.equal(snapshot.revision, 4);
    assert.equal(snapshot.peerCount, 1);
    assert.equal(snapshot.presence.users[0]?.displayName, 'Alice');
  });

  it('acknowledges changesets and drains the offline queue', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    const session = new CollaborationSession(runtime, { actorId: 'actor-1' });
    session.enqueueLocalMutations([{
      id: 'cell.set',
      unitId: 'wb-1',
      sheetId: 'sheet-1',
      params: { row: 1, column: 1, value: { value: 2 } },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    }], 'wb-1', 'op-ack');
    assert.equal(session.offlineQueue.getPendingCount(), 1);
    acknowledgeChangeSet(session, 5, 'op-ack');
    assert.equal(session.getRevision(), 5);
    assert.equal(session.offlineQueue.getPendingCount(), 0);
  });
});
