import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CollaborationSession } from './collaboration-session';
import { buildCollaborationSnapshot, mapPeerCursor, updatePresenceFromPeer } from './helpers';

describe('collaboration presence and revision projection', () => {
  it('projects the server revision and peer presence without durable client state', () => {
    const session = new CollaborationSession();
    session.setRevision(4);
    const peer = mapPeerCursor('peer-1', { row: 2, column: 3, sheetId: 'sheet-1', name: 'Alice' }, 'sheet-1');
    updatePresenceFromPeer(session, peer);
    const snapshot = buildCollaborationSnapshot(session, [peer]);
    assert.equal(snapshot.revision, 4);
    assert.equal(snapshot.peerCount, 1);
    assert.equal(snapshot.presence.users[0]?.displayName, 'Alice');
    assert.equal('pendingCount' in snapshot, false);
    assert.equal('offlineQueueState' in snapshot, false);
  });

  it('rejects invalid and regressing revision announcements', () => {
    const session = new CollaborationSession();
    session.setRevision(7);
    assert.throws(() => session.setRevision(6), /COLLABORATION_REVISION_REGRESSION/);
    assert.throws(() => session.setRevision(-1), /COLLABORATION_REVISION_INVALID/);
    assert.equal(session.getRevision(), 7);
  });
});
