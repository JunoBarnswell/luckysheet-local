import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';
import { CollaborationSession } from './collaboration/collaboration-session';

describe('WorkbookSession collaboration integration', () => {
  it('exposes revision and presence only', () => {
    const app = new WorkbookSession();
    const snapshot = app.getCollaborationSnapshot();
    assert.equal(snapshot.revision, 0);
    assert.equal(snapshot.peerCount, 0);
    assert.equal('pendingCount' in snapshot, false);
  });

  it('publishes edit target/status without broadcasting draft characters', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    runtime.collaboration = new CollaborationSession();
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
});
