import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  LocalDraftStore,
  buildLocalDraftRecord,
  buildPersistenceMeta,
  isDraftNewerThanServer,
  verifyLocalDraft,
} from './storage';

describe('persistence storage', () => {
  it('builds local draft records with checksum verification', () => {
    const snapshot = new WorkbookModel('wb-draft', 'Draft').snapshot();
    const record = buildLocalDraftRecord(snapshot, 3);
    assert.equal(record.unitId, 'wb-draft');
    assert.equal(record.revision, 3);
    assert.equal(verifyLocalDraft(record), true);
    assert.equal(isDraftNewerThanServer(record, 2), true);
    assert.equal(isDraftNewerThanServer(record, 3), false);
  });

  it('tracks persistence metadata and local draft presence', () => {
    const snapshot = new WorkbookModel('wb-meta', 'Meta').snapshot();
    const draft = buildLocalDraftRecord(snapshot, 1);
    const meta = buildPersistenceMeta(snapshot, 0, draft);
    assert.equal(meta.hasLocalDraft, true);
    assert.equal(meta.checksum.length, 64);
    assert.equal(meta.draftUpdatedAt, draft.updatedAt);
  });

  it('persists drafts through LocalDraftStore when window is available', () => {
    if (typeof globalThis.window === 'undefined') return;
    const store = new LocalDraftStore();
    const snapshot = new WorkbookModel('wb-store', 'Store').snapshot();
    const record = buildLocalDraftRecord(snapshot, 5);
    store.write(record);
    const loaded = store.read('wb-store');
    assert.equal(loaded?.revision, 5);
    store.clear('wb-store');
    assert.equal(store.read('wb-store'), null);
  });
});
