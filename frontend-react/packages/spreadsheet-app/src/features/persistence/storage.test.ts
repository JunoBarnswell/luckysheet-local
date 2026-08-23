import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import type { OperationEnvelopeV2 } from '@react-sheets/protocol';
import {
  LocalDraftStore,
  LocalOperationStore,
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

  it('persists only a monotonic pending-operation journal with checksum validation', () => {
    const store = new LocalOperationStore();
    const operation: OperationEnvelopeV2 = {
      schema: 'OperationEnvelopeV2',
      operationId: 'offline-op-1',
      unitId: 'wb-operation-store',
      clientSequence: 7,
      baseRevision: 3,
      mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { row: 0, column: 0, value: { value: 1 } } }],
      createdAt: '2026-08-23T00:00:00.000Z',
    };
    store.write(operation.unitId, [operation], operation.clientSequence);
    const loaded = store.read(operation.unitId);
    assert.equal(loaded?.schema, 'PendingOperationJournalV1');
    assert.equal(loaded?.nextClientSequence, 7);
    assert.deepEqual(loaded?.operations, [operation]);
    store.clear(operation.unitId);
    assert.equal(store.read(operation.unitId), null);
  });
});
