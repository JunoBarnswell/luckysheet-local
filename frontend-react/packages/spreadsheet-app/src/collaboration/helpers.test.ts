import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import type { OperationEnvelope } from '@react-sheets/protocol';
import { CollaborationSession } from './collaboration-session';
import { OfflineQueue } from './offline-queue';
import {
  buildOperation,
  buildCollaborationSnapshot,
  mapPeerCursor,
  updatePresenceFromPeer,
} from './helpers';

describe('collaboration helpers', () => {
  it('builds the single client operation contract without server-owned fields', () => {
    const operation = buildOperation('op-1', 'wb-1', 1, 0, [{ id: 'cell.set', sheetId: 'sheet-1', params: {} }], '2026-08-23T00:00:00.000Z');
    assert.deepEqual(operation, {
      schema: 'OperationEnvelope',
      operationId: 'op-1',
      unitId: 'wb-1',
      clientSequence: 1,
      baseRevision: 0,
      mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: {} }],
      createdAt: '2026-08-23T00:00:00.000Z',
    });
    assert.equal('actorId' in operation, false);
    assert.equal('affectedRanges' in operation, false);
  });

  it('builds collaboration snapshot from operation session state', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const session = new CollaborationSession(runtime);
    session.setRevision(4);
    const peer = mapPeerCursor('peer-1', { row: 2, column: 3, sheetId: 'sheet-1', name: 'Alice' }, 'sheet-1');
    updatePresenceFromPeer(session, peer);
    const snapshot = buildCollaborationSnapshot(session, [peer]);
    assert.equal(snapshot.revision, 4);
    assert.equal(snapshot.peerCount, 1);
    assert.equal(snapshot.presence.users[0]?.displayName, 'Alice');
  });

  it('acknowledges operations and drains the offline queue', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const session = new CollaborationSession(runtime);
    session.enqueueLocalMutations([{
      id: 'cell.set',
      unitId: 'wb-1',
      sheetId: 'sheet-1',
      params: { sheetId: 'sheet-1', row: 1, column: 1, value: { value: 2 } },
      affectedRanges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    }], 'wb-1', 'op-ack');
    assert.equal(session.offlineQueue.getPendingCount(), 1);
    session.acknowledge('op-ack', 5);
    assert.equal(session.getRevision(), 5);
    assert.equal(session.offlineQueue.getPendingCount(), 0);
    assert.equal(session.offlineQueue.getStatus('op-ack'), 'acked');
  });

  it('flushes a REST-style async transport and clears only after the returned revision', async () => {
    const workbook = new WorkbookModel('wb-rest', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const committed: string[] = [];
    const session = new CollaborationSession(runtime, {
      send: async (operation) => {
        committed.push(operation.operationId);
        return 8;
      },
    });
    session.enqueueLocalMutations([{
      id: 'cell.set',
      unitId: workbook.unitId,
      sheetId: workbook.primarySheetId,
      params: { sheetId: workbook.primarySheetId, row: 0, column: 0, value: { value: 8 } },
      affectedRanges: [{ sheetId: workbook.primarySheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    }], workbook.unitId, 'op-rest');
    session.offlineQueue.setOnline(true);
    await session.offlineQueue.flushAll();
    assert.deepEqual(committed, ['op-rest']);
    assert.equal(session.getRevision(), 8);
    assert.equal(session.offlineQueue.getPendingCount(), 0);
  });

  it('rejects unknown mutations and unmatched acknowledgements', () => {
    const workbook = new WorkbookModel('wb-unknown', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const session = new CollaborationSession(runtime);
    assert.throws(() => session.enqueueLocalMutations([{
      id: 'mutation.does-not-exist',
      unitId: workbook.unitId,
      sheetId: workbook.primarySheetId,
      params: {},
      affectedRanges: [],
    }], workbook.unitId), /Unknown mutation/);
    assert.equal(session.acknowledge('operation.does-not-exist', 1), false);
    assert.equal(session.getRevision(), 0);
  });

  it('rejects malformed durable operation records and remote unknown mutations', () => {
    assert.throws(() => new OfflineQueue({
      load: () => [{
        schema: 'WrongEnvelope',
        operationId: 'op-invalid',
        unitId: 'wb-1',
        clientSequence: 1,
        baseRevision: 0,
        mutations: [],
        createdAt: '2026-08-23T00:00:00.000Z',
      } as never],
    }), /Unsupported operation schema/);

    const workbook = new WorkbookModel('wb-remote-unknown', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const session = new CollaborationSession(runtime);
    assert.throws(() => session.applyRemote({
      schema: 'OperationEnvelope',
      operationId: 'remote-invalid',
      unitId: workbook.unitId,
      actorId: 'peer',
      clientSequence: 1,
      baseRevision: 0,
      revision: 1,
      committedAt: '2026-08-23T00:00:00.000Z',
      createdAt: '2026-08-23T00:00:00.000Z',
      mutations: [{
        id: 'mutation.does-not-exist',
        sheetId: workbook.primarySheetId,
        params: {},
        affectedRanges: [],
      }],
    }), /Unknown mutation/);
  });

  it('rewrites durable pending intent after a remote structural revision', () => {
    const workbook = new WorkbookModel('wb-1', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    const persisted: OperationEnvelope[] = [];
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
      params: { sheetId: 'sheet-1', row: 9, column: 0, value: { value: 2 } },
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
      schema: 'OperationEnvelope',
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
