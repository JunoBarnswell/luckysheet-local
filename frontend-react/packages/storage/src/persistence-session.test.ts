import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { PersistenceSession, computeSnapshotChecksum } from './persistence-session';

test('PersistenceSession triggers snapshot on threshold', async () => {
  const snapshots: number[] = [];
  const session = new PersistenceSession({
    changesetThreshold: 2,
    persistSnapshot: async (record) => {
      snapshots.push(record.revision);
    },
  });

  const wb = new WorkbookModel('u1', 'Test');
  await session.writeSnapshot(wb.snapshot(), 0);
  assert.equal(snapshots.length, 1);

  await session.recordChangeSet({
    schema: 'CollaborationChangeSetV1',
    operationId: 'op-1',
    unitId: 'u1',
    actorId: 'a1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [],
    createdAt: new Date().toISOString(),
  });
  assert.equal(snapshots.length, 1);

  await session.recordChangeSet({
    schema: 'CollaborationChangeSetV1',
    operationId: 'op-2',
    unitId: 'u1',
    actorId: 'a1',
    clientSequence: 2,
    baseRevision: 1,
    mutations: [],
    createdAt: new Date().toISOString(),
  });
  assert.equal(snapshots.length, 2);
  assert.equal(snapshots[1], 2);
});

test('computeSnapshotChecksum is stable', () => {
  const json = '{"schema":"WorkbookSnapshotV1"}';
  const a = computeSnapshotChecksum(json);
  const b = computeSnapshotChecksum(json);
  assert.equal(a, b);
  assert.equal(a.length, 64);
});
