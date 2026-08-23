import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import type { CommittedOperationEnvelopeV2 } from '@react-sheets/protocol';
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

  await session.recordOperation({
    schema: 'OperationEnvelopeV2',
    operationId: 'op-1',
    unitId: 'u1',
    actorId: 'a1',
    clientSequence: 1,
    baseRevision: 0,
    revision: 1,
    committedAt: new Date().toISOString(),
    mutations: [],
    createdAt: new Date().toISOString(),
  });
  assert.equal(snapshots.length, 1);

  await session.recordOperation({
    schema: 'OperationEnvelopeV2',
    operationId: 'op-2',
    unitId: 'u1',
    actorId: 'a1',
    clientSequence: 2,
    baseRevision: 1,
    revision: 2,
    committedAt: new Date().toISOString(),
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

test('PersistenceSession isolates revision and snapshot thresholds per unit', async () => {
  const snapshots: Array<{ unitId: string; revision: number; schema: string }> = [];
  const session = new PersistenceSession({
    changesetThreshold: 1,
    persistSnapshot: async (record) => {
      snapshots.push({ unitId: record.unitId, revision: record.revision, schema: record.schema });
    },
  });
  const first = new WorkbookModel('unit-a', 'A');
  const second = new WorkbookModel('unit-b', 'B');
  await session.writeSnapshot(first.snapshot(), 0);
  await session.writeSnapshot(second.snapshot(), 0);
  await session.recordOperation({
    schema: 'OperationEnvelopeV2',
    operationId: 'unit-a-op',
    unitId: 'unit-a',
    actorId: 'actor',
    clientSequence: 1,
    baseRevision: 0,
    revision: 1,
    committedAt: new Date().toISOString(),
    mutations: [],
    createdAt: new Date().toISOString(),
  });
  assert.equal(session.getRevision('unit-a'), 1);
  assert.equal(session.getRevision('unit-b'), 0);
  assert.equal(snapshots.at(-1)?.unitId, 'unit-a');
  assert.equal(snapshots.at(-1)?.revision, 1);
  assert.equal(snapshots.at(-1)?.schema, 'SnapshotV2');
});

test('PersistenceSession records server-authored V2 operation revisions', async () => {
  const operations: CommittedOperationEnvelopeV2[] = [];
  const session = new PersistenceSession({
    appendOperation: (operation) => { operations.push(operation); },
  });
  await session.recordOperation({
    schema: 'OperationEnvelopeV2',
    operationId: 'v2-op',
    unitId: 'unit-v2',
    actorId: 'actor-v2',
    clientSequence: 1,
    baseRevision: 4,
    revision: 5,
    committedAt: '2026-08-23T00:00:00.000Z',
    createdAt: '2026-08-23T00:00:00.000Z',
    mutations: [],
  });
  assert.equal(operations[0]?.revision, 5);
  assert.equal(session.getRevision('unit-v2'), 5);
});
