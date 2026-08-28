import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { createNativeDocumentArtifact, exportSnapshotToOoxmlBuffer, loadOpcPackageGraph } from '@react-sheets/exchange-excel-ooxml';
import type { OpcPackageGraph } from '@react-sheets/exchange-excel-ooxml';
import type { OperationEnvelope } from '@react-sheets/protocol';
import {
  OperationJournalStore,
  WorkspacePersistence,
  buildWorkspaceRecord,
  buildPersistenceMeta,
  verifyWorkspaceRecord,
} from './storage';
import {
  WorkspaceMemoryCoordinator,
  WorkspaceStorageError,
} from './memory';

describe('page-session memory persistence', () => {
  it('tracks pending local operation metadata', () => {
    const snapshot = new WorkbookModel('wb-meta', 'Meta').snapshot();
    const meta = buildPersistenceMeta(snapshot, 0, 1);
    assert.equal(meta.hasPendingOperations, true);
    assert.equal(meta.checksum.length, 64);
    assert.equal(meta.pendingOperationCount, 1);
  });

  it('persists only a monotonic pending-operation journal with checksum validation', () => {
    const store = new OperationJournalStore();
    const operation: OperationEnvelope = {
      schema: 'OperationEnvelope', operationId: 'offline-op-1', unitId: 'wb-operation-store',
      clientSequence: 7, baseRevision: 3, mutations: [], createdAt: '2026-08-23T00:00:00.000Z',
    };
    store.write(operation.unitId, [operation], operation.clientSequence);
    assert.deepEqual(store.read(operation.unitId)?.operations, [operation]);
    store.clear(operation.unitId);
    assert.equal(store.read(operation.unitId), null);
  });

  it('builds and validates a canonical workspace record', () => {
    const snapshot = new WorkbookModel('wb-workspace', 'Workspace').snapshot();
    const record = buildWorkspaceRecord({
      unitId: snapshot.unitId, snapshot, localRevision: 4, serverRevision: 2,
      syncMode: 'local-only', operations: [], nextClientSequence: 0,
    });
    assert.equal(verifyWorkspaceRecord(record), true);
  });

  it('creates, lists, checkpoints, and deletes workbooks in one memory context', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('wb-catalog', 'Catalog').snapshot();
    const record = await persistence.store.create({
      unitId: snapshot.unitId, snapshot, localRevision: 1, serverRevision: 0,
      syncMode: 'local-only', operations: [], nextClientSequence: 0,
    });
    assert.equal((await persistence.load(snapshot.unitId))?.checksum, record.checksum);
    assert.equal((await persistence.list()).length, 1);
    const checkpoint = await persistence.checkpoint(snapshot, 2, 0, 'local-only');
    assert.equal(checkpoint.localRevision, 2);
    await persistence.clear(snapshot.unitId);
    assert.equal(await persistence.load(snapshot.unitId), null);
  });

  it('checkpoints the workspace and native artifact atomically', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('wb-artifact', 'Artifact').snapshot();
    const sourceBytes = exportSnapshotToOoxmlBuffer(snapshot);
    const artifact = await createNativeDocumentArtifact({
      fileName: 'artifact.xlsx', buffer: sourceBytes, dateSystem: '1900',
      nativeGraph: { kind: 'opc', package: loadOpcPackageGraph(sourceBytes).packageGraph satisfies OpcPackageGraph },
      detectedFeatures: ['worksheet'],
    });
    const record = await persistence.checkpointWithArtifact(snapshot, 3, 0, 'local-only', artifact);
    assert.equal((await persistence.load(snapshot.unitId))?.checksum, record.checksum);
    assert.equal((await persistence.nativeDocuments.load(snapshot.unitId))?.checksum, artifact.checksum);
  });

  it('commits operation journal and rejects a stale storage revision', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('wb-operation', 'Operation').snapshot();
    const created = await persistence.checkpoint(snapshot, 1, 0, 'local-only');
    const operation: OperationEnvelope = {
      schema: 'OperationEnvelope', operationId: 'operation-1', unitId: snapshot.unitId,
      clientSequence: 1, baseRevision: 0, mutations: [], createdAt: new Date().toISOString(),
    };
    const nextRevision = await persistence.commitOperationJournal(snapshot.unitId, [operation], 1, created.storageRevision);
    assert.equal(nextRevision, created.storageRevision + 1);
    assert.equal((await persistence.load(snapshot.unitId))?.pending.operations[0]?.operationId, operation.operationId);
    await assert.rejects(
      persistence.commitOperationJournal(snapshot.unitId, [], 2, created.storageRevision),
      (error: unknown) => (error as WorkspaceStorageError).code === 'STORAGE_REVISION_CONFLICT',
    );
  });

  it('patches catalog without changing the checkpoint snapshot', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('wb-catalog-patch', 'Catalog patch').snapshot();
    const created = await persistence.checkpoint(snapshot, 1, 0, 'local-only');
    const patched = await persistence.updateMetadata(snapshot.unitId, { lifecycle: 'trashed' }, created.storageRevision);
    assert.equal(patched.metadata.lifecycle, 'trashed');
    assert.deepEqual(patched.snapshot, snapshot);
    assert.equal(patched.storageRevision, created.storageRevision + 1);
  });

  it('rolls back a failed multi-store transaction completely', async () => {
    const coordinator = new WorkspaceMemoryCoordinator();
    await coordinator.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { value: 'before' }));
    await assert.rejects(coordinator.transaction((transaction) => {
      transaction.set('workspaceCatalog', 'unit', { value: 'after' });
      transaction.set('workspaceHeads', 'unit', { value: 'after' });
      throw new Error('forced transaction failure');
    }));
    assert.deepEqual(await coordinator.read((transaction) => transaction.get('workspaceCatalog', 'unit')), { value: 'before' });
    assert.equal(await coordinator.read((transaction) => transaction.get('workspaceHeads', 'unit')), undefined);
  });

  it('serializes transactions and rejects a concurrent writer for the same workbook', async () => {
    const coordinator = new WorkspaceMemoryCoordinator();
    let release!: () => void;
    const first = coordinator.withWorkbookWriter('unit', () => new Promise<void>((resolve) => { release = resolve; }));
    await assert.rejects(
      coordinator.withWorkbookWriter('unit', async () => undefined),
      (error: unknown) => (error as WorkspaceStorageError).code === 'STORAGE_WRITER_UNAVAILABLE',
    );
    release();
    await first;
  });

  it('does not share data between page-session persistence instances', async () => {
    const first = new WorkspacePersistence();
    const second = new WorkspacePersistence();
    const snapshot = new WorkbookModel('isolated-unit', 'Isolated').snapshot();
    await first.checkpoint(snapshot, 1, 0, 'local-only');
    assert.equal((await second.list()).length, 0);
  });

  it('rejects operations after dispose and clears the session', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('disposed-unit', 'Disposed').snapshot();
    await persistence.checkpoint(snapshot, 1, 0, 'local-only');
    await persistence.disposeAsync();
    assert.equal(persistence.state, 'disposed');
    await assert.rejects(persistence.list(), (error: unknown) => (error as WorkspaceStorageError).code === 'STORAGE_MEMORY_DISPOSED');
  });
});
