import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { createNativePackageState, exportSnapshotToXlsxBuffer, loadOpcPackageGraph } from '@react-sheets/exchange-excel-ooxml';
import type { OpcPackageGraph } from '@react-sheets/exchange-excel-ooxml';
import type { OperationEnvelope } from '@react-sheets/protocol';
import {
  LocalWorkspaceStore,
  OperationJournalStore,
  WorkspacePersistence,
  buildWorkspaceRecord,
  buildPersistenceMeta,
  verifyWorkspaceRecord,
} from './storage';
import {
  ASSET_STORE_NAME,
  DATA_BLOCK_STORE_NAME,
  NATIVE_PACKAGE_STORE_NAME,
  OVERLAY_STORE_NAME,
  WORKSPACE_STORE_NAME,
  WORKSPACE_HEAD_STORE_NAME,
  WORKSPACE_SNAPSHOT_STORE_NAME,
  WORKSPACE_OPERATION_STORE_NAME,
  WORKSPACE_CATALOG_STORE_NAME,
  WorkspaceDatabaseCoordinator,
  WorkspaceStorageError,
  disposeOwnedWorkspaceCoordinators,
  requestResult,
  resolveWorkspaceDatabaseCoordinator,
  transactionComplete,
} from './indexed-db';

interface FakeOpenRequest {
  result: IDBDatabase;
  onupgradeneeded: (() => void) | null;
  onsuccess: (() => void) | null;
  onerror: (() => void) | null;
  onblocked: (() => void) | null;
  error?: DOMException | null;
}

function fakeDatabase(): IDBDatabase {
  return {
    onversionchange: null,
    close: () => undefined,
    objectStoreNames: {
        contains: (name: string) => [WORKSPACE_STORE_NAME, WORKSPACE_HEAD_STORE_NAME, WORKSPACE_SNAPSHOT_STORE_NAME, WORKSPACE_OPERATION_STORE_NAME, WORKSPACE_CATALOG_STORE_NAME, DATA_BLOCK_STORE_NAME, NATIVE_PACKAGE_STORE_NAME, OVERLAY_STORE_NAME, ASSET_STORE_NAME].includes(name),
    },
  } as unknown as IDBDatabase;
}

describe('persistence storage', () => {
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
      schema: 'OperationEnvelope',
      operationId: 'offline-op-1',
      unitId: 'wb-operation-store',
      clientSequence: 7,
      baseRevision: 3,
      mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { row: 0, column: 0, value: { value: 1 } } }],
      createdAt: '2026-08-23T00:00:00.000Z',
    };
    store.write(operation.unitId, [operation], operation.clientSequence);
    const loaded = store.read(operation.unitId);
    assert.equal(loaded?.schema, 'PendingOperationJournal');
    assert.equal(loaded?.nextClientSequence, 7);
    assert.deepEqual(loaded?.operations, [operation]);
    store.clear(operation.unitId);
    assert.equal(store.read(operation.unitId), null);
  });

  it('builds a canonical workspace record with a checksummed pending journal', () => {
    const snapshot = new WorkbookModel('wb-workspace', 'Workspace').snapshot();
    const record = buildWorkspaceRecord({
      unitId: snapshot.unitId,
      snapshot,
      localRevision: 4,
      serverRevision: 2,
      syncMode: 'local-only',
      operations: [],
      nextClientSequence: 0,
    });
    assert.equal(record.schema, 'WorkspaceRecord');
    assert.equal(record.snapshot.schema, 'WorkbookSnapshot');
    assert.equal(record.localRevision, 4);
    assert.equal(record.pending.schema, 'PendingOperationJournal');
    assert.equal(verifyWorkspaceRecord(record), true);
  });

  it('opens, lists, checkpoints, and deletes local workspaces through IndexedDB', async () => {
    const store = new LocalWorkspaceStore({ databaseName: 'persistence-test-catalog', indexedDB });
    const persistence = new WorkspacePersistence({ databaseName: 'persistence-test-catalog', indexedDB });
    const snapshot = new WorkbookModel('wb-catalog', 'Catalog').snapshot();
    const record = await store.create({
      unitId: snapshot.unitId,
      snapshot,
      localRevision: 1,
      serverRevision: 0,
      syncMode: 'local-only',
      operations: [],
      nextClientSequence: 0,
    });
    assert.equal((await store.open(snapshot.unitId))?.checksum, record.checksum);
    assert.equal((await store.list()).map((entry) => entry.unitId).includes(snapshot.unitId), true);
    persistence.operationJournal.hydrate(record);
    const checkpoint = await persistence.checkpoint(snapshot, 2, 0, 'local-only');
    assert.equal(checkpoint.localRevision, 2);
    await store.delete(snapshot.unitId);
    assert.equal(await store.open(snapshot.unitId), null);
  });

  it('checkpoints the workspace and source artifact through the same persistence namespace', async () => {
    const databaseName = `persistence-artifact-${Date.now()}-${Math.random()}`;
    const persistence = new WorkspacePersistence({ databaseName, indexedDB });
    const snapshot = new WorkbookModel('wb-artifact', 'Artifact').snapshot();
    const sourceBytes = exportSnapshotToXlsxBuffer(snapshot);
    const artifact = await createNativePackageState({
      fileName: 'artifact.xlsx',
      buffer: sourceBytes,
      dateSystem: '1900',
      packageGraph: loadOpcPackageGraph(sourceBytes).packageGraph satisfies OpcPackageGraph,
      detectedFeatures: ['worksheet'],
    });

    const record = await persistence.checkpointWithArtifact(snapshot, 3, 0, 'local-only', artifact);
    assert.equal((await persistence.load(snapshot.unitId))?.checksum, record.checksum);
    assert.equal((await persistence.nativePackages.load(snapshot.unitId))?.checksum, artifact.checksum);
  });

  it('commits operation journal and rejects a stale storage revision', async () => {
    const databaseName = `persistence-operation-${Date.now()}-${Math.random()}`;
    const persistence = new WorkspacePersistence({ databaseName, indexedDB });
    const snapshot = new WorkbookModel('wb-operation', 'Operation').snapshot();
    const created = await persistence.checkpoint(snapshot, 1, 0, 'local-only');
    const operation = {
      schema: 'OperationEnvelope' as const,
      operationId: 'operation-1',
      unitId: snapshot.unitId,
      clientSequence: 1,
      baseRevision: 0,
      mutations: [],
      createdAt: new Date().toISOString(),
    };
    const nextRevision = await persistence.commitOperationJournal(snapshot.unitId, [operation], 1, created.storageRevision);
    assert.equal(nextRevision, created.storageRevision + 1);
    const loaded = await persistence.load(snapshot.unitId);
    assert.equal(loaded?.pending.operations[0]?.operationId, operation.operationId);
    await assert.rejects(
      persistence.commitOperationJournal(snapshot.unitId, [], 2, created.storageRevision),
      (error: unknown) => (error as WorkspaceStorageError).code === 'STORAGE_REVISION_CONFLICT',
    );
  });

  it('patches catalog atomically without changing the checkpoint snapshot', async () => {
    const databaseName = `persistence-catalog-patch-${Date.now()}-${Math.random()}`;
    const persistence = new WorkspacePersistence({ databaseName, indexedDB });
    const snapshot = new WorkbookModel('wb-catalog-patch', 'Catalog patch').snapshot();
    const created = await persistence.checkpoint(snapshot, 1, 0, 'local-only');
    const patched = await persistence.updateMetadata(snapshot.unitId, { lifecycle: 'trashed' }, created.storageRevision);
    assert.equal(patched.metadata.lifecycle, 'trashed');
    assert.deepEqual(patched.snapshot, snapshot);
    assert.equal(patched.storageRevision, created.storageRevision + 1);
  });

  it('fails immediately when no IndexedDB factory is available', async () => {
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'unavailable-storage-test', indexedDB: null, broadcast: false });
    await assert.rejects(coordinator.open(), (error: unknown) => {
      assert.equal((error as WorkspaceStorageError).code, 'STORAGE_UNAVAILABLE');
      return true;
    });
  });

  it('upgrades v6 to v8 without deleting workspace, package, block, or overlay records', async () => {
    const databaseName = `persistence-v6-upgrade-${Date.now()}-${Math.random()}`;
    const request = indexedDB.open(databaseName, 6);
    request.onupgradeneeded = () => {
      const database = request.result;
      database.createObjectStore(WORKSPACE_STORE_NAME, { keyPath: 'unitId' });
      database.createObjectStore(DATA_BLOCK_STORE_NAME, { keyPath: ['sourceId', 'blockId'] });
      database.createObjectStore(NATIVE_PACKAGE_STORE_NAME, { keyPath: 'unitId' });
      database.createObjectStore(OVERLAY_STORE_NAME, { keyPath: ['sourceId', 'blockId', 'revision'] });
    };
    const legacyDatabase = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error('Unable to create v6 database'));
    });
    const seed = legacyDatabase.transaction(
      [WORKSPACE_STORE_NAME, DATA_BLOCK_STORE_NAME, NATIVE_PACKAGE_STORE_NAME, OVERLAY_STORE_NAME],
      'readwrite',
    );
    const legacySnapshot = new WorkbookModel('legacy-workbook', 'Legacy').snapshot();
    seed.objectStore(WORKSPACE_STORE_NAME).put({
      ...buildWorkspaceRecord({
        unitId: legacySnapshot.unitId,
        snapshot: legacySnapshot,
        localRevision: 1,
        serverRevision: 0,
        syncMode: 'local-only',
        operations: [],
        nextClientSequence: 0,
      }),
      unknownField: { preserved: true },
    });
    seed.objectStore(DATA_BLOCK_STORE_NAME).put({ sourceId: 'legacy-source', blockId: 'block-1', bytes: new Uint8Array([1, 2, 3]) });
    seed.objectStore(NATIVE_PACKAGE_STORE_NAME).put({ unitId: 'legacy-workbook', opaqueBinary: new Uint8Array([4, 5, 6]) });
    seed.objectStore(OVERLAY_STORE_NAME).put({ sourceId: 'legacy-source', blockId: 'block-1', revision: 2, marker: 'overlay' });
    await transactionComplete(seed);
    legacyDatabase.close();

    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName, indexedDB, broadcast: false });
    const upgraded = await coordinator.open();
    assert.equal(upgraded.version, 8);
    assert.equal(upgraded.objectStoreNames.contains(ASSET_STORE_NAME), true);
    const read = upgraded.transaction(
      [WORKSPACE_STORE_NAME, DATA_BLOCK_STORE_NAME, NATIVE_PACKAGE_STORE_NAME, OVERLAY_STORE_NAME],
      'readonly',
    );
    const [workspace, block, nativePackage, overlay] = await Promise.all([
      requestResult(read.objectStore(WORKSPACE_STORE_NAME).get('legacy-workbook')),
      requestResult(read.objectStore(DATA_BLOCK_STORE_NAME).get(['legacy-source', 'block-1'])),
      requestResult(read.objectStore(NATIVE_PACKAGE_STORE_NAME).get('legacy-workbook')),
      requestResult(read.objectStore(OVERLAY_STORE_NAME).get(['legacy-source', 'block-1', 2])),
    ]);
    await transactionComplete(read);
    assert.deepEqual((workspace as { unknownField: unknown }).unknownField, { preserved: true });
    assert.deepEqual([...((block as { bytes: Uint8Array }).bytes)], [1, 2, 3]);
    assert.deepEqual([...((nativePackage as { opaqueBinary: Uint8Array }).opaqueBinary)], [4, 5, 6]);
    assert.equal((overlay as { marker: string }).marker, 'overlay');
    const migrated = await new LocalWorkspaceStore({ databaseName, indexedDB }).open('legacy-workbook');
    assert.equal(migrated?.snapshot.name, 'Legacy');
    assert.equal(migrated?.storageRevision, 0);
    await coordinator.close();
  });

  it('fails closed when the database upgrade is blocked and permits an explicit retry', async () => {
    let attempts = 0;
    const database = fakeDatabase();
    const factory = {
      open: () => {
        const request: FakeOpenRequest = {
          result: database,
          onupgradeneeded: null,
          onsuccess: null,
          onerror: null,
          onblocked: null,
        };
        attempts += 1;
        queueMicrotask(() => {
          if (attempts === 1) request.onblocked?.();
          else request.onsuccess?.();
        });
        return request;
      },
    };

    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'blocked-retry-test', indexedDB: factory, openTimeoutMs: 10, broadcast: false });
    await assert.rejects(
      coordinator.open(),
      (error: unknown) => {
        assert.equal(error instanceof WorkspaceStorageError, true);
        assert.equal((error as WorkspaceStorageError).code, 'STORAGE_UPGRADE_BLOCKED');
        assert.match((error as WorkspaceStorageError).recovery, /刷新|关闭旧页面/);
        return true;
      },
    );

    assert.equal(await coordinator.open(), database);
  });

  it('returns STORAGE_OPEN_TIMEOUT when an open request never becomes blocked or ready', async () => {
    const factory = { open: () => ({ result: fakeDatabase(), onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null }) };
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'open-timeout-test', indexedDB: factory, openTimeoutMs: 5, broadcast: false });
    await assert.rejects(coordinator.open(), (error: unknown) => {
      assert.equal((error as WorkspaceStorageError).code, 'STORAGE_OPEN_TIMEOUT');
      return true;
    });
    assert.equal(coordinator.state, 'failed');
  });

  it('closes a late zombie success without replacing the next generation connection', async () => {
    const requests: FakeOpenRequest[] = [];
    const databases = [fakeDatabase(), fakeDatabase()];
    let opens = 0;
    const factory = { open: () => {
      const request: FakeOpenRequest = { result: databases[opens++]!, onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null };
      requests.push(request);
      return request;
    } };
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'zombie-open-test', indexedDB: factory, openTimeoutMs: 5, broadcast: false });
    await assert.rejects(coordinator.open(), (error: unknown) => (error as WorkspaceStorageError).code === 'STORAGE_OPEN_TIMEOUT');
    const retry = coordinator.open();
    requests[1]!.onsuccess?.();
    assert.equal(await retry, databases[1]);
    let zombieClosed = false;
    databases[0]!.close = () => { zombieClosed = true; };
    requests[0]!.onsuccess?.();
    assert.equal(zombieClosed, true);
    assert.equal(coordinator.state, 'ready');
  });

  it('keeps a blocked open in the same generation after a peer closed message', async () => {
    const request: FakeOpenRequest = { result: fakeDatabase(), onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null };
    const factory = { open: () => { queueMicrotask(() => request.onblocked?.()); return request; } };
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'peer-closed-test', indexedDB: factory, openTimeoutMs: 50, broadcast: false });
    const opening = coordinator.open();
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    coordinator.deliverPeerMessageForTest({ type: 'closed', databaseName: 'peer-closed-test', targetVersion: 8, instanceId: 'peer' });
    request.onsuccess?.();
    assert.equal(await opening, request.result);
  });

  it('closes an open connection when a later schema upgrade requests a version change', async () => {
    let closed = false;
    const database = fakeDatabase();
    database.close = () => { closed = true; };
    const factory = {
      open: () => {
        const request: FakeOpenRequest = {
          result: database,
          onupgradeneeded: null,
          onsuccess: null,
          onerror: null,
          onblocked: null,
        };
        queueMicrotask(() => request.onsuccess?.());
        return request;
      },
    };

    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'version-change-test', indexedDB: factory, broadcast: false });
    assert.equal(await coordinator.open(), database);
    (database as unknown as { onversionchange: (() => void) | null }).onversionchange?.();
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    assert.equal(closed, true);
  });

  it('serializes reopen behind an in-flight lifecycle close', async () => {
    let openAttempts = 0;
    const databases = [fakeDatabase(), fakeDatabase()];
    const factory = {
      open: () => {
        const request: FakeOpenRequest = {
          result: databases[openAttempts]!,
          onupgradeneeded: null,
          onsuccess: null,
          onerror: null,
          onblocked: null,
        };
        openAttempts += 1;
        queueMicrotask(() => request.onsuccess?.());
        return request;
      },
    };
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'close-reopen-test', indexedDB: factory, broadcast: false });
    assert.equal(await coordinator.open(), databases[0]);

    const closing = coordinator.close('dispose');
    const reopening = coordinator.open();
    await closing;

    assert.equal(await reopening, databases[1]);
    assert.equal(coordinator.state, 'ready');
    assert.equal(openAttempts, 2);
  });

  it('fails a close when transactions do not drain, then remains reopenable', async () => {
    const database = fakeDatabase();
    let opens = 0;
    const factory = { open: () => {
      const request: FakeOpenRequest = { result: database, onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null };
      queueMicrotask(() => request.onsuccess?.());
      opens += 1;
      return request;
    } };
    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName: 'drain-timeout-test', indexedDB: factory, closeDrainTimeoutMs: 5, broadcast: false });
    await coordinator.open();
    (coordinator as unknown as { activeTransactions: number }).activeTransactions = 1;
    await assert.rejects(coordinator.close(), (error: unknown) => {
      assert.equal((error as WorkspaceStorageError).code, 'STORAGE_CLOSE_TIMEOUT');
      return true;
    });
    assert.equal(coordinator.state, 'failed');
    (coordinator as unknown as { activeTransactions: number }).activeTransactions = 0;
    assert.equal(await coordinator.open(), database);
    assert.equal(opens, 2);
  });

  it('removes disposed coordinators from the registry before the next HMR open', async () => {
    const databaseName = `hmr-registry-${Date.now()}-${Math.random()}`;
    const database = fakeDatabase();
    const factory = {
      open: () => {
        const request: FakeOpenRequest = { result: database, onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null };
        queueMicrotask(() => request.onsuccess?.());
        return request;
      },
    };
    const first = resolveWorkspaceDatabaseCoordinator({ databaseName, indexedDB: factory, broadcast: false });
    await first.open();
    await disposeOwnedWorkspaceCoordinators();
    const second = resolveWorkspaceDatabaseCoordinator({ databaseName, indexedDB: factory, broadcast: false });
    assert.notEqual(second, first);
    await second.open();
    await disposeOwnedWorkspaceCoordinators();
  });
});
