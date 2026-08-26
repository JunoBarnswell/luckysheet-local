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
  WorkspaceDatabaseCoordinator,
  WorkspaceStorageError,
  requestResult,
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
      contains: (name: string) => [WORKSPACE_STORE_NAME, DATA_BLOCK_STORE_NAME, NATIVE_PACKAGE_STORE_NAME, OVERLAY_STORE_NAME, ASSET_STORE_NAME].includes(name),
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
    const store = new LocalWorkspaceStore({ databaseName: 'persistence-test-catalog' });
    const persistence = new WorkspacePersistence({ databaseName: 'persistence-test-catalog' });
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
    const persistence = new WorkspacePersistence({ databaseName });
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

  it('upgrades v6 to v7 without deleting workspace, package, block, or overlay records', async () => {
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
    seed.objectStore(WORKSPACE_STORE_NAME).put({ unitId: 'legacy-workbook', unknownField: { preserved: true } });
    seed.objectStore(DATA_BLOCK_STORE_NAME).put({ sourceId: 'legacy-source', blockId: 'block-1', bytes: new Uint8Array([1, 2, 3]) });
    seed.objectStore(NATIVE_PACKAGE_STORE_NAME).put({ unitId: 'legacy-workbook', opaqueBinary: new Uint8Array([4, 5, 6]) });
    seed.objectStore(OVERLAY_STORE_NAME).put({ sourceId: 'legacy-source', blockId: 'block-1', revision: 2, marker: 'overlay' });
    await transactionComplete(seed);
    legacyDatabase.close();

    const coordinator = new WorkspaceDatabaseCoordinator({ databaseName, indexedDB, broadcast: false });
    const upgraded = await coordinator.open();
    assert.equal(upgraded.version, 7);
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
});
