import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { createNativePackageState } from '@react-sheets/exchange-excel-ooxml';
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
    const artifact = await createNativePackageState({
      fileName: 'artifact.xlsx',
      buffer: new Uint8Array([80, 75, 3, 4]).buffer,
      dateSystem: '1900',
      packageGraph: {
        schema: 'OpcPackageGraph',
        workbookPart: 'xl/workbook.xml',
        parts: {},
        opaqueParts: {},
        relationships: {},
        sheetPartById: {},
        dateSystem: '1900',
        format: { family: 'ooxml', profile: 'transitional', variant: 'xlsx' },
        profile: 'transitional',
      } satisfies OpcPackageGraph,
      detectedFeatures: ['worksheet'],
    });

    const record = await persistence.checkpointWithArtifact(snapshot, 3, 0, 'local-only', artifact);
    assert.equal((await persistence.load(snapshot.unitId))?.checksum, record.checksum);
    assert.equal((await persistence.nativePackages.load(snapshot.unitId))?.checksum, artifact.checksum);
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
    const database = {
      onversionchange: null as (() => void) | null,
      close: () => { closed = true; },
    } as unknown as IDBDatabase;
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
    assert.equal(closed, true);
  });
});
