import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exchangeExportXlsx } from '../xlsx';
import { WorkspacePersistence } from '../persistence';
import { LocalDataBlockStore } from '../persistence';
import { computeBinaryChecksum } from '../persistence/checksum';
import { LocalSparseOverlayStore } from '../data-source/overlay-store';
import { filterWorkbookCatalog, WORKBOOK_SYNC_STATE_PRIORITY } from './state';
import { WorkbookCatalogService } from './service';
import { createTemplateSnapshot } from './templates';
import type { WorkbookCatalogRemoteClient } from './types';

function service(name: string): WorkbookCatalogService {
  return new WorkbookCatalogService({
    persistence: new WorkspacePersistence({ databaseName: `catalog-${name}`, indexedDB: null }),
    unitIdFactory: (() => {
      let sequence = 0;
      return () => `catalog-${name}-${++sequence}`;
    })(),
  });
}

describe('WorkbookCatalogService', () => {
  it('creates independent template workbooks and keeps one local record per unitId', async () => {
    const catalog = service('identity');
    const first = await catalog.create({ snapshot: createTemplateSnapshot('blank', 'workbook-a') });
    const second = await catalog.create({ snapshot: createTemplateSnapshot('budget', 'workbook-b') });

    assert.notEqual(first.unitId, second.unitId);
    assert.equal((await catalog.open(first.unitId)).snapshot.name, '空白工作簿');
    assert.equal((await catalog.open(second.unitId)).snapshot.name, '预算模板');
    assert.equal((await catalog.list()).length, 2);
  });

  it('moves local workbooks to trash, restores them, and purges source artifacts', async () => {
    const catalog = service('trash');
    const created = await catalog.create({ snapshot: createTemplateSnapshot('blank', 'trash-source') });
    const trashed = await catalog.moveToTrash(created.unitId);
    assert.equal(trashed.lifecycle, 'trashed');
    assert.equal((await catalog.list()).some((entry) => entry.unitId === created.unitId), false);
    assert.equal((await catalog.list({ view: 'trash' })).some((entry) => entry.unitId === created.unitId), true);

    const restored = await catalog.restore(created.unitId);
    assert.equal(restored.lifecycle, 'active');
    assert.equal((await catalog.list()).some((entry) => entry.unitId === created.unitId), true);
    await catalog.purge(created.unitId);
    await assert.rejects(() => catalog.open(created.unitId), /not found/i);
  });

  it('imports XLSX as a new workbook and exports it through the same catalog boundary', async () => {
    const catalog = service('xlsx');
    const original = await catalog.create({ snapshot: createTemplateSnapshot('template', 'original') });
    const source = new WorkbookModel('source-xlsx', 'Source XLSX');
    source.getSheet(source.primarySheetId).cells.set(0, 0, { value: 'preserve' });
    const generated = await exchangeExportXlsx(source.snapshot(), { fileName: 'source.xlsx', execution: 'inline-test' });
    assert.ok(generated.buffer);

    const imported = await catalog.importXlsx({ fileName: 'source.xlsx', buffer: generated.buffer!, execution: 'inline-test' });
    assert.notEqual(imported.entry.unitId, original.unitId);
    assert.notEqual(imported.entry.unitId, source.unitId);
    assert.equal((await catalog.open(original.unitId)).snapshot.name, '会议记录模板');
    const exported = await catalog.exportXlsx(imported.entry.unitId, { execution: 'inline-test' });
    assert.ok(exported.buffer.byteLength > 0);
    assert.equal(exported.fileName, 'Imported Workbook.xlsx');
  });

  it('upgrades an explicit local workbook to the remote catalog through create plus checkpoint', async () => {
    let remoteSnapshot: ReturnType<WorkbookModel['snapshot']> | null = null;
    let syncedUnitId: string | null = null;
    const remote: WorkbookCatalogRemoteClient = {
      getSnapshot: async () => ({ snapshot: remoteSnapshot!, revision: 1 }),
      getAccess: async (unitId) => ({ unitId, role: 'owner' }),
      listWorkbookAcl: async () => [],
      putWorkbookAcl: async (unitId, subject, role) => ({ unitId, subject, role, createdAt: '', updatedAt: '' }),
      deleteWorkbookAcl: async () => undefined,
      createWorkbook: async (snapshot) => {
        remoteSnapshot = structuredClone(snapshot);
        syncedUnitId = snapshot.unitId;
        return { snapshot: structuredClone(snapshot), revision: 1 };
      },
      listWorkbooks: async () => remoteSnapshot ? [{ unitId: remoteSnapshot.unitId, name: remoteSnapshot.name, revision: 1, updatedAt: new Date().toISOString(), role: 'owner' }] : [],
      updateWorkbook: async (unitId, patch) => ({ unitId, name: patch.name ?? remoteSnapshot?.name ?? '', revision: 1, updatedAt: new Date().toISOString(), role: 'owner' }),
      copyWorkbook: async (unitId) => ({ unitId: `${unitId}-copy`, name: remoteSnapshot?.name ?? '', revision: 1, updatedAt: new Date().toISOString(), role: 'owner' }),
      moveToTrash: async () => undefined,
      restoreFromTrash: async (unitId) => ({ unitId, name: remoteSnapshot?.name ?? '', revision: 1, updatedAt: new Date().toISOString(), role: 'owner' }),
      purgeWorkbook: async () => undefined,
      getWorkbookUserState: async (unitId) => ({ unitId }),
      putWorkbookUserState: async (unitId) => ({ unitId }),
      listSpaces: async () => [],
      createSpace: async (input) => ({ ...input, spaceId: 'space', createdAt: '', createdBy: '', updatedAt: '' }),
      listFolders: async () => [],
      createFolder: async () => ({ folderId: 'folder', name: 'folder', spaceId: 'space', updatedAt: '' }),
      updateFolder: async () => ({ folderId: 'folder', name: 'folder', spaceId: 'space', updatedAt: '' }),
      deleteFolder: async () => undefined,
      listSpaceMembers: async () => [],
      putSpaceMember: async (spaceId, subject, role) => ({ spaceId, subject, role, updatedAt: '' }),
      deleteSpaceMember: async () => undefined,
      createWorkbookImport: async (input) => ({
        snapshot: structuredClone(input.snapshot),
        summary: { unitId: input.snapshot.unitId, name: input.snapshot.name, revision: 1, updatedAt: new Date().toISOString(), role: 'owner' },
        artifact: { unitId: input.snapshot.unitId, fileName: input.artifactFileName, byteLength: input.artifact.size, checksum: '', updatedAt: new Date().toISOString() },
      }),
      putWorkbookSourceArtifact: async (unitId, artifact, fileName) => ({ unitId, fileName, byteLength: artifact.size, checksum: '', updatedAt: '' }),
      getWorkbookSourceArtifact: async () => { throw new Error('no artifact'); },
      commitOperation: async (_unitId, operation) => ({ operation: { ...operation, actorId: 'actor', revision: 1, committedAt: new Date().toISOString(), mutations: operation.mutations.map((mutation) => ({ ...mutation, affectedRanges: [] })) } }),
      checkpointWorkbook: async () => ({ created: true, snapshot: { snapshot: structuredClone(remoteSnapshot!), revision: 1 } }),
    };
    const catalog = new WorkbookCatalogService({
      persistence: new WorkspacePersistence({ databaseName: 'catalog-sync', indexedDB: null }),
      remote,
    });
    const local = await catalog.create({ destination: 'local', snapshot: createTemplateSnapshot('blank', 'offline-unit') });
    const synced = await catalog.syncToServer(local.unitId);
    assert.equal(syncedUnitId, local.unitId);
    assert.equal(synced.entry.storage, 'remote');
  });
});

describe('Workbook catalog state', () => {
  it('keeps failure/conflict states ahead of offline and synced entries', () => {
    assert.ok(WORKBOOK_SYNC_STATE_PRIORITY.error < WORKBOOK_SYNC_STATE_PRIORITY.conflict);
    assert.ok(WORKBOOK_SYNC_STATE_PRIORITY.conflict < WORKBOOK_SYNC_STATE_PRIORITY.offline);
    const entries = [
      { unitId: 'synced', name: 'Synced', revision: 1, updatedAt: '2026-01-01', storage: 'remote' as const, syncState: 'synced' as const, role: 'owner' as const, lifecycle: 'active' as const, source: 'native' as const, locationPath: [], favorite: false, pendingOperationCount: 0 },
      { unitId: 'error', name: 'Error', revision: 1, updatedAt: '2026-01-02', storage: 'remote' as const, syncState: 'error' as const, role: 'owner' as const, lifecycle: 'active' as const, source: 'native' as const, locationPath: [], favorite: false, pendingOperationCount: 0 },
    ];
    assert.deepEqual(filterWorkbookCatalog(entries).map((entry) => entry.unitId), ['error', 'synced']);
  });

  it('namespaces block storage by workbook while keeping the manifest source id stable', async () => {
    const bytesA = new TextEncoder().encode('workbook-a').buffer;
    const bytesB = new TextEncoder().encode('workbook-b').buffer;
    const checksumA = await computeBinaryChecksum(bytesA);
    const checksumB = await computeBinaryChecksum(bytesB);
    const ref = (checksum: string) => ({
      id: 'block-1', dataSourceId: 'source-1', startRow: 0, rowCount: 1,
      storageKey: 'source-1:block-1', checksum, byteLength: 10, encoding: 'columnar-v1' as const, revision: 1,
    });
    const first = new LocalDataBlockStore({ databaseName: 'catalog-blocks', indexedDB: null, unitId: 'unit-a' });
    const second = new LocalDataBlockStore({ databaseName: 'catalog-blocks', indexedDB: null, unitId: 'unit-b' });
    await first.put(ref(checksumA), bytesA);
    await second.put(ref(checksumB), bytesB);
    assert.deepEqual(new Uint8Array((await first.get(ref(checksumA)))!.bytes), new Uint8Array(bytesA));
    assert.deepEqual(new Uint8Array((await second.get(ref(checksumB)))!.bytes), new Uint8Array(bytesB));
  });

  it('keeps sparse overlays isolated by workbook namespace', async () => {
    const overlay = { schema: 'SparseCellOverlayMetadata' as const, revision: 1, cells: [{ row: 0, column: 0, formula: '=1' }] };
    const first = new LocalSparseOverlayStore({ databaseName: 'catalog-overlays', indexedDB: null, unitId: 'unit-a' });
    const second = new LocalSparseOverlayStore({ databaseName: 'catalog-overlays', indexedDB: null, unitId: 'unit-b' });
    await first.put('source-1', 'block-1', overlay);
    assert.ok(await first.get('source-1', 'block-1', 1));
    assert.equal(await second.get('source-1', 'block-1', 1), null);
  });
});
