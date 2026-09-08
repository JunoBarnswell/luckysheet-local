import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkspaceMemoryCoordinator } from '../persistence';
import { LocalDataBlockStore } from '../persistence';
import { computeBinaryChecksum } from '../persistence/checksum';
import { LocalSparseOverlayStore } from '../data-source/overlay-store';
import { filterWorkbookCatalog, WORKBOOK_SYNC_STATE_PRIORITY } from './state';
import { WorkbookCatalogError, WorkbookCatalogService } from './service';
import { WorkbookResolutionError } from './resolver';
import type { WorkbookCatalogRemoteClient } from './types';

const manifest = {
  schema: 'KernelReplicaManifest' as const,
  version: 11 as const,
  unitId: 'unit-1',
  name: 'Cloud workbook',
  revision: 7,
  sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 1000, columnCount: 26, metadata: {} }],
  pages: [],
  metadata: {},
};

function remote(overrides: Partial<WorkbookCatalogRemoteClient> = {}): WorkbookCatalogRemoteClient {
  return {
    getManifest: async () => structuredClone(manifest),
    getAccess: async () => ({ unitId: manifest.unitId, role: 'owner' }),
    getWorkbookUserState: async () => ({ unitId: manifest.unitId, favorite: false }),
    putWorkbookUserState: async (_unitId, state) => ({ unitId: manifest.unitId, ...state }),
    listWorkbookPage: async () => ({ items: [{ unitId: manifest.unitId, name: manifest.name, revision: manifest.revision, updatedAt: '2026-01-01T00:00:00.000Z', role: 'owner', syncStatus: 'synced' }], nextCursor: null }),
    ...overrides,
  } as WorkbookCatalogRemoteClient;
}

describe('cloud-only workbook catalog', () => {
  it('resolves an authoritative manifest and records MRU only after ready', async () => {
    const savedStates: Array<{ lastOpenedAt?: string }> = [];
    const catalog = new WorkbookCatalogService({
      remote: remote({ putWorkbookUserState: async (_unitId, state) => (savedStates.push(state), { unitId: manifest.unitId, ...state }) }),
      now: () => new Date('2026-08-26T00:01:00.000Z'),
    });
    const resolution = await catalog.resolve(manifest.unitId);
    assert.equal(resolution.mode, 'remote');
    assert.equal(resolution.revision, 7);
    assert.equal(savedStates.length, 0);
    await catalog.markOpened(resolution);
    assert.equal(savedStates[0]?.lastOpenedAt, '2026-08-26T00:01:00.000Z');
  });

  it('fails closed when cloud authority is unavailable', async () => {
    const catalog = new WorkbookCatalogService({ remote: remote(), remoteAvailable: () => false });
    await assert.rejects(
      () => catalog.resolve(manifest.unitId),
      (error: unknown) => error instanceof WorkbookResolutionError && error.code === 'remote-unavailable',
    );
  });

  it('rejects obsolete pending/offline catalog states from the server', async () => {
    const catalog = new WorkbookCatalogService({
      remote: remote({ listWorkbookPage: async () => ({ items: [{ unitId: manifest.unitId, name: manifest.name, revision: 7, updatedAt: '', syncStatus: 'pending' }], nextCursor: null }) as never }),
    });
    await assert.rejects(() => catalog.list(), (error: unknown) => error instanceof WorkbookCatalogError && error.code === 'conflict');
  });
});

describe('catalog projection and browser caches', () => {
  it('orders actionable cloud failures ahead of saved workbooks', () => {
    assert.ok(WORKBOOK_SYNC_STATE_PRIORITY.error < WORKBOOK_SYNC_STATE_PRIORITY.synced);
    const base = { revision: 1, role: 'owner' as const, lifecycle: 'active' as const, source: 'native' as const, locationPath: [], favorite: false };
    const entries = [
      { ...base, unitId: 'saved', name: 'Saved', updatedAt: '2026-01-01', syncState: 'synced' as const },
      { ...base, unitId: 'error', name: 'Error', updatedAt: '2026-01-02', syncState: 'error' as const },
    ];
    assert.deepEqual(filterWorkbookCatalog(entries).map((entry) => entry.unitId), ['error', 'saved']);
  });

  it('namespaces block and sparse overlay caches by workbook', async () => {
    const bytes = new TextEncoder().encode('unit-a').buffer;
    const checksum = await computeBinaryChecksum(bytes);
    const ref = { id: 'block-1', dataSourceId: 'source-1', startRow: 0, rowCount: 1, storageKey: 'source-1:block-1', checksum, byteLength: bytes.byteLength, encoding: 'columnar-v1' as const, revision: 1 };
    const coordinator = new WorkspaceMemoryCoordinator();
    const first = new LocalDataBlockStore(coordinator, 'unit-a');
    const second = new LocalDataBlockStore(coordinator, 'unit-b');
    await first.put(ref, bytes);
    assert.ok(await first.get(ref));
    assert.equal(await second.get(ref), null);
    const overlay = { schema: 'SparseCellOverlayMetadata' as const, revision: 1, cells: [{ row: 0, column: 0, formula: '=1' }] };
    const firstOverlay = new LocalSparseOverlayStore({ coordinator, unitId: 'unit-a' });
    const secondOverlay = new LocalSparseOverlayStore({ coordinator, unitId: 'unit-b' });
    await firstOverlay.put('source-1', 'block-1', overlay);
    assert.ok(await firstOverlay.get('source-1', 'block-1', 1));
    assert.equal(await secondOverlay.get('source-1', 'block-1', 1), null);
  });
});
