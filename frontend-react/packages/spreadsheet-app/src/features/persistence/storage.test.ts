import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';
import { WorkspacePersistence } from './storage';
import { WorkspaceMemoryCoordinator, WorkspaceStorageError } from './memory';

describe('browser cache persistence', () => {
  it('coordinates data blocks and sparse overlays without storing workbook snapshots', async () => {
    const persistence = new WorkspacePersistence({ unitId: 'cache-workbook' });
    const bytes = new TextEncoder().encode('cache block').buffer;
    const checksum = await computeBinaryChecksum(bytes);
    const ref: DataBlockRef = {
      id: 'block-0',
      dataSourceId: 'source-1',
      startRow: 0,
      rowCount: 1,
      storageKey: 'source-1/block-0',
      checksum,
      byteLength: bytes.byteLength,
      encoding: 'columnar-v1',
      revision: 1,
    };

    await persistence.dataBlocks.put(ref, bytes);
    const cached = await persistence.dataBlocks.get(ref);
    assert.equal(cached?.schema, 'DataBlockRecord');
    assert.equal(cached?.sourceId, 'unit:cache-workbook:source:source-1');
    assert.equal(cached?.checksum, checksum);
    assert.deepEqual(cached?.bytes, bytes);
    await persistence.sparseOverlays.put('source-1', 'block-0', {
      schema: 'SparseCellOverlayMetadata',
      revision: 1,
      cells: [{ row: 0, column: 0, formula: '=1+1' }],
    });
    assert.equal((await persistence.sparseOverlays.get('source-1', 'block-0', 1))?.overlay.cells[0]?.formula, '=1+1');
  });

  it('rolls back failed cache transactions and rejects access after disposal', async () => {
    const coordinator = new WorkspaceMemoryCoordinator();
    await coordinator.transaction((transaction) => transaction.set('dataBlocks', 'block', { value: 'before' }));
    await assert.rejects(coordinator.transaction((transaction) => {
      transaction.set('dataBlocks', 'block', { value: 'after' });
      transaction.set('assets', 'asset', { value: 'after' });
      throw new Error('forced transaction failure');
    }));
    assert.deepEqual(await coordinator.read((transaction) => transaction.get('dataBlocks', 'block')), { value: 'before' });
    assert.equal(await coordinator.read((transaction) => transaction.get('assets', 'asset')), undefined);

    await coordinator.disposeAsync();
    await assert.rejects(
      coordinator.read((transaction) => transaction.get('dataBlocks', 'block')),
      (error: unknown) => error instanceof WorkspaceStorageError && error.code === 'STORAGE_MEMORY_DISPOSED',
    );
  });
});
