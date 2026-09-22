import assert from 'node:assert/strict';
import test from 'node:test';
import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';
import { LocalDataBlockStore } from './data-block-store';
import { WorkspaceMemoryCoordinator } from './memory';

async function block(text: string): Promise<{ ref: DataBlockRef; bytes: ArrayBuffer }> {
  const bytes = new TextEncoder().encode(text).buffer as ArrayBuffer;
  return {
    bytes,
    ref: { id: 'block', dataSourceId: 'source', startRow: 0, rowCount: 1,
      storageKey: 'source/block', checksum: await computeBinaryChecksum(bytes),
      byteLength: bytes.byteLength, encoding: 'columnar-v1', revision: 0 },
  };
}

test('block writes are idempotent but cannot replace bytes at an existing identity', async () => {
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator(), 'workbook');
  const original = await block('original');
  const replacement = await block('modified');
  const first = await store.put(original.ref, original.bytes);
  assert.deepEqual(await store.put(original.ref, original.bytes), first);
  await assert.rejects(store.put(replacement.ref, replacement.bytes), { code: 'STORAGE_REVISION_CONFLICT' });
  assert.deepEqual((await store.get(original.ref))?.bytes, original.bytes);
});

test('block writes reject a mismatched manifest length without storing partial content', async () => {
  const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator(), 'workbook');
  const original = await block('original');
  await assert.rejects(store.put({ ...original.ref, byteLength: 1 }, original.bytes), { code: 'STORAGE_SCHEMA_INVALID' });
  assert.equal(await store.get(original.ref), null);
  await store.put(original.ref, original.bytes);
  assert.deepEqual((await store.get(original.ref))?.bytes, original.bytes);
});
