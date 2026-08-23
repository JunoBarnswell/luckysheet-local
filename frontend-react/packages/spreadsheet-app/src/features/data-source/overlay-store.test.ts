import assert from 'node:assert/strict';
import test from 'node:test';
import type { SparseCellOverlayMetadata } from './import';
import {
  LocalSparseOverlayStore,
  MemorySparseOverlayStore,
} from './overlay-store';

function overlay(revision: number, row = 2, column = 1): SparseCellOverlayMetadata {
  return {
    schema: 'SparseCellOverlayMetadata',
    revision,
    cells: [{
      row,
      column,
      formula: '=SUM(A1:A2)',
      style: { bold: true, background: '#fff7ed' },
      comment: {
        id: `comment-${String(revision)}`,
        author: 'user',
        text: 'Preserve metadata',
        createdAt: '2026-08-24T00:00:00.000Z',
      },
    }],
  };
}

test('memory overlay store persists by source, block, and exact revision without aliasing values', async () => {
  const store = new MemorySparseOverlayStore();
  const written = await store.put('source-1', 'block-1', overlay(4));
  assert.equal(written.schema, 'SparseCellOverlayRecord');
  assert.equal(written.revision, 4);
  assert.equal(written.overlay.revision, 4);
  assert.equal((await store.get('source-1', 'block-1', 3)), null);

  written.overlay.cells[0]!.style!.bold = false;
  const loaded = await store.get('source-1', 'block-1', 4);
  assert.ok(loaded);
  assert.equal(loaded.overlay.cells[0]!.style!.bold, true);
  assert.equal(loaded.overlay.cells[0]!.comment?.text, 'Preserve metadata');
});

test('multiple revisions remain independently addressable and can be deleted precisely', async () => {
  const store = new MemorySparseOverlayStore();
  await store.put('source-1', 'block-1', overlay(1, 0, 0));
  await store.put('source-1', 'block-1', overlay(2, 1, 0));
  await store.put('source-1', 'block-2', overlay(1, 0, 1));

  await store.remove('source-1', 'block-1', 1);
  assert.equal(await store.get('source-1', 'block-1', 1), null);
  assert.ok(await store.get('source-1', 'block-1', 2));
  assert.ok(await store.get('source-1', 'block-2', 1));

  await store.removeBlock('source-1', 'block-1');
  assert.equal(await store.get('source-1', 'block-1', 2), null);
  assert.ok(await store.get('source-1', 'block-2', 1));
  await store.removeSource('source-1');
  assert.equal(await store.get('source-1', 'block-2', 1), null);
});

test('overlay writes reject mismatched revisions, duplicate coordinates, and empty metadata cells', async () => {
  const store = new MemorySparseOverlayStore();
  await assert.rejects(
    store.put('source-1', 'block-1', { ...overlay(1), revision: 2 }),
    /schema or revision/,
  );
  await assert.rejects(
    store.put('source-1', 'block-1', {
      ...overlay(1),
      cells: [{ row: 1, column: 1 }, { row: 1, column: 1, formula: '=1' }],
    }),
    /empty metadata cell/,
  );
  await assert.rejects(
    store.put('source-1', 'block-1', {
      ...overlay(1),
      cells: [{ row: 1, column: 1, formula: '=1' }, { row: 1, column: 1, formula: '=2' }],
    }),
    /duplicate cell/,
  );
});

test('local store shares its memory namespace across instances when IndexedDB is absent', async () => {
  const databaseName = `overlay-test-${Date.now()}-${Math.random()}`;
  const first = new LocalSparseOverlayStore({ databaseName, indexedDB: null });
  const second = new LocalSparseOverlayStore({ databaseName, indexedDB: null });
  await first.put('source-2', 'block-1', overlay(7));
  assert.ok(await second.get('source-2', 'block-1', 7));
  await second.removeSource('source-2');
  assert.equal(await first.get('source-2', 'block-1', 7), null);
});
