import assert from 'node:assert/strict';
import test from 'node:test';
import { createHash } from 'node:crypto';
import type { WorkbookImportResponse, WorkbookSourceArtifactMetadata } from '@react-sheets/protocol';
import { WorkbookApiNativeDocumentTransport } from './server-transport';

const bytes = new Uint8Array([80, 75, 3, 4]);
const checksum = createHash('sha256').update(bytes).digest('hex');

const manifest = {
  schema: 'WorkbookManifest' as const,
  version: 11 as const,
  unitId: 'chunk-unit',
  name: 'Chunk',
  revision: 0,
  sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 1, columnCount: 1, metadata: {} }],
  pages: [],
  metadata: {},
};

const metadata: WorkbookSourceArtifactMetadata = {
  unitId: 'chunk-unit', fileName: 'chunk.xlsx', mimeType: 'application/octet-stream', checksum,
  revision: 0, byteLength: bytes.byteLength, updatedAt: '2026-09-08T00:00:00.000Z',
  nativeMetadata: {
    revision: 0, checksum, byteLength: bytes.byteLength, format: 'xlsx', codecRevision: 1,
    documentMetadata: { dateSystem: 'excel1900', features: [{ feature: 'cells', support: 'edit', reason: 'canonical cells' }] },
  },
};

const result: WorkbookImportResponse = {
  unitId: 'chunk-unit', revision: 0, checksum, artifact: metadata, manifest, summary: {} as never,
};

function task(overrides: Record<string, unknown> = {}) {
  return {
    taskId: 'task-1', state: 'uploading' as const, uploadedBytes: 0, byteLength: bytes.byteLength, ...overrides,
  };
}

test('chunk transport uploads opaque bytes and commits the server result', async () => {
  const uploaded: Array<{ offset: number; bytes: ArrayBuffer }> = [];
  let cancelled = false;
  const transport = new WorkbookApiNativeDocumentTransport({
    createNativeDocumentTask: async (request) => {
      assert.equal(request.byteLength, bytes.byteLength);
      assert.equal(request.sha256, checksum);
      return task();
    },
    uploadNativeDocumentTaskChunk: async (_taskId, offset, chunk) => {
      uploaded.push({ offset, bytes: chunk });
      return task({ uploadedBytes: offset + chunk.byteLength });
    },
    commitNativeDocumentTask: async () => task({ state: 'completed', uploadedBytes: bytes.byteLength, result }),
    cancelNativeDocumentTask: async () => { cancelled = true; return task({ state: 'cancelled' }); },
    saveNativeDocumentArtifact: async () => metadata,
    getWorkbookSourceArtifact: async () => ({ artifact: new Blob([bytes]), metadata }),
  });

  const imported = await transport.import({ fileName: 'chunk.xlsx', content: bytes.buffer, options: { compatibilityTarget: 'B' } });
  assert.equal(imported.unitId, 'chunk-unit');
  assert.deepEqual(uploaded.map((entry) => [entry.offset, entry.bytes.byteLength]), [[0, bytes.byteLength]]);
  assert.equal(cancelled, false);
});

test('chunk transport rejects a server offset contract violation and cancels the task', async () => {
  let cancelled = false;
  const transport = new WorkbookApiNativeDocumentTransport({
    createNativeDocumentTask: async () => task(),
    uploadNativeDocumentTaskChunk: async () => task({ uploadedBytes: 1 }),
    commitNativeDocumentTask: async () => task({ state: 'completed', uploadedBytes: bytes.byteLength, result }),
    cancelNativeDocumentTask: async () => { cancelled = true; return task({ state: 'cancelled' }); },
    saveNativeDocumentArtifact: async () => metadata,
    getWorkbookSourceArtifact: async () => ({ artifact: new Blob([bytes]), metadata }),
  });

  await assert.rejects(
    () => transport.import({ fileName: 'chunk.xlsx', content: bytes.buffer, options: { compatibilityTarget: 'B' } }),
    /NATIVE_DOCUMENT_IMPORT_TASK_INVALID/,
  );
  assert.equal(cancelled, true);
});
