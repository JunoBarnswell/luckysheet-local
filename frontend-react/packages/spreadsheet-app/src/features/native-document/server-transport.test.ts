import assert from 'node:assert/strict';
import test from 'node:test';
import { createHash } from 'node:crypto';
import { WorkbookApiNativeDocumentTransport } from './server-transport';
import type { WorkbookSourceArtifactMetadata } from '@react-sheets/protocol';
import type { NativeDocumentServerApi } from './server-transport';

const bytes = new Uint8Array([80, 75, 3, 4]);
const checksum = createHash('sha256').update(bytes).digest('hex');
const manifest = {
  schema: 'WorkbookManifest' as const,
  version: 11 as const,
  unitId: 'native-unit',
  name: 'Native',
  revision: 0,
  sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 1000, columnCount: 26, metadata: {} }],
  pages: [],
  metadata: {},
};

function metadata(overrides: Partial<WorkbookSourceArtifactMetadata> = {}): WorkbookSourceArtifactMetadata {
  return {
    unitId: 'native-unit',
    fileName: 'native.xlsx',
    mimeType: 'application/octet-stream',
    checksum,
    revision: 0,
    byteLength: bytes.byteLength,
    nativeMetadata: {
      revision: 0,
      checksum,
      byteLength: bytes.byteLength,
      format: 'xlsx',
      codecRevision: 1,
      documentMetadata: {
        dateSystem: 'excel1900',
        features: [
          { feature: 'cells', support: 'edit', reason: 'canonical cells' },
          { feature: 'macros', support: 'preserve', reason: 'preserved package part' },
        ],
      },
    },
    createdAt: '2026-09-07T00:00:00.000Z',
    updatedAt: '2026-09-07T00:00:00.000Z',
    ...overrides,
  };
}

function importApi(artifact: WorkbookSourceArtifactMetadata): NativeDocumentServerApi {
  return {
    createNativeDocumentTask: async (request) => ({ taskId: 'task-1', state: 'uploading', uploadedBytes: 0, byteLength: request.byteLength }),
    uploadNativeDocumentTaskChunk: async (taskId, offset, chunk) => ({ taskId, state: 'uploading', uploadedBytes: offset + chunk.byteLength, byteLength: bytes.byteLength }),
    commitNativeDocumentTask: async (taskId) => ({
      taskId, state: 'completed', uploadedBytes: bytes.byteLength, byteLength: bytes.byteLength,
      result: { unitId: 'native-unit', revision: 0, checksum, summary: {} as never, manifest, artifact },
    }),
    cancelNativeDocumentTask: async (taskId) => ({ taskId, state: 'cancelled', uploadedBytes: bytes.byteLength, byteLength: bytes.byteLength }),
    saveNativeDocumentArtifact: async () => artifact,
    getWorkbookSourceArtifact: async () => ({ artifact: new Blob([bytes]), metadata: artifact }),
  };
}

test('server native transport maps committed import identity and compatibility metadata', async () => {
  const artifact = metadata();
  const transport = new WorkbookApiNativeDocumentTransport(importApi(artifact), { name: 'Native', spaceId: 'space-1', folderId: 'folder-1' });

  const result = await transport.import({
    fileName: 'native.xlsx',
    content: bytes.buffer,
    options: { compatibilityTarget: 'B' },
  });

  assert.equal(result.unitId, 'native-unit');
  assert.equal(result.artifact.sourceRevision, 0);
  assert.equal(result.artifact.dateSystem, '1900');
  assert.deepEqual(result.artifact.detectedFeatures, ['cells', 'macros']);
  assert.equal(result.report.summary.editableFeatures, 1);
  assert.equal(result.report.summary.preservedOnly, 1);
});

test('server native transport verifies downloaded export bytes and rejects checksum drift', async () => {
  const published = metadata({ revision: 3 });
  published.nativeMetadata = { ...(published.nativeMetadata as Record<string, unknown>), revision: 3 };
  const transport = new WorkbookApiNativeDocumentTransport({
    ...importApi(published),
    saveNativeDocumentArtifact: async () => published,
    getWorkbookSourceArtifact: async () => ({ artifact: new Blob([new Uint8Array([1, 2, 3, 4])]), metadata: published }),
  });

  await assert.rejects(() => transport.export({
    unitId: 'native-unit', revision: 3, fileName: 'native.xlsx', options: { compatibilityTarget: 'B' },
  }), /NATIVE_DOCUMENT_EXPORT_CHECKSUM_MISMATCH/);
});

test('server native transport rejects unsupported export options before publishing an artifact', async () => {
  const published = metadata({ revision: 3 });
  published.nativeMetadata = { ...(published.nativeMetadata as Record<string, unknown>), revision: 3 };
  let publishCalls = 0;
  const transport = new WorkbookApiNativeDocumentTransport({
    ...importApi(published),
    saveNativeDocumentArtifact: async () => { publishCalls += 1; return published; },
  });

  await assert.rejects(() => transport.export({
    unitId: 'native-unit', revision: 3, fileName: 'native.xlsx',
    options: { compatibilityTarget: 'B', includeCachedValues: false },
  }), /UNSUPPORTED_FEATURE: native export cannot disable cached formula values/);
  await assert.rejects(() => transport.export({
    unitId: 'native-unit', revision: 3, fileName: 'native.xlsx',
    options: { compatibilityTarget: 'B', preserveMacros: false },
  }), /UNSUPPORTED_FEATURE: native export cannot remove or rewrite macro parts/);
  await assert.rejects(() => transport.export({
    unitId: 'native-unit', revision: 3, fileName: 'native.xlsx',
    options: { compatibilityTarget: 'B', dateSystem: '1904' },
  }), /UNSUPPORTED_FEATURE: native export cannot convert the workbook date system/);
  assert.equal(publishCalls, 0);
});

test('server native transport rejects import metadata whose nested revision disagrees', async () => {
  const invalid = metadata();
  invalid.nativeMetadata = { ...(invalid.nativeMetadata as Record<string, unknown>), revision: 9 };
  const transport = new WorkbookApiNativeDocumentTransport(importApi(invalid));

  await assert.rejects(() => transport.import({
    fileName: 'native.xlsx', content: bytes.buffer, options: { compatibilityTarget: 'B' },
  }), /NATIVE_DOCUMENT_ARTIFACT_METADATA_INVALID/);
});
