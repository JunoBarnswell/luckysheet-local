import assert from 'node:assert/strict';
import test from 'node:test';
import { createNativeDocumentTransaction } from './exchange';
import type { NativeDocumentArtifact, NativeDocumentTransport } from '@react-sheets/exchange-excel-ooxml';

const report = {
  schema: 'CompatibilityReport' as const,
  fileName: 'source.xlsx',
  importLevel: 'B' as const,
  exportLevel: 'B' as const,
  dateSystem: '1900' as const,
  issues: [],
  summary: { editableFeatures: 1, preservedOnly: 0, unsupported: 0 },
};
const artifact: NativeDocumentArtifact = {
  schema: 'NativeDocumentArtifact',
  unitId: 'unit-1',
  fileName: 'source.xlsx',
  format: { family: 'ooxml', profile: 'transitional', variant: 'xlsx' },
  checksum: 'a'.repeat(64),
  byteLength: 1,
  sourceRevision: 1,
  dateSystem: '1900',
  detectedFeatures: ['cells'],
  codecRevision: 1,
  compatibility: report,
};
const manifest = {
  schema: 'WorkbookManifest' as const,
  version: 11 as const,
  unitId: 'unit-1',
  name: 'source',
  revision: 1,
  sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 1, columnCount: 1, metadata: {} }],
  pages: [],
  metadata: {},
};

function transport(): NativeDocumentTransport {
  return {
    import: async () => ({ unitId: 'unit-1', manifest, report, artifact }),
    export: async (request) => ({ unitId: request.unitId, revision: request.revision, content: new ArrayBuffer(1), fileName: request.fileName, report, artifact }),
  };
}

test('native transaction imports the server manifest and exports through the transport', async () => {
  const transaction = createNativeDocumentTransaction(transport());
  const imported = await transaction.import({ fileName: 'source.xlsx', content: new ArrayBuffer(1), options: { compatibilityTarget: 'B' } });
  assert.equal(imported.manifest.version, 11);
  assert.equal(transaction.artifact?.checksum, artifact.checksum);
  const exported = await transaction.export({ unitId: 'unit-1', revision: 1, fileName: 'source.xlsx', mode: 'save', options: { compatibilityTarget: 'B' } });
  assert.equal(exported.content?.byteLength, 1);
  assert.equal(transaction.status, 'exported');
});

test('native transaction fails closed when no server transport is supplied', async () => {
  const transaction = createNativeDocumentTransaction();
  await assert.rejects(
    transaction.import({ fileName: 'source.xlsx', content: new ArrayBuffer(1), options: { compatibilityTarget: 'B' } }),
    /NATIVE_DOCUMENT_TRANSPORT_FAILED: NATIVE_DOCUMENT_TRANSPORT_UNAVAILABLE/,
  );
  assert.equal(transaction.status, 'failed');
  assert.equal(transaction.artifact, undefined);
});

test('Save As does not retarget the imported artifact baseline', async () => {
  const transaction = createNativeDocumentTransaction(transport());
  await transaction.import({ fileName: 'source.xlsx', content: new ArrayBuffer(1), options: { compatibilityTarget: 'B' } });
  await transaction.export({ unitId: 'unit-1', revision: 1, fileName: 'copy.xlsx', mode: 'save-as', options: { compatibilityTarget: 'B' } });
  assert.equal(transaction.artifact?.fileName, 'source.xlsx');
  assert.equal(transaction.status, 'imported');
});
