import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  assertNativeDocumentWorkerRequest,
  assertNativeDocumentWorkerResult,
  createNativeDocumentImportRequest,
  NATIVE_DOCUMENT_WORKER_PROTOCOL,
  NATIVE_DOCUMENT_WORKER_VERSION,
} from './worker-protocol';
import { consumeNativeDocumentWorkerTask } from './worker-entry';
import { createNativeDocumentExportRequest } from './worker-protocol';

describe('native document worker protocol', () => {
  it('accepts a strict transferable import request and progress result', () => {
    const request = createNativeDocumentImportRequest('task-1', 4, {
      fileName: 'book.xlsx',
      buffer: new ArrayBuffer(0),
      options: { compatibilityTarget: 'B', preserveMacros: false },
    });
    assertNativeDocumentWorkerRequest(request);
    assertNativeDocumentWorkerResult({
      protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL,
      version: NATIVE_DOCUMENT_WORKER_VERSION,
      taskId: 'task-1',
      revision: 4,
      status: 'progress',
      stage: 'read',
      percent: 10,
    });
  });

  it('rejects an invalid protocol version or revision', () => {
    assert.throws(() => assertNativeDocumentWorkerRequest({ protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: 99, kind: 'import', taskId: 'task-1', revision: 0, payload: {} }), /protocol or version/);
    assert.throws(() => assertNativeDocumentWorkerResult({ protocol: NATIVE_DOCUMENT_WORKER_PROTOCOL, version: NATIVE_DOCUMENT_WORKER_VERSION, taskId: 'task-1', revision: -1, status: 'cancelled' }), /envelope/);
  });

  it('dispatches non-OOXML codecs through the same worker entry', async () => {
    const result = await consumeNativeDocumentWorkerTask(createNativeDocumentImportRequest('task-text', 1, {
      fileName: 'values.csv',
      buffer: new TextEncoder().encode('A,B\n1,2').buffer as ArrayBuffer,
      options: { compatibilityTarget: 'B' },
    }));
    assert.equal(result.status, 'completed');
    if (result.status !== 'completed') return;
    assert.equal('snapshot' in result.result, true);
    if (!('snapshot' in result.result)) return;
    assert.equal(result.result.artifact.nativeGraph.kind, 'text');
    const importedSnapshot = result.result.snapshot;
    const exported = await consumeNativeDocumentWorkerTask(createNativeDocumentExportRequest('task-text-export', 1, {
      fileName: 'values.csv',
      snapshot: importedSnapshot,
      options: { compatibilityTarget: 'B' },
      mode: 'save-as',
    }));
    assert.equal(exported.status, 'completed');
    if (exported.status === 'completed' && 'fileName' in exported.result) assert.equal(exported.result.fileName, 'values.csv');
  });
});
