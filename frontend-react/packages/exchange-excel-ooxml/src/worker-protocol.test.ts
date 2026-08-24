import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  assertXlsxWorkerRequest,
  assertXlsxWorkerResult,
  createXlsxImportRequest,
  XLSX_WORKER_PROTOCOL,
  XLSX_WORKER_VERSION,
} from './worker-protocol';

describe('xlsx worker protocol', () => {
  it('accepts a strict transferable import request and progress result', () => {
    const request = createXlsxImportRequest('task-1', 4, {
      fileName: 'book.xlsx',
      buffer: new ArrayBuffer(0),
      options: { compatibilityTarget: 'B', preserveMacros: false },
    });
    assertXlsxWorkerRequest(request);
    assertXlsxWorkerResult({
      protocol: XLSX_WORKER_PROTOCOL,
      version: XLSX_WORKER_VERSION,
      taskId: 'task-1',
      revision: 4,
      status: 'progress',
      stage: 'read',
      percent: 10,
    });
  });

  it('rejects an invalid protocol version or revision', () => {
    assert.throws(() => assertXlsxWorkerRequest({ protocol: XLSX_WORKER_PROTOCOL, version: 99, kind: 'import', taskId: 'task-1', revision: 0, payload: {} }), /protocol or version/);
    assert.throws(() => assertXlsxWorkerResult({ protocol: XLSX_WORKER_PROTOCOL, version: XLSX_WORKER_VERSION, taskId: 'task-1', revision: -1, status: 'cancelled' }), /envelope/);
  });
});
