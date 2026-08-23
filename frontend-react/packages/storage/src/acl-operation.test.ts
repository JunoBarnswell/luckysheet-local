import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { WorkbookModel } from '@react-sheets/core-model';
import { StorageAccessError, WorkbookStorage } from './index';

test('WorkbookStorage persists ACL and commits a server-owned OperationEnvelopeV2', () => {
  const directory = mkdtempSync(join(tmpdir(), 'react-sheets-storage-'));
  const storage = new WorkbookStorage({ databasePath: join(directory, 'workbook.sqlite') });
  try {
    const snapshot = new WorkbookModel('unit-acl-test', 'ACL test').snapshot();
    const sheetId = snapshot.sheets[0]!.id;
    storage.createWorkbook(snapshot, 'oidc-owner');
    storage.grantAccess('unit-acl-test', 'oidc-owner', 'oidc-editor', 'editor');

    assert.equal(storage.getRole('unit-acl-test', 'oidc-editor'), 'editor');
    assert.equal(storage.listWorkbooks('oidc-editor').length, 1);
    assert.throws(() => storage.getSnapshot('unit-acl-test', 'unrelated-subject'), StorageAccessError);

    const result = storage.appendOperation({
      schema: 'OperationEnvelopeV2',
      operationId: 'operation-acl-test',
      unitId: 'unit-acl-test',
      clientSequence: 1,
      baseRevision: 0,
      mutations: [{
        id: 'cell.set',
        sheetId,
        params: { sheetId, row: 0, column: 0, value: { value: 'server-authorized' } },
      }],
      createdAt: new Date().toISOString(),
    }, 'oidc-editor');

    assert.equal(result.revision, 1);
    assert.equal(result.operation.actorId, 'oidc-editor');
    assert.equal(result.operation.mutations[0]?.affectedRanges[0]?.sheetId, sheetId);
    assert.equal(storage.getSnapshot('unit-acl-test', 'oidc-editor').snapshot.sheets[0]!.cells['0']?.['0']?.value, 'server-authorized');

    assert.throws(() => storage.appendOperation({
      schema: 'OperationEnvelopeV2',
      operationId: 'operation-spoof-test',
      unitId: 'unit-acl-test',
      clientSequence: 2,
      baseRevision: 1,
      mutations: [{ id: 'cell.set', sheetId, params: { actorId: 'spoofed', affectedRanges: [] } }],
      createdAt: new Date().toISOString(),
      actorId: 'spoofed',
    } as never, 'oidc-editor'), /server-owned/);
  } finally {
    storage.close();
    rmSync(directory, { recursive: true, force: true });
  }
});
