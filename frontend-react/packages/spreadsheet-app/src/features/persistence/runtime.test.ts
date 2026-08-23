import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { createSpreadsheetRuntime, startCollaborationSession, startPersistenceSession } from '../../runtime';
import { LocalWorkspaceStore } from './storage';

describe('local workspace runtime persistence', () => {
  it('restores local-only workspaces without invoking API or collaboration transport', async () => {
    const runtime = createSpreadsheetRuntime({
      localOnly: true,
      persistence: { databaseName: 'runtime-local-only-test' },
    });
    let apiCalls = 0;
    runtime.api = {
      getSnapshot: async () => { apiCalls += 1; throw new Error('API must not be called'); },
      createWorkbook: async () => { apiCalls += 1; throw new Error('API must not be called'); },
    } as unknown as typeof runtime.api;
    const disposePersistence = startPersistenceSession(runtime);
    const disposeCollaboration = startCollaborationSession(runtime, () => 'sheet-1:0:0');
    await runtime.persistenceReady;
    disposeCollaboration();
    disposePersistence();
    assert.equal(apiCalls, 0);
    assert.equal(runtime.localOnly, true);
    assert.equal(runtime.remoteConnected, false);
    assert.equal(runtime.collab, null);
    assert.equal(runtime.workspaceRecord?.schema, 'WorkspaceRecord');
    assert.equal(runtime.workspaceRecord?.syncMode, 'local-only');
  });

  it('checkpoints the canonical snapshot after each root transaction', async () => {
    const runtime = createSpreadsheetRuntime({
      localOnly: true,
      persistence: { databaseName: 'runtime-checkpoint-test' },
    });
    const dispose = startPersistenceSession(runtime);
    await runtime.persistenceReady;
    const before = runtime.localRevision;
    runtime.collaboration = null;
    runtime.commands.execute('sheet.cell.set', {
      sheetId: runtime.model.activeSheetId,
      row: 0,
      column: 0,
      value: { value: 'checkpointed' },
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    dispose();
    assert.equal(runtime.localRevision, before + 1);
    assert.equal(runtime.workspaceRecord?.snapshot.schema, 'WorkbookSnapshot');
    assert.equal(runtime.workspaceRecord?.snapshot.sheets[0]?.cells['0']?.['0']?.value, 'checkpointed');
  });

  it('uses a stored local-only record before any network path', async () => {
    const databaseName = 'runtime-local-first-test';
    const snapshot = new WorkbookModel('wb-server-default', 'Stored locally').snapshot();
    const store = new LocalWorkspaceStore({ databaseName });
    await store.create({
      unitId: snapshot.unitId,
      snapshot,
      localRevision: 8,
      serverRevision: 3,
      syncMode: 'local-only',
      operations: [],
      nextClientSequence: 0,
    });
    const runtime = createSpreadsheetRuntime({ persistence: { databaseName } });
    let apiCalls = 0;
    runtime.api = {
      getSnapshot: async () => { apiCalls += 1; throw new Error('API must not be called'); },
      createWorkbook: async () => { apiCalls += 1; throw new Error('API must not be called'); },
    } as unknown as typeof runtime.api;
    const dispose = startPersistenceSession(runtime);
    await runtime.persistenceReady;
    dispose();
    assert.equal(apiCalls, 0);
    assert.equal(runtime.workspaceRecord?.localRevision, 8);
    assert.equal(runtime.workspaceRecord?.syncMode, 'local-only');
    assert.equal(runtime.localOnly, true);
  });
});
