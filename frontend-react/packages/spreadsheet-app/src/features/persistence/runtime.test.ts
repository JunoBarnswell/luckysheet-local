import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { ApiRequestError } from '@react-sheets/protocol';
import {
  createSpreadsheetRuntime,
  isLocalSnapshotCheckpointDue,
  startCollaborationSession,
  startPersistenceSession,
} from '../../runtime';
import { WorkspacePersistence } from './storage';

describe('local workspace runtime persistence', () => {
  it('schedules a full local snapshot every fifty local revisions', () => {
    assert.equal(isLocalSnapshotCheckpointDue(49, 0), false);
    assert.equal(isLocalSnapshotCheckpointDue(50, 0), true);
    assert.equal(isLocalSnapshotCheckpointDue(75, 50), false);
  });

  it('restores local-only workspaces without invoking API or collaboration transport', async () => {
    const runtime = createSpreadsheetRuntime({
      localOnly: true,
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

  it('checkpoints the canonical snapshot when explicitly requested after a root transaction', async () => {
    const runtime = createSpreadsheetRuntime({
      localOnly: true,
    });
    const dispose = startPersistenceSession(runtime);
    await runtime.persistenceReady;
    const before = runtime.localRevision;
    runtime.collaboration = null;
    const target = { sheetId: runtime.model.primarySheetId, row: 0, column: 0 };
    const value = { value: 'checkpointed' };
    runtime.commands.execute('sheet.cell.set', {
      sheetId: runtime.model.primarySheetId,
      row: 0,
      column: 0,
      value,
      entryIntent: {
        kind: 'direct-entry',
        target,
        candidate: value,
        validationDecision: { status: 'accepted' },
      },
    });
    await runtime.checkpointWorkspace(false);
    dispose();
    assert.equal(runtime.localRevision, before + 1);
    assert.equal(runtime.workspaceRecord?.snapshot.schema, 'WorkbookSnapshot');
    assert.equal(runtime.workspaceRecord?.snapshot.sheets[0]?.cells['0']?.['0']?.value, 'checkpointed');
  });

  it('uses a stored local-only record before any network path', async () => {
    const persistence = new WorkspacePersistence();
    const snapshot = new WorkbookModel('wb-server-default', 'Stored locally').snapshot();
    const store = persistence.store;
    await store.create({
      unitId: snapshot.unitId,
      snapshot,
      localRevision: 8,
      serverRevision: 3,
      syncMode: 'local-only',
      operations: [],
      nextClientSequence: 0,
    });
    const runtime = createSpreadsheetRuntime({ workspacePersistence: persistence });
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

  it('keeps REST data blocks readable after remote hydration before WebSocket open', async () => {
    const snapshot = new WorkbookModel('remote-blocks-before-socket', 'Remote blocks').snapshot();
    const runtime = createSpreadsheetRuntime({
      unitId: snapshot.unitId,
      authTokenProvider: () => 'verified-host-token',
    });
    runtime.api = {
      getSnapshot: async () => ({ unitId: snapshot.unitId, snapshot, revision: 4 }),
      getAccess: async () => ({ unitId: snapshot.unitId, role: 'editor' }),
      listRevisions: async () => [],
    } as unknown as typeof runtime.api;

    const dispose = startPersistenceSession(runtime);
    await runtime.persistenceReady;
    dispose();

    assert.equal(runtime.localOnly, false);
    assert.equal(runtime.remoteDataAvailable, true);
    assert.equal(runtime.remoteConnected, false);
  });

  it('starts a usable local workbook when the connected backend returns a server error', async () => {
    const runtime = createSpreadsheetRuntime({
      authTokenProvider: () => 'verified-host-token',
    });
    const phases: string[] = [];
    const saveStates: string[] = [];
    runtime.handlers.onPhaseChange = (phase) => phases.push(phase);
    runtime.handlers.onSaveState = (state) => saveStates.push(state);
    runtime.api = {
      getSnapshot: async () => { throw new ApiRequestError('Backend is unavailable', 503, 'INTERNAL_ERROR'); },
      getAccess: async () => { throw new Error('getAccess must not run after a failed snapshot'); },
      createWorkbook: async () => { throw new Error('startup must not create a remote workbook after a backend error'); },
    } as unknown as typeof runtime.api;

    const dispose = startPersistenceSession(runtime);
    await runtime.persistenceReady;
    dispose();

    assert.equal(runtime.localOnly, true);
    assert.equal(runtime.remoteConnected, false);
    assert.equal(runtime.workspaceRecord?.syncMode, 'local-only');
    assert.deepEqual(phases, ['ready']);
    assert.equal(saveStates.at(-1), 'offline');
  });
});
