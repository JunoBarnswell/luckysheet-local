import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import type { CommittedOperationEnvelope } from '@react-sheets/protocol';
import { createCellSetMutationParams, createPasteSpecialSpec } from '@react-sheets/sheet-features';
import { registerSpreadsheetFeatures } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { disposeSpreadsheetRuntime, hydrateRuntime, loadHistoryAndReplayPending } from './runtime';
import { createRemoteReadySessionFixture } from './session-test-fixtures';
import { WorkbookSession } from './workbook-session';
import { CollaborationSession } from './collaboration/collaboration-session';
import { classifyMutation } from './collaboration/operation-types';
import { rebaseMutation } from './collaboration/ot-rebase';
import { createSpreadsheetRuntime } from './runtime';

describe('WorkbookSession collaboration integration', () => {
  it('rejects offline structural edits before changing the workbook', async () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    const before = runtime.model.snapshot();

    assert.equal(app.canExecute('sheet.rows.insert'), false);
    assert.throws(
      () => app.runCommand('sheet.rows.insert', { sheetId: runtime.model.primarySheetId, at: 0, count: 1 }),
      /STRUCTURAL_PLANNER_OFFLINE/,
    );
    assert.equal(app.canExecute('pivot.drillDown'), false);
    await assert.rejects(
      app.drillDownPivot('unresolved-pivot', 'Details', [{ sheetId: runtime.model.primarySheetId, row: 0 }]),
      /STRUCTURAL_PLANNER_OFFLINE/,
    );
    assert.deepEqual(runtime.model.snapshot(), before);
    assert.equal(runtime.commands.getHistoryDepth().undo, 0);
  });

  it('allows structural mutation execution only after the remote connection is ready', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    runtime.localOnly = false;
    runtime.remoteConnected = true;

    assert.equal(app.canExecute('sheet.rows.insert'), true);
    const result = app.runCommand('sheet.rows.insert', {
      sheetId: runtime.model.primarySheetId,
      at: 0,
      count: 1,
    });

    assert.equal(result.mutationCount, 1);
    assert.equal(runtime.commands.getHistoryDepth().undo, 1);
  });

  it('exposes collaboration snapshot defaults when session is offline', () => {
    const app = new WorkbookSession();
    const snapshot = app.getCollaborationSnapshot();
    assert.equal(snapshot.pendingCount, 0);
    assert.equal(snapshot.offlineQueueState, 'offline');
    assert.equal(app.getUiSnapshot().pendingChangeSetCount, 0);
  });

  it('publishes edit target/status lifecycle without broadcasting draft characters', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    runtime.collaboration = new CollaborationSession(runtime.commands);
    const broadcasts: unknown[] = [];
    runtime.broadcastPresence = (state) => { broadcasts.push(structuredClone(state)); return true; };
    app.cellEdit.dispatch({ type: 'begin.request', source: 'direct-typing', initialText: '=' });
    app.cellEdit.dispatch({ type: 'text.insert', text: 'SENSITIVE-DRAFT' });
    app.cellEdit.dispatch({ type: 'reference.begin' });
    const active = runtime.collaboration.presence.snapshot().editSessions[0];
    assert.equal(active?.status, 'point');
    assert.equal('draftPreview' in (active ?? {}), false);
    assert.equal(JSON.stringify(broadcasts).includes('SENSITIVE-DRAFT'), false);
    app.cellEdit.dispatch({ type: 'cancel' });
    assert.equal(runtime.collaboration.presence.snapshot().editSessions.length, 0);
    assert.deepEqual((broadcasts.at(-1) as { edit?: unknown }).edit, null);
  });

  it('rebaseMutation shifts cell references after structural row inserts', () => {
    const committed = classifyMutation('rows.inserted', { at: 5, count: 1 }, 'sheet-1', [{
      sheetId: 'sheet-1', startRow: 5, endRow: 5, startColumn: 0, endColumn: 0,
    }]);
    const pending = classifyMutation('cell.set', { row: 9, column: 0 }, 'sheet-1', [{
      sheetId: 'sheet-1', startRow: 9, endRow: 9, startColumn: 0, endColumn: 0,
    }]);
    const { rebased, transformed } = rebaseMutation(pending, committed);
    assert.equal(transformed, true);
    assert.equal(rebased.affectedRanges[0]?.startRow, 10);
    assert.equal((rebased.params as { row: number }).row, 10);
  });

  it('applies remote changesets through CollaborationSession without local undo pollution', () => {
    const workbook = new WorkbookModel('wb-collab', 'Collab');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session',
      operationId: 'remote-op',
      unitId: 'wb-collab',
      actorId: 'actor-2',
      origin: 'client',
      clientSequence: 1,
      baseRevision: 0,
      revision: 1,
      committedAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set',
        sheetId: 'sheet-1',
        params: createCellSetMutationParams(
          workbook.getSheet('sheet-1'),
          { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'remote' } },
          'external-sync',
        ),
        affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      }],
      createdAt: new Date().toISOString(),
    });
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.value, 'remote');
    assert.equal(runtime.undo(), false);
  });

  it('rejects a skipped remote revision before applying its mutations', () => {
    const workbook = new WorkbookModel('wb-collab-gap', 'Collab revision gap');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);

    const skippedRevision: CommittedOperationEnvelope = {
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session',
      operationId: 'remote-op-after-gap',
      unitId: 'wb-collab-gap',
      actorId: 'actor-2',
      origin: 'client',
      clientSequence: 2,
      baseRevision: 1,
      revision: 2,
      committedAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set',
        sheetId: 'sheet-1',
        params: createCellSetMutationParams(
          workbook.getSheet('sheet-1'),
          { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'must-not-apply' } },
          'external-sync',
        ),
        affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      }],
      createdAt: new Date().toISOString(),
    };
    assert.throws(() => session.applyRemote(skippedRevision), /COLLABORATION_REVISION_GAP: expected revision 1, received 2/);
    assert.throws(() => session.loadCommittedHistory([skippedRevision], 0), /COLLABORATION_HISTORY_AHEAD_OF_MODEL/);

    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
    assert.equal(session.getRevision(), 0);
  });

  it('hydrates the latest exact snapshot when commits race the initial snapshot read', async () => {
    const runtime = createSpreadsheetRuntime({ unitId: 'wb-bootstrap-race', localOnly: false });
    try {
      const serverWorkbook = new WorkbookModel('wb-bootstrap-race', 'Concurrent server state');
      const sheetId = serverWorkbook.primarySheetId;
      const mutationParams = createCellSetMutationParams(
        serverWorkbook.getSheet(sheetId),
        { sheetId, row: 4, column: 2, value: { value: 'revision-one' } },
        'external-sync',
      );
      serverWorkbook.getSheet(sheetId).cells.set(4, 2, { value: 'revision-one' });
      const operation = {
        schema: 'OperationEnvelope' as const,
        clientSessionId: 'fixture-session',
        operationId: 'commit-after-snapshot',
        unitId: serverWorkbook.unitId,
        actorId: 'actor-2',
        origin: 'client' as const,
        clientSequence: 1,
        baseRevision: 0,
        revision: 1,
        createdAt: new Date().toISOString(),
        committedAt: new Date().toISOString(),
        mutations: [{
          id: 'cell.set',
          sheetId,
          params: mutationParams,
          affectedRanges: [{ sheetId, startRow: 4, endRow: 4, startColumn: 2, endColumn: 2 }],
        }],
      };
      const snapshot = { unitId: serverWorkbook.unitId, snapshot: serverWorkbook.snapshot(), revision: 1, checksum: 'fixture' };
      let requestedRevision: number | undefined;
      runtime.api = {
        getOperationResult: async () => null,
        listRevisions: async () => [{
          operationId: operation.operationId,
          revision: operation.revision,
          committedAt: operation.committedAt,
          payload: operation,
        }],
        getRevisionSnapshot: async (_unitId: string, revision: number) => {
          requestedRevision = revision;
          return snapshot;
        },
      } as unknown as typeof runtime.api;
      runtime.collaboration = new CollaborationSession(runtime.commands);
      hydrateRuntime(runtime, { ...snapshot, snapshot: new WorkbookModel('wb-bootstrap-race', 'Initial').snapshot(), revision: 0 });

      await loadHistoryAndReplayPending(runtime, 0);

      assert.equal(requestedRevision, 1);
      assert.equal(runtime.remoteRevision, 1);
      assert.equal(runtime.model.getSheet(sheetId).cells.get(4, 2)?.value, 'revision-one');
    } finally {
      disposeSpreadsheetRuntime(runtime);
    }
  });

  it('acknowledges a pending commit already included in the hydrated snapshot without applying its patch twice', async () => {
    const runtime = createSpreadsheetRuntime({ unitId: 'wb-commit-in-snapshot', localOnly: false });
    try {
      const sheetId = runtime.model.primarySheetId;
      const params = { sheetId, at: 0, count: 1 };
      const affectedRanges = [...runtime.commands.registry.getMutationMetadata('rows.inserted').affectedRanges.resolve(params)];
      const mutation = { id: 'rows.inserted', unitId: runtime.model.unitId, sheetId, params, affectedRanges };
      const collaboration = new CollaborationSession(runtime.commands, { clientSessionId: 'snapshot-recovery-session' });
      runtime.collaboration = collaboration;
      const pending = collaboration.enqueueLocalMutations([mutation], runtime.model.unitId, 'committed-before-snapshot');
      const beforeAddress = { sheetId, row: 0, column: 0 };
      const afterAddress = { sheetId, row: 1, column: 0 };
      const committed = {
        ...pending,
        actorId: 'actor-1',
        origin: 'client' as const,
        revision: 1,
        committedAt: new Date().toISOString(),
        mutations: [{
          ...pending.mutations[0]!,
          structuralImpactRanges: [beforeAddress, afterAddress].map((address) => ({
            sheetId,
            startRow: address.row,
            endRow: address.row,
            startColumn: address.column,
            endColumn: address.column,
          })),
          structuralPatch: {
            version: 9 as const,
            mutationId: 'rows.inserted',
            formulaOwnerDeltas: [{
              kind: 'formula-cell' as const,
              beforeAddress,
              afterAddress,
              before: { formula: '=A1', sourceFormula: null, barcodeFormula: null },
              after: { formula: '=A2', sourceFormula: null, barcodeFormula: null },
            }],
            definedNameOwnerDeltas: [],
            rangeOwnerDeltas: [],
          },
        }],
      };
      const authoritative = new WorkbookModel(runtime.model.unitId, 'Authoritative snapshot');
      authoritative.getSheet(sheetId).cells.set(1, 0, { value: null, formula: '=A2' });
      const snapshot = { unitId: runtime.model.unitId, snapshot: authoritative.snapshot(), revision: 1, checksum: 'committed' };
      runtime.api = {
        getOperationResult: async () => ({ operation: committed }),
        checkpointWorkbook: async () => ({ revision: 1 }),
        listRevisions: async () => [{
          operationId: committed.operationId,
          revision: committed.revision,
          committedAt: committed.committedAt,
          payload: committed,
        }],
      } as unknown as typeof runtime.api;
      hydrateRuntime(runtime, snapshot);

      await loadHistoryAndReplayPending(runtime, snapshot.revision);

      assert.equal(runtime.model.getSheet(sheetId).cells.get(1, 0)?.formula, '=A2');
      assert.equal(collaboration.offlineQueue.getPendingCount(), 0);
      assert.equal(collaboration.getRevision(), 1);
      assert.equal(runtime.commands.isMutationRecoveryRequired, false);
    } finally {
      disposeSpreadsheetRuntime(runtime);
    }
  });

  it('replays an authoritative sheet-rename patch without rebuilding the formula engine', async () => {
    const app = createRemoteReadySessionFixture();
    try {
      const runtime = app['runtime'];
      const source = runtime.model.getSheet(runtime.model.primarySheetId);
      source.name = 'Source';
      source.cells.set(0, 0, { value: 42 });
      const owner = runtime.model.addSheet('formula-owner', 'Formula Owner');
      const beforeFormula = "='Source'!A1";
      const afterFormula = "='Renamed'!A1";
      owner.cells.set(0, 0, { value: null, formula: beforeFormula });
      const sourceSheetId = source.id;
      const ownerSheetId = owner.id;
      hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
      await app.waitForFormulaCalculation();
      const formulaEngine = runtime.formula;
      const collaboration = new CollaborationSession(runtime.commands, { clientSessionId: 'fixture-session' });
      runtime.collaboration = collaboration;
      const hydratedSource = runtime.model.getSheet(sourceSheetId);
      const hydratedOwner = runtime.model.getSheet(ownerSheetId);

      const ownerAddress = { sheetId: hydratedOwner.id, row: 0, column: 0 };
      const formulaOwnerDelta = {
        kind: 'formula-cell' as const,
        beforeAddress: ownerAddress,
        afterAddress: ownerAddress,
        before: { formula: beforeFormula, sourceFormula: null, barcodeFormula: null },
        after: { formula: afterFormula, sourceFormula: null, barcodeFormula: null },
      };
      collaboration.applyRemote({
        schema: 'OperationEnvelope',
        clientSessionId: 'fixture-session',
        operationId: 'remote-sheet-rename',
        unitId: runtime.model.unitId,
        actorId: 'actor-2',
        origin: 'client',
        clientSequence: 1,
        baseRevision: 0,
        revision: 1,
        createdAt: new Date().toISOString(),
        committedAt: new Date().toISOString(),
        mutations: [{
          id: 'sheet.rename',
          sheetId: hydratedSource.id,
          params: { sheetId: hydratedSource.id, name: 'Renamed' },
          affectedRanges: [],
          structuralImpactRanges: [{
            sheetId: hydratedOwner.id,
            startRow: 0,
            endRow: 0,
            startColumn: 0,
            endColumn: 0,
          }],
          structuralPatch: {
            version: 9,
            mutationId: 'sheet.rename',
            formulaOwnerDeltas: [formulaOwnerDelta],
            definedNameOwnerDeltas: [],
            rangeOwnerDeltas: [],
          },
        }],
      });
      await app.waitForFormulaCalculation();

      assert.equal(hydratedSource.name, 'Renamed');
      assert.equal(hydratedOwner.cells.get(0, 0)?.formula, afterFormula);
      assert.equal(runtime.formula, formulaEngine);
      assert.equal(runtime.formula.getCellResult({ sheetId: hydratedOwner.id, row: 0, column: 0 })?.value, 42);
      assert.equal(runtime.commands.getHistoryDepth().undo, 0);
    } finally {
      app.dispose();
    }
  });

  it('applies committed structural owner patches before acknowledging local operations', () => {
    const createPendingCommit = (formula: string) => {
      const workbook = new WorkbookModel('wb-structural-ack', 'Structural ACK');
      const sheetId = workbook.primarySheetId;
      const runtime = new CommandRuntime(workbook);
      registerSpreadsheetFeatures(runtime, new DrawingRuntime());
      const session = new CollaborationSession(runtime, { clientSessionId: 'fixture-session' });
      workbook.getSheet(sheetId).cells.set(1, 0, { value: null, formula });

      const params = { sheetId, at: 0, count: 1 };
      const affectedRanges = [...runtime.registry.getMutationMetadata('rows.inserted').affectedRanges.resolve(params)];
      const mutation = { id: 'rows.inserted', unitId: workbook.unitId, sheetId, params, affectedRanges };
      const pending = session.enqueueLocalMutations([mutation], workbook.unitId, 'local-structural-ack');
      const beforeAddress = { sheetId, row: 0, column: 0 };
      const afterAddress = { sheetId, row: 1, column: 0 };
      const formulaOwnerDelta = {
        kind: 'formula-cell' as const,
        beforeAddress,
        afterAddress,
        before: { formula: '=A1', sourceFormula: null, barcodeFormula: null },
        after: { formula: '=A2', sourceFormula: null, barcodeFormula: null },
      };
      const impact = [beforeAddress, afterAddress].map((address) => ({
        sheetId: address.sheetId,
        startRow: address.row,
        endRow: address.row,
        startColumn: address.column,
        endColumn: address.column,
      }));
      const committed = {
        ...pending,
        actorId: 'actor-1',
        origin: 'client' as const,
        revision: 1,
        committedAt: new Date().toISOString(),
        mutations: [{
          ...pending.mutations[0]!,
          affectedRanges,
          structuralImpactRanges: impact,
          structuralPatch: {
            version: 9 as const,
            mutationId: 'rows.inserted',
            formulaOwnerDeltas: [formulaOwnerDelta],
            definedNameOwnerDeltas: [],
            rangeOwnerDeltas: [],
          },
        }],
      };
      return { workbook, sheetId, session, committed };
    };

    const success = createPendingCommit('=A1');
    success.session.applyRemote(success.committed);
    assert.equal(success.workbook.getSheet(success.sheetId).cells.get(1, 0)?.formula, '=A2');
    assert.equal(success.session.offlineQueue.getPendingCount(), 0);
    assert.equal(success.session.getRevision(), 1);

    const rejected = createPendingCommit('=Broken');
    assert.throws(() => rejected.session.applyRemote(rejected.committed), /STRUCTURAL_PATCH_PRECONDITION/);
    assert.equal(rejected.workbook.getSheet(rejected.sheetId).cells.get(1, 0)?.formula, '=Broken');
    assert.equal(rejected.session.offlineQueue.getPendingCount(), 1);
    assert.equal(rejected.session.getRevision(), 0);

    const missingPatch = createPendingCommit('=A1');
    const malformedCommit = {
      ...missingPatch.committed,
      mutations: [{
        ...missingPatch.committed.mutations[0]!,
        structuralPatch: undefined,
        structuralImpactRanges: [],
      }],
    };
    assert.throws(() => missingPatch.session.applyCommittedStructuralPatches(malformedCommit),
      /requires a server-derived StructuralPatch/);
    assert.equal(missingPatch.workbook.getSheet(missingPatch.sheetId).cells.get(1, 0)?.formula, '=A1');
    assert.equal(missingPatch.session.offlineQueue.getPendingCount(), 1);
    assert.equal(missingPatch.session.getRevision(), 0);
  });

  it('invalidates overlapping local undo after a committed remote cell write', () => {
    const workbook = new WorkbookModel('wb-collab-history-overlap', 'Collaboration history overlap');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 2, endRow: 2, startColumn: 3, endColumn: 3 };
    runtime.execute('sheet.cell.set', { sheetId, row: 2, column: 3, value: { value: 'local' } });
    runtime.execute('sheet.cell.set', { sheetId, row: 8, column: 8, value: { value: 'unrelated-local' } });

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-overlap',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'cell.set', sheetId,
        params: createCellSetMutationParams(workbook.getSheet(sheetId), {
          sheetId, row: 2, column: 3, value: { value: 'remote' },
        }, 'external-sync'),
        affectedRanges: [range],
      }],
    });

    assert.equal(workbook.getSheet(sheetId).cells.get(2, 3)?.value, 'remote');
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(sheetId).cells.get(2, 3)?.value, 'remote');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('invalidates local undo after a committed range move without a canonical history transform', () => {
    const workbook = new WorkbookModel('wb-collab-history-move', 'Collaboration history move');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const sourceRange = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
    const targetRange = { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 };
    sheet.cells.set(0, 0, { value: 'source' });
    sheet.cells.set(1, 1, { value: 'previous-target' });
    runtime.execute('sheet.cell.set', { sheetId: sheet.id, row: 1, column: 1, value: { value: 'local' } });

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-range-move',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'range.move', sheetId: sheet.id,
        params: { sheetId: sheet.id, sourceRange, targetOrigin: { row: 1, column: 1 } },
        affectedRanges: [sourceRange, targetRange],
      }],
    });

    assert.equal(sheet.cells.get(1, 1)?.value, 'source');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('rejects structural mutation envelopes whose sheet differs from their target', () => {
    const workbook = new WorkbookModel('wb-structural-scope', 'Structural scope');
    const target = workbook.addSheet('target-sheet', 'Target');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    target.cells.set(4, 0, { value: 'stays-at-row-five' });

    assert.throws(() => runtime.applyRemoteMutations([{
      id: 'rows.inserted',
      unitId: workbook.unitId,
      sheetId: workbook.primarySheetId,
      params: { sheetId: target.id, at: 2, count: 1 },
      affectedRanges: [{ sheetId: target.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }],
    }]), /envelope sheetId differs from its target/);
    assert.equal(target.cells.get(4, 0)?.value, 'stays-at-row-five');
    assert.equal(target.cells.get(5, 0), undefined);
  });

  it('invalidates unscoped history when another unscoped mutation arrives', () => {
    const workbook = new WorkbookModel('wb-defined-name-history', 'Defined name history');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheetId = workbook.primarySheetId;
    runtime.execute('workbook.name.set', { name: 'Rate', formula: '=Sheet1!A1', scope: 'workbook' });
    const session = new CollaborationSession(runtime);

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-name-set',
      unitId: workbook.unitId, actorId: 'actor-2', origin: 'client', clientSequence: 1, baseRevision: 0,
      revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'name.set', sheetId,
        params: { model: { name: 'Rate', formula: '=Sheet1!A2', scope: 'workbook' } },
        affectedRanges: [],
      }],
    });

    assert.equal(workbook.getDefinedName('Rate')?.formula, '=Sheet1!A2');
    assert.equal(runtime.undo(), false);
    assert.equal(runtime.getInvalidHistoryEntries()[0]?.status, 'invalid');
  });

  it('rejects range-paste history with omitted or out-of-footprint affected cells', () => {
    const workbook = new WorkbookModel('wb-paste-footprint', 'Paste footprint');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.range.paste', {
      sheetId,
      targetOrigin: { row: 2, column: 2 },
      clipboard: {
        schema: 'SparseClipboardPayload',
        range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        sourceExtent: { rows: 1, columns: 1 },
        occupiedCells: [{ rowOffset: 0, columnOffset: 0, value: { value: 'paste' } }],
        transfer: 'copy',
        rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
      },
      transfer: 'copy',
      spec: createPasteSpecialSpec({
        formatting: 'none',
        metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false },
      }),
    });
    const mutation = runtime.getUndoEntries().at(-1)?.forwardMutations[0];
    assert.ok(mutation);

    assert.ok(runtime.registry.validateMutationInfo({ ...mutation, affectedRanges: [] })
      .some((issue) => issue.code === 'invalid-affected-ranges'));
    const params = mutation.params as {
      snapshot: { cells: Array<{ row: number; column: number; value?: unknown }> };
      clipboard: { range: { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number }; [key: string]: unknown };
    };
    const escapedSnapshot = {
      ...mutation,
      params: {
        ...params,
        snapshot: { ...params.snapshot, cells: [...params.snapshot.cells, { row: 9, column: 9, value: { value: 'outside' } }] },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(escapedSnapshot)
      .some((issue) => issue.code === 'invalid-params'));
    const outOfBoundsSource = {
      ...mutation,
      params: {
        ...params,
        clipboard: {
          ...params.clipboard,
          range: { sheetId, startRow: 1_048_576, endRow: 1_048_576, startColumn: 0, endColumn: 0 },
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(outOfBoundsSource)
      .some((issue) => issue.code === 'invalid-params'));
    const invalidWidth = {
      ...mutation,
      params: { ...params, snapshot: { ...params.snapshot, columnWidths: [{ column: 2, widthPx: 0 }] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(invalidWidth)
      .some((issue) => issue.code === 'invalid-params'));
    const unexpectedMetadata = {
      ...mutation,
      params: { ...params, snapshot: { ...params.snapshot, validations: [] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(unexpectedMetadata)
      .some((issue) => issue.code === 'invalid-params'));
    const validationSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: true, columnWidths: false, conditionalFormats: false, hyperlinks: false },
    });
    const targetRange = { sheetId, startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 };
    const foreignRule = {
      ...mutation,
      params: {
        ...params,
        spec: validationSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          validations: [{ id: 'foreign-rule', sheetId: 'another-sheet', ranges: [{ sheetId: 'another-sheet', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], type: 'whole', formula1: '1' }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(foreignRule)
      .some((issue) => issue.code === 'invalid-params'));
    const malformedRule = {
      ...mutation,
      params: {
        ...params,
        spec: validationSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          validations: [{ id: 'unknown-rule', sheetId, ranges: [targetRange], type: 'unsupported', formula1: '1' }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(malformedRule)
      .some((issue) => issue.code === 'invalid-params'));

    const widthsSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: false, columnWidths: true, conditionalFormats: false, hyperlinks: false },
    });
    const outOfRangeWidth = {
      ...mutation,
      params: { ...params, spec: widthsSpec, snapshot: { ...params.snapshot, clearMetadataRanges: [targetRange], columnWidths: [{ column: 7, widthPx: 120 }] } },
    };
    assert.ok(runtime.registry.validateMutationInfo(outOfRangeWidth)
      .some((issue) => issue.code === 'invalid-params'));

    const hyperlinksSpec = createPasteSpecialSpec({
      formatting: 'none',
      metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: true },
    });
    const nonCanonicalCellKey = {
      ...mutation,
      params: {
        ...params,
        spec: hyperlinksSpec,
        snapshot: {
          ...params.snapshot,
          clearMetadataRanges: [targetRange],
          hyperlinks: [{ key: '02:2', value: { id: 'link', target: { kind: 'url', url: 'https://example.com' } } }],
        },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(nonCanonicalCellKey)
      .some((issue) => issue.code === 'invalid-params'));

    const overlapRange = { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
    const overlappingMove = {
      ...mutation,
      affectedRanges: [overlapRange, overlapRange],
      params: {
        ...params,
        targetOrigin: { row: 0, column: 0 },
        transfer: 'move',
        clearSource: true,
        sourceRange: overlapRange,
        clipboard: { ...params.clipboard, transfer: 'move' },
        snapshot: { ...params.snapshot, clearRanges: [overlapRange, overlapRange], cells: [] },
      },
    };
    assert.ok(runtime.registry.validateMutationInfo(overlappingMove)
      .some((issue) => issue.code === 'invalid-params'));
  });

  it('replays a canonical bulk row visibility mutation without splitting history semantics', () => {
    const workbook = new WorkbookModel('wb-rows-visibility-replay', 'Rows visibility replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheetId = workbook.primarySheetId;
    const affectedRanges = [1, 2].map((row) => ({ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: 0 }));

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-rows-visibility', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'rows.visibility', sheetId, params: { sheetId, states: [{ row: 1, hidden: true }, { row: 2, hidden: true }] }, affectedRanges }],
    });
    assert.deepEqual([...workbook.getSheet(sheetId).hiddenRows].sort((left, right) => left - right), [1, 2]);
    assert.equal(runtime.undo(), false);
  });

  it('replays canonical font-family mutations and rejects an empty family atomically', () => {
    const workbook = new WorkbookModel('wb-font-replay', 'Font replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const session = new CollaborationSession(runtime);
    const sheet = workbook.getSheet('sheet-1');
    const range = { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };

    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-font', unitId: workbook.unitId, actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'style.set', sheetId: sheet.id, params: { sheetId: sheet.id, range, style: { fontFamily: '  aRiAl  ' } }, affectedRanges: [range] }],
    });
    assert.equal(sheet.cells.get(0, 0)?.style?.fontFamily, 'Arial');
    assert.equal(runtime.undo(), false);

    assert.throws(() => session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-empty-font', unitId: workbook.unitId, actorId: 'actor-3', origin: 'client',
      clientSequence: 2, baseRevision: 1, revision: 2, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{ id: 'style.set', sheetId: sheet.id, params: { sheetId: sheet.id, range, style: { fontFamily: '   ' } }, affectedRanges: [range] }],
    }), /must not be empty/);
    assert.equal(sheet.cells.get(0, 0)?.style?.fontFamily, 'Arial');
  });

  it('replays canonical clear families and conditional-format cropping remotely', () => {
    const workbook = new WorkbookModel('wb-clear-replay', 'Clear Replay');
    const runtime = new CommandRuntime(workbook);
    registerSpreadsheetFeatures(runtime, new DrawingRuntime());
    const sheet = workbook.getSheet('sheet-1');
    sheet.conditionalFormats.push({
      id: 'cf-replay',
      sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 }],
      type: 'highlight',
    });
    const session = new CollaborationSession(runtime);
    session.applyRemote({
      schema: 'OperationEnvelope', clientSessionId: 'fixture-session', operationId: 'remote-clear', unitId: 'wb-clear-replay', actorId: 'actor-2', origin: 'client',
      clientSequence: 1, baseRevision: 0, revision: 1, committedAt: new Date().toISOString(), createdAt: new Date().toISOString(),
      mutations: [{
        id: 'range.clear', sheetId: sheet.id,
        params: { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 }, family: 'formats' },
        affectedRanges: [{ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 }],
      }],
    });
    assert.equal(sheet.conditionalFormats[0]?.ranges.length, 4);
    assert.equal(runtime.undo(), false);
  });

  it('turns local undo into a durable compensating operation', () => {
    const runtime = createSpreadsheetRuntime();
    const sheetId = runtime.model.primarySheetId;
    runtime.commands.execute('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'local' },
    });
    assert.equal(runtime.collaboration?.offlineQueue.getPendingCount(), 1);
    assert.equal(runtime.commands.undo(), true);
    const pending = runtime.collaboration?.offlineQueue.getPending() ?? [];
    assert.equal(pending.length, 2);
    assert.notEqual(pending[0]?.operation.operationId, pending[1]?.operation.operationId);
    assert.equal(runtime.model.getSheet(sheetId).cells.get(0, 0), undefined);
  });
});
