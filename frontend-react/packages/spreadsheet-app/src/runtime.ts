import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry, type MutationInfo } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  WorkbookApiClient,
  type AuthTokenProvider,
  type OperationMessageV2,
  type SnapshotResponse,
} from '@react-sheets/protocol';
import { CollabSocketClient } from '@react-sheets/protocol';
import { registerSpreadsheetFeatures } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { createDefaultConnectorRegistry, type ConnectorRegistry } from './features/query';
import { createDefaultCapabilityRegistry, type CapabilityRegistry } from './features/extended';
import { CollaborationSession } from './collaboration/collaboration-session';
import { mapPeerCursor, updatePresenceFromPeer } from './collaboration';
import {
  configureFormulaSpillEnvironment,
  configureWorkbookSpillEnvironments,
  syncWorkbookSheetTables,
  syncWorkbookSpills,
} from './formula-spill-sync';
import {
  LocalDraftStore,
  buildLocalDraftRecord,
  isDraftNewerThanServer,
  scheduleDebounced,
} from './features/persistence';

export interface RuntimeHandlers {
  onSaveState?: (state: import('./types').SaveState) => void;
  onNotice?: (message: string) => void;
  onMutationsApplied?: () => void;
  onPhaseChange?: (phase: import('./types').AppPhase) => void;
  onActiveSheetChange?: (sheetId: string) => void;
  onRemoteRevisions?: (revisions: import('@react-sheets/protocol').RevisionRecord[]) => void;
  onCollabStatus?: (status: 'connecting' | 'open' | 'closed') => void;
  onPeersChange?: (peers: import('./types').PeerCursor[]) => void;
  onDraftUpdated?: () => void;
}

export interface SpreadsheetRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  model: WorkbookModel;
  commands: CommandRuntime;
  drawing: DrawingRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
  pendingMutations: MutationInfo[];
  detachers: Array<() => void>;
  handlers: RuntimeHandlers;
  ownOperationIds: Set<string>;
  nextClientSequence: number;
  pivotResults: Record<string, import('@react-sheets/core-model').PivotResultTree>;
  collab: CollabSocketClient | null;
  collabDispose: (() => void) | null;
  collaboration: CollaborationSession | null;
  bootstrapDispose: (() => void) | null;
  draftStore: LocalDraftStore;
  scheduleDraftWrite: () => void;
  connectors: ConnectorRegistry;
  capabilities: CapabilityRegistry;
  authTokenProvider?: AuthTokenProvider;
}

const UNIT_ID_STORAGE_KEY = 'react-sheets:unitId';
const ACTOR_ID_STORAGE_KEY = 'react-sheets:actorId';

export function resolveUnitId(): string {
  if (typeof window === 'undefined') return 'wb-server-default';
  const routeMatch = /^\/workbooks\/([^/]+)\/?$/.exec(window.location.pathname);
  if (routeMatch?.[1]) return decodeURIComponent(routeMatch[1]);
  const existing = window.localStorage.getItem(UNIT_ID_STORAGE_KEY);
  if (existing) return existing;
  const generated = typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID()
    : 'wb-' + Date.now().toString(36);
  window.localStorage.setItem(UNIT_ID_STORAGE_KEY, generated);
  return generated;
}

export function resolveActorId(): string {
  if (typeof window === 'undefined') return 'actor-server';
  const existing = window.localStorage.getItem(ACTOR_ID_STORAGE_KEY);
  if (existing) return existing;
  const generated = typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID().slice(0, 8)
    : 'actor-' + Date.now().toString(36);
  window.localStorage.setItem(ACTOR_ID_STORAGE_KEY, generated);
  return generated;
}

export function createSpreadsheetRuntime(options: { authTokenProvider?: AuthTokenProvider } = {}): SpreadsheetRuntime {
  const model = new WorkbookModel(resolveUnitId(), 'Untitled workbook');
  const commands = new CommandRuntime(model);
  const drawing = new DrawingRuntime();
  const connectors = createDefaultConnectorRegistry();
  const capabilities = createDefaultCapabilityRegistry();
  registerSpreadsheetFeatures(commands, drawing);
  const draftStore = new LocalDraftStore();
  const runtime: SpreadsheetRuntime = {
    api: new WorkbookApiClient({ authTokenProvider: options.authTokenProvider }),
    formula: new FormulaEngine({ defaultSheetId: 'sheet-1' }),
    model,
    commands,
    drawing,
    remoteConnected: false,
    remoteRevision: 0,
    pendingMutations: [],
    detachers: [],
    handlers: {},
    ownOperationIds: new Set(),
    nextClientSequence: 0,
    pivotResults: {},
    collab: null,
    collabDispose: null,
    collaboration: null,
    bootstrapDispose: null,
    draftStore,
    scheduleDraftWrite: () => undefined,
    connectors,
    capabilities,
    authTokenProvider: options.authTokenProvider,
  };
  runtime.scheduleDraftWrite = scheduleDebounced(() => {
    runtime.draftStore.write(buildLocalDraftRecord(runtime.model.snapshot(), runtime.remoteRevision));
    runtime.handlers.onDraftUpdated?.();
  }, 800);
  attachCoreListeners(runtime);
  return runtime;
}

const FORMULA_SYNC_MUTATIONS = new Set([
  'cell.set',
  'cell.restore',
  'range.set',
  'range.clear',
  'range.paste',
  'cells.shifted',
  'cells.shifted.restore',
  'rows.inserted',
  'rows.deleted',
  'columns.inserted',
  'columns.deleted',
  'sheet.rename',
  'sheet.remove',
  'sheet.restore',
  'sheet.add',
  'sheet.duplicated',
  'sheetTable.add',
  'sheetTable.remove',
  'sheetTable.update',
  'table.add',
  'table.remove',
  'name.set',
  'name.remove',
]);

const DIRECT_CELL_WRITE_MUTATIONS = new Set([
  'cell.set',
  'cell.restore',
  'range.set',
  'range.clear',
  'range.paste',
  'cells.shifted',
  'cells.shifted.restore',
]);

function synchronizeManualCellMutation(engine: FormulaEngine, workbook: WorkbookModel, mutation: MutationInfo): void {
  for (const range of mutation.affectedRanges) {
    const sheet = workbook.getSheet(range.sheetId);
    configureFormulaSpillEnvironment(engine, sheet);
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const cell = sheet.cells.get(row, column);
        const address = { sheetId: sheet.id, row, column };
        if (!cell || (cell.formula === undefined && cell.value == null)) engine.clearCell(address);
        else if (cell.formula !== undefined) engine.setFormula(address, cell.formula);
        else engine.setValue(address, cell.value as never);
      }
    }
  }
  syncWorkbookSpills(engine, workbook);
}

/**
 * Rehydrate one stable FormulaEngine instance from the canonical workbook
 * model.  Structural transforms move formula owners as well as references;
 * rebuilding the dependency index here avoids leaving an old owner address
 * behind after rows/columns or bounded shifts.
 */
function synchronizeFormulaEngine(engine: FormulaEngine, workbook: WorkbookModel): void {
  const mode = engine.getRecalculationMode();
  engine.reset();
  engine.setRecalculationMode('manual');
  engine.setDefinedNames(workbook.definedNames);
  configureWorkbookSpillEnvironments(engine, workbook);
  syncWorkbookSheetTables(engine, workbook);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
  engine.setRecalculationMode(mode);
  engine.recalculate();
  syncWorkbookSpills(engine, workbook);
}

function assertNoSpillChildWrite(
  workbook: WorkbookModel,
  mutation: MutationInfo,
): void {
  const sheet = workbook.getSheet(mutation.sheetId);
  for (const range of mutation.affectedRanges) {
    if (range.sheetId !== sheet.id) continue;
    for (const spill of sheet.spillRanges) {
      const startRow = Math.max(range.startRow, spill.range.startRow);
      const endRow = Math.min(range.endRow, spill.range.endRow);
      const startColumn = Math.max(range.startColumn, spill.range.startColumn);
      const endColumn = Math.min(range.endColumn, spill.range.endColumn);
      if (startRow > endRow || startColumn > endColumn) continue;
      const overlapCells = (endRow - startRow + 1) * (endColumn - startColumn + 1);
      const includesAnchor = spill.anchor.row >= startRow
        && spill.anchor.row <= endRow
        && spill.anchor.column >= startColumn
        && spill.anchor.column <= endColumn;
      if (overlapCells - (includesAnchor ? 1 : 0) > 0) {
        throw new Error('Spill cells are read-only');
      }
    }
  }
}

export function attachCoreListeners(runtime: SpreadsheetRuntime): void {
  detachCoreListeners(runtime);

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      // CommandRuntime invokes listeners after the mutation handler.  Throwing
      // here still causes the command transaction to run its inverse, so a
      // direct write into a dynamic-array child cannot leave partial model or
      // formula state behind.  Undo/redo replay is allowed to restore the
      // exact prior snapshot.
      if (source === 'command' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
        assertNoSpillChildWrite(runtime.model, mutation);
      }

      const changedSheet = runtime.model.getSheets().find((sheet) => sheet.id === mutation.sheetId);
      for (const pivot of changedSheet?.pivots ?? []) {
        delete runtime.pivotResults[pivot.id];
        if (
          mutation.id === 'cell.set' ||
          mutation.id === 'cell.restore' ||
          mutation.id === 'range.set' ||
          mutation.id === 'range.clear' ||
          mutation.id === 'range.paste' ||
          mutation.id === 'cells.shifted' ||
          mutation.id === 'cells.shifted.restore' ||
          mutation.id === 'rows.inserted' ||
          mutation.id === 'rows.deleted' ||
          mutation.id === 'columns.inserted' ||
          mutation.id === 'columns.deleted'
        ) {
          pivot.fieldCatalog = undefined;
        }
      }
      if (FORMULA_SYNC_MUTATIONS.has(mutation.id)) {
        if (runtime.formula.getRecalculationMode() === 'manual' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
          synchronizeManualCellMutation(runtime.formula, runtime.model, mutation);
        } else {
          synchronizeFormulaEngine(runtime.formula, runtime.model);
        }
      }
    }),
  );

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (source !== 'command') return;
      runtime.pendingMutations.push({
        id: mutation.id,
        unitId: mutation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: [...mutation.affectedRanges],
      });
      runtime.scheduleDraftWrite();
    }),
  );

  runtime.detachers.push(
    runtime.commands.onCommand((_commandId, _params, result) => {
      if (runtime.commands.activeDepth > 0) return;
      const batch = runtime.pendingMutations;
      runtime.pendingMutations = [];
      runtime.handlers.onMutationsApplied?.();
      if (batch.length === 0) return;
      submitChangeset(runtime, result.operationId, batch);
    }),
  );
}

function detachCoreListeners(runtime: SpreadsheetRuntime): void {
  for (const detach of runtime.detachers) detach();
  runtime.detachers = [];
  runtime.pendingMutations = [];
}

function submitChangeset(
  runtime: SpreadsheetRuntime,
  operationId: string,
  mutations: MutationInfo[],
): void {
  if (!runtime.collaboration) {
    runtime.handlers.onSaveState?.('offline');
    return;
  }
  runtime.ownOperationIds.add(operationId);
  runtime.collaboration.enqueueLocalMutations(mutations, runtime.model.unitId, operationId);
  runtime.handlers.onSaveState?.('saving');
  void runtime.collaboration.offlineQueue.flushAll().then(({ failed }) => {
    if (failed > 0) runtime.handlers.onNotice?.('Some offline changes could not be synced');
  });
}

export function rehydrateFormulaAfterRestore(runtime: SpreadsheetRuntime, revision?: number): void {
  runtime.formula = rebuildFormulaEngine(runtime.model);
  if (revision != null) {
    runtime.remoteRevision = revision;
    runtime.collaboration?.setRevision(revision);
  }
  runtime.pivotResults = {};
}

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.activeSheetId });
  synchronizeFormulaEngine(engine, workbook);
  return engine;
}

export function hydrateRuntime(runtime: SpreadsheetRuntime, response: SnapshotResponse): void {
  const workbook = WorkbookModel.fromSnapshot(response.snapshot);
  detachCoreListeners(runtime);
  runtime.model = workbook;
  runtime.commands = new CommandRuntime(workbook);
  registerSpreadsheetFeatures(runtime.commands, runtime.drawing);
  runtime.formula = rebuildFormulaEngine(workbook);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  runtime.collaboration?.setRevision(response.revision);
  runtime.collaboration?.rebindCommands(runtime.commands);
  runtime.pivotResults = {};
}

export function startCollaborationSession(
  runtime: SpreadsheetRuntime,
  getSelectionKey: () => string,
  authTokenProvider: AuthTokenProvider = runtime.authTokenProvider ?? (() => null),
): () => void {
  if (typeof window === 'undefined') return () => undefined;

  runtime.collaboration = new CollaborationSession(runtime.commands, {
    send: (operation) => {
      return runtime.collab?.send({ type: 'changeset.submit', payload: operation }) ?? false;
    },
  });
  runtime.collaboration.setRevision(runtime.remoteRevision);

  const protocol = window.location.protocol === 'https:' ? 'wss' : 'ws';
  const client = new CollabSocketClient(protocol + '://' + window.location.host + '/api/v1/collab', {
    authTokenProvider,
  });
  runtime.collab = client;

  const applyRemote = (message: OperationMessageV2) => {
    if (message.type === 'snapshot.response') {
      const response = message.payload;
      if (response.snapshot.unitId !== runtime.model.unitId) return;
      hydrateRuntime(runtime, response);
      runtime.remoteRevision = response.revision;
      runtime.collaboration?.setRevision(response.revision);
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
      runtime.handlers.onMutationsApplied?.();
    } else if (message.type === 'revision.created') {
      if (message.payload.unitId !== runtime.model.unitId) return;
      if (runtime.ownOperationIds.has(message.payload.operationId)) return;
      runtime.collaboration?.applyRemote(message.payload);
      runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
      runtime.collaboration?.setRevision(runtime.remoteRevision);
      runtime.handlers.onMutationsApplied?.();
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
    } else if (message.type === 'changeset.ack') {
      runtime.ownOperationIds.add(message.operationId);
      runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
      if (runtime.collaboration) {
        runtime.collaboration.acknowledge(message.operationId, message.revision);
      }
      runtime.handlers.onSaveState?.('saved');
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
    } else if (message.type === 'changeset.reject') {
      runtime.ownOperationIds.delete(message.operationId);
      runtime.collaboration?.reject(message.operationId, message.error);
      runtime.handlers.onSaveState?.('conflict');
      runtime.handlers.onNotice?.(`Change rejected: ${message.error.message}`);
      void runtime.api.getSnapshot(runtime.model.unitId).then((snapshot) => {
        hydrateRuntime(runtime, snapshot);
        runtime.collaboration?.setRevision(snapshot.revision);
        runtime.handlers.onMutationsApplied?.();
      }).catch(() => undefined);
    } else if (message.type === 'cursor.broadcast' || message.type === 'presence.broadcast') {
      if (!message.unitId || message.unitId !== runtime.model.unitId) return;
      if (message.type === 'presence.broadcast' && (message.state as { status?: string } | null)?.status === 'offline') {
        runtime.handlers.onPeersChange?.([]);
        runtime.collaboration?.presence.removeUser(message.actorId);
        return;
      }
      const cursorState = message.state as { row?: number; column?: number; name?: string; sheetId?: string } | null;
      const peer = mapPeerCursor(message.actorId, cursorState, runtime.model.activeSheetId);
      runtime.collaboration?.presence.upsertUser({
        actorId: peer.actorId,
        displayName: peer.name,
        color: peer.color,
      });
      updatePresenceFromPeer(runtime.collaboration!, peer);
      runtime.handlers.onPeersChange?.([peer]);
    }
  };

  const detachMessage = client.onMessage(applyRemote);
  const detachStatus = client.onStatus((status: 'connecting' | 'open' | 'closed') => {
    runtime.handlers.onCollabStatus?.(status);
    runtime.remoteConnected = status !== 'closed';
    runtime.collaboration?.offlineQueue.setOnline(status === 'open');
    if (status === 'closed') runtime.collaboration?.transportClosed();
    if (status === 'closed') runtime.handlers.onSaveState?.('offline');
    else if (status === 'connecting') runtime.handlers.onSaveState?.('syncing');
    if (status === 'open') {
      client.send({ type: 'snapshot.request', unitId: runtime.model.unitId });
      void runtime.collaboration?.offlineQueue.flushAll().then(({ failed }) => {
        if (failed > 0) runtime.handlers.onNotice?.('Some offline changes could not be synced');
      });
    }
  });
  client.open();

  let lastBroadcast = '';
  const broadcastTimer = window.setInterval(() => {
    const key = getSelectionKey();
    if (key === lastBroadcast) return;
    lastBroadcast = key;
    const parts = key.split(':');
    const state = { row: Number(parts[1]), column: Number(parts[2]), sheetId: parts[0] };
    client.send({
      type: 'cursor.updated',
      unitId: runtime.model.unitId,
      state,
    });
  }, 400);

  return () => {
    window.clearInterval(broadcastTimer);
    detachMessage();
    detachStatus();
    client.close();
    runtime.collab = null;
    runtime.collaboration = null;
  };
}

export function startPersistenceSession(runtime: SpreadsheetRuntime): () => void {
  let active = true;
  void (async () => {
    try {
      const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
      const localDraft = runtime.draftStore.read(runtime.model.unitId);
      if (localDraft && isDraftNewerThanServer(localDraft, snapshotResponse.revision)) {
        hydrateRuntime(runtime, { snapshot: localDraft.snapshot, revision: localDraft.revision });
        runtime.handlers.onNotice?.('Recovered newer local draft');
      } else {
        hydrateRuntime(runtime, snapshotResponse);
      }
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => runtime.handlers.onRemoteRevisions?.([]));
      runtime.remoteConnected = true;
      if (!active) return;
      runtime.handlers.onSaveState?.('saved');
      runtime.handlers.onNotice?.('Workbook restored from server');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.activeSheetId);
      runtime.handlers.onMutationsApplied?.();
      runtime.handlers.onDraftUpdated?.();
    } catch {
      try {
        const created = await runtime.api.createWorkbook(runtime.model.snapshot());
        void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => runtime.handlers.onRemoteRevisions?.([]));
        runtime.remoteConnected = true;
        runtime.remoteRevision = Math.max(runtime.remoteRevision, created.revision);
        if (!active) return;
        runtime.handlers.onSaveState?.('saved');
        runtime.handlers.onNotice?.('SQLite sync connected');
        runtime.handlers.onPhaseChange?.('ready');
        runtime.handlers.onMutationsApplied?.();
      } catch {
        const localDraft = runtime.draftStore.read(runtime.model.unitId);
        if (localDraft) {
          hydrateRuntime(runtime, { snapshot: localDraft.snapshot, revision: localDraft.revision });
          runtime.remoteConnected = false;
          if (!active) return;
          runtime.handlers.onSaveState?.('offline');
          runtime.handlers.onNotice?.('Running from local draft');
          runtime.handlers.onPhaseChange?.('ready');
          runtime.handlers.onMutationsApplied?.();
          runtime.handlers.onDraftUpdated?.();
          return;
        }
        if (!active) return;
        runtime.handlers.onSaveState?.('offline');
        runtime.handlers.onNotice?.('Running local in-memory engine');
        runtime.handlers.onPhaseChange?.('ready');
        runtime.handlers.onMutationsApplied?.();
      }
    }
  })();
  return () => {
    active = false;
  };
}

export type { HistoryEntry };
