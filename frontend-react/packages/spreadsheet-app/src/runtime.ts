import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry, type MutationInfo } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  ApiRequestError,
  WorkbookApiClient,
  type AuthTokenProvider,
  type ShareTokenProvider,
  type WorkbookAclRole,
  type OperationMessage,
  type SnapshotResponse,
  mutationCapability,
} from '@react-sheets/protocol';
import { CollabSocketClient } from '@react-sheets/protocol';
import { registerSpreadsheetFeatures } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { createDefaultConnectorRegistry, type ConnectorRegistry } from './features/query';
import { FormulaAuditController, registerFormulaAuditCommands } from './features/formula-audit';
import { DataSourceContentQuery, migrateDataRegionCellPatches } from './features/data-source';
import { CollaborationSession } from './collaboration/collaboration-session';
import { mapPeerCursor, updatePresenceFromPeer } from './collaboration';
import {
  configureFormulaSpillEnvironment,
  configureWorkbookSpillEnvironments,
  syncWorkbookSheetTables,
  syncWorkbookSpills,
} from './formula-spill-sync';
import {
  OperationJournalStore,
  WorkspacePersistence,
  DataBlockSynchronizer,
  type IndexedDbWorkspaceStoreOptions,
  type WorkspaceRecord,
} from './features/persistence';

export interface RuntimeHandlers {
  onSaveState?: (state: import('./types').SaveState) => void;
  onNotice?: (message: string) => void;
  onMutationsApplied?: () => void;
  onPhaseChange?: (phase: import('./types').AppPhase) => void;
  onActiveSheetChange?: (sheetId: string) => void;
  onRemoteRevisions?: (revisions: import('@react-sheets/protocol').RevisionRecord[]) => void;
  onCollabStatus?: (status: 'connecting' | 'open' | 'closed') => void;
  onAccessRole?: (role: WorkbookAclRole | null) => void;
  onPeersChange?: (peers: import('./types').PeerCursor[]) => void;
  onWorkspacePersisted?: () => void;
}

export interface SpreadsheetRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  formulaAudit: FormulaAuditController;
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
  operationJournal: OperationJournalStore;
  workspacePersistence: WorkspacePersistence;
  dataBlocks: DataBlockSynchronizer;
  dataContent: Map<string, DataSourceContentQuery>;
  dataContentDetachers: Array<() => void>;
  workspaceRecord: WorkspaceRecord | null;
  localRevision: number;
  localOnly: boolean;
  remoteSyncRequested: boolean;
  formulaCalculation: Promise<void>;
  persistenceReady: Promise<void>;
  pendingLocalOperations: Array<{ operationId: string; mutations: MutationInfo[] }>;
  checkpointWorkspace: (advanceLocalRevision?: boolean) => Promise<void>;
  connectors: ConnectorRegistry;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  /** Runtime lifecycle is explicit so late Worker/IndexedDB callbacks cannot
   * publish into a disposed session. */
  disposed: boolean;
}

let localActorSequence = 0;

/** Test-only default; browser routes must provide unitId through the session factory. */
export function resolveUnitId(): string { return 'wb-local-default'; }

export function resolveActorId(): string {
  if (typeof window === 'undefined') return 'actor-server';
  localActorSequence += 1;
  return typeof crypto !== 'undefined' && crypto.randomUUID
    ? `local-${crypto.randomUUID().slice(0, 8)}`
    : `local-${Date.now().toString(36)}-${localActorSequence}`;
}

export function resolveShareToken(): string | null {
  if (typeof window === 'undefined') return null;
  return new URLSearchParams(window.location.search).get('share')?.trim() || null;
}

export function createSpreadsheetRuntime(options: {
  unitId?: string;
  api?: WorkbookApiClient;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  localOnly?: boolean;
  persistence?: IndexedDbWorkspaceStoreOptions;
  workspacePersistence?: WorkspacePersistence;
} = {}): SpreadsheetRuntime {
  const model = new WorkbookModel(options.unitId ?? resolveUnitId(), 'Untitled workbook');
  const commands = new CommandRuntime(model);
  const drawing = new DrawingRuntime();
  const connectors = createDefaultConnectorRegistry();
  const formula = new FormulaEngine({ defaultSheetId: 'sheet-1' });
  const formulaAudit = new FormulaAuditController(formula);
  registerSpreadsheetFeatures(commands, drawing);
  registerFormulaAuditCommands(commands.registry, formulaAudit);
  const operationJournal = new OperationJournalStore();
  const workspacePersistence = options.workspacePersistence ?? new WorkspacePersistence({
    ...options.persistence,
    unitId: () => runtime?.model.unitId ?? model.unitId,
  }, operationJournal);
  const api = options.api ?? new WorkbookApiClient({ authTokenProvider: options.authTokenProvider, shareTokenProvider: options.shareTokenProvider });
  let runtime!: SpreadsheetRuntime;
  const dataBlocks = new DataBlockSynchronizer(workspacePersistence.dataBlocks, api, {
    unitId: () => runtime.model.unitId,
    isRemoteAvailable: () => !runtime.localOnly && runtime.remoteConnected,
  });
  runtime = {
    api,
    formula,
    formulaAudit,
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
    operationJournal,
    workspacePersistence,
    dataBlocks,
    dataContent: new Map(),
    dataContentDetachers: [],
    workspaceRecord: null,
    localRevision: 0,
    localOnly: options.localOnly ?? (!options.authTokenProvider && !options.shareTokenProvider),
    remoteSyncRequested: Boolean(options.authTokenProvider || options.shareTokenProvider),
    formulaCalculation: Promise.resolve(),
    persistenceReady: Promise.resolve(),
    pendingLocalOperations: [],
    checkpointWorkspace: () => Promise.resolve(),
    connectors,
    authTokenProvider: options.authTokenProvider,
    shareTokenProvider: options.shareTokenProvider,
    disposed: false,
  };
  // The offline journal records operation intent and its client sequence.
  // The same IndexedDB transaction also checkpoints the canonical local
  // workbook snapshot, so a closed browser can resume without any service.
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    loadPending: () => {
      const journal = operationJournal.read(runtime.model.unitId);
      return journal
        ? { operations: journal.operations, nextClientSequence: journal.nextClientSequence }
        : null;
    },
    persistPending: (operations, nextClientSequence) => {
      operationJournal.write(runtime.model.unitId, operations, nextClientSequence);
      runtime.handlers.onWorkspacePersisted?.();
    },
  });
  runtime.checkpointWorkspace = () => checkpointWorkspace(runtime);
  attachCoreListeners(runtime);
  return runtime;
}

const FORMULA_SYNC_MUTATIONS = new Set([
  'cell.set',
  'cell.restore',
  'range.set',
  'range.clear',
  'range.paste',
  'format.painter.applied',
  'style.preset.set',
  'dataRegion.materialize.commit',
  'dataRegion.materialize.restore',
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
        else if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) engine.setFormula(address, cell.formula);
        else if (cell.value != null) engine.setValue(address, cell.value as never);
        else engine.setValue(address, cell.value as never);
      }
    }
  }
}

/**
 * Load only formula inputs from the canonical workbook. The actual evaluation
 * is intentionally scheduled separately through FormulaEngine.recalculateAsync
 * so browser calculation stays in its Worker.
 */
function loadFormulaInputs(engine: FormulaEngine, workbook: WorkbookModel): void {
  const mode = engine.getRecalculationMode();
  engine.cancelCalculation();
  engine.reset();
  engine.setRecalculationMode('manual');
  engine.setDefinedNameModels(workbook.definedNameModels);
  configureWorkbookSpillEnvironments(engine, workbook);
  syncWorkbookSheetTables(engine, workbook);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
  engine.setRecalculationMode(mode);
}

interface FormulaQueueState {
  tail: Promise<void>;
  scheduled: boolean;
  epoch: number;
  force: boolean;
}

const formulaQueueStates = new WeakMap<SpreadsheetRuntime, FormulaQueueState>();

function localFormulaIdleState(runtime: SpreadsheetRuntime): import('./types').SaveState {
  if (runtime.localOnly) return runtime.remoteSyncRequested ? 'offline' : 'saved';
  if (!runtime.remoteConnected) return 'offline';
  return runtime.collaboration?.offlineQueue.getPendingCount() ? 'syncing' : 'saved';
}

/**
 * Coalesce formula input changes into one Worker task. A new mutation cancels
 * the active task and advances the epoch, so a late worker result cannot
 * mutate spills or render projections for an older workbook state.
 */
export function scheduleFormulaRecalculation(runtime: SpreadsheetRuntime, force = false): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  const state = formulaQueueStates.get(runtime) ?? {
    tail: Promise.resolve(),
    scheduled: false,
    epoch: 0,
    force: false,
  } satisfies FormulaQueueState;
  formulaQueueStates.set(runtime, state);
  state.epoch += 1;
  state.force ||= force;
  runtime.formula.cancelCalculation();
  if (state.scheduled) return runtime.formulaCalculation;

  state.scheduled = true;
  state.tail = state.tail
    .catch(() => undefined)
    .then(async () => {
      if (runtime.disposed) return;
      state.scheduled = false;
      const epoch = state.epoch;
      const forceCalculation = state.force;
      state.force = false;
      const engine = runtime.formula;
      const workbook = runtime.model;
      loadFormulaInputs(engine, workbook);
      if (engine.getRecalculationMode() === 'manual' && !forceCalculation) {
        if (!runtime.disposed && epoch === state.epoch && runtime.formula === engine && runtime.model === workbook) {
          runtime.handlers.onMutationsApplied?.();
        }
        return;
      }

      runtime.handlers.onSaveState?.('calculating');
      try {
        await engine.recalculateAsync();
        if (runtime.disposed || epoch !== state.epoch || runtime.formula !== engine || runtime.model !== workbook) return;
        syncWorkbookSpills(engine, workbook);
        void checkpointWorkspace(runtime, false).catch((error: unknown) => {
          runtime.handlers.onSaveState?.('error');
          runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Local formula checkpoint failed');
        });
        runtime.handlers.onMutationsApplied?.();
        runtime.handlers.onSaveState?.(localFormulaIdleState(runtime));
      } catch (error) {
        if (runtime.disposed || epoch !== state.epoch || runtime.formula !== engine || runtime.model !== workbook) return;
        runtime.handlers.onSaveState?.('error');
        runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Formula calculation failed');
      }
    });
  runtime.formulaCalculation = state.tail;
  return state.tail;
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

const checkpointChains = new WeakMap<SpreadsheetRuntime, Promise<void>>();

function checkpointWorkspace(runtime: SpreadsheetRuntime, advanceLocalRevision = true): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  if (advanceLocalRevision) runtime.localRevision += 1;
  const snapshot = runtime.model.snapshot();
  const localRevision = runtime.localRevision;
  const serverRevision = runtime.remoteRevision;
  const syncMode = runtime.localOnly ? 'local-only' as const : 'remote' as const;
  const pendingJournal = runtime.operationJournal.read(runtime.model.unitId);
  const previous = checkpointChains.get(runtime) ?? Promise.resolve();
  const next = previous
    .catch(() => undefined)
    .then(async () => {
      if (runtime.disposed) return;
      const record = await runtime.workspacePersistence.checkpoint(
        snapshot,
        localRevision,
        serverRevision,
        syncMode,
        pendingJournal,
      );
      if (runtime.disposed) return;
      runtime.workspaceRecord = record;
      runtime.handlers.onWorkspacePersisted?.();
    });
  checkpointChains.set(runtime, next);
  return next;
}

export function attachCoreListeners(runtime: SpreadsheetRuntime): void {
  detachCoreListeners(runtime);

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (runtime.disposed) return;
      // CommandRuntime invokes listeners after the mutation handler.  Throwing
      // here still causes the command transaction to run its inverse, so a
      // direct write into a dynamic-array child cannot leave partial model or
      // formula state behind.  Undo/redo replay is allowed to restore the
      // exact prior snapshot.
      if (source === 'command' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
        assertNoSpillChildWrite(runtime.model, mutation);
      }

      const changedSheet = runtime.model.getSheets().find((sheet) => sheet.id === mutation.sheetId);
      if (mutation.id === 'dataSource.add' || mutation.id === 'dataSource.update' || mutation.id === 'dataSource.remove'
        || mutation.id === 'dataRegion.add' || mutation.id === 'dataRegion.remove'
        || mutation.id === 'dataRegion.materialize.commit' || mutation.id === 'dataRegion.materialize.restore') {
        initializeDataContent(runtime);
      }
      for (const pivot of changedSheet?.pivots ?? []) {
        delete runtime.pivotResults[pivot.id];
      }
      if (FORMULA_SYNC_MUTATIONS.has(mutation.id)) {
        if (runtime.formula.getRecalculationMode() === 'manual' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
          synchronizeManualCellMutation(runtime.formula, runtime.model, mutation);
        } else {
          void scheduleFormulaRecalculation(runtime);
        }
      }
    }),
  );

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (runtime.disposed) return;
      if (source !== 'command') return;
      if (mutationCapability(mutation.id)?.durability === 'transient') return;
      runtime.pendingMutations.push({
        id: mutation.id,
        unitId: mutation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: [...mutation.affectedRanges],
      });
    }),
  );

  runtime.detachers.push(
    runtime.commands.onCommand((_commandId, _params, result) => {
      if (runtime.disposed) return;
      if (runtime.commands.activeDepth > 0) return;
      const batch = runtime.pendingMutations;
      runtime.pendingMutations = [];
      if (batch.length === 0) return;
      runtime.handlers.onMutationsApplied?.();
      const history = runtime.commands.getUndoEntries().at(-1);
      if (history) {
        runtime.collaboration?.recordLocalUndo({
          operationId: result.operationId,
          undoMutations: history.undo,
        });
      }
      if (runtime.collaboration) submitChangeset(runtime, result.operationId, batch);
      else runtime.pendingLocalOperations.push({ operationId: result.operationId, mutations: batch });
      void runtime.checkpointWorkspace();
    }),
  );

  runtime.detachers.push(
    runtime.commands.onHistoryReplay((source, entry) => {
      if (runtime.disposed) return;
      if (!runtime.collaboration || entry.undo.length === 0) return;
      const operation = source === 'undo'
        ? runtime.collaboration.enqueueCompensatingMutations(
          runtime.collaboration.undoOwnLast() ?? entry.undo,
          runtime.model.unitId,
        )
        : runtime.collaboration.enqueueLocalMutations(entry.redo, runtime.model.unitId);
      if (source === 'redo') {
        runtime.collaboration.recordLocalUndo({ operationId: entry.operationId, undoMutations: entry.undo });
      }
      scheduleOperation(runtime, operation);
      void runtime.checkpointWorkspace();
    }),
  );
}

function detachCoreListeners(runtime: SpreadsheetRuntime): void {
  for (const detach of runtime.detachers) detach();
  runtime.detachers = [];
  runtime.pendingMutations = [];
}

function replaceCollaborationSession(runtime: SpreadsheetRuntime, record: WorkspaceRecord | null): void {
  const existingPending = runtime.collaboration?.getPendingOperations() ?? [];
  const buffered = runtime.pendingLocalOperations.splice(0);
  const byId = new Map<string, import('@react-sheets/protocol').OperationEnvelope>();
  for (const operation of record?.pending.operations ?? []) byId.set(operation.operationId, operation);
  for (const operation of existingPending) byId.set(operation.operationId, operation);
  const pending = [...byId.values()].sort((left, right) => left.clientSequence - right.clientSequence);
  const nextClientSequence = Math.max(
    record?.pending.nextClientSequence ?? 0,
    ...pending.map((operation) => operation.clientSequence),
  );
  runtime.operationJournal.write(runtime.model.unitId, pending, nextClientSequence);
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    loadPending: () => {
      const journal = runtime.operationJournal.read(runtime.model.unitId);
      return journal ? { operations: journal.operations, nextClientSequence: journal.nextClientSequence } : null;
    },
    persistPending: (operations, sequence) => {
      runtime.operationJournal.write(runtime.model.unitId, operations, sequence);
      runtime.handlers.onWorkspacePersisted?.();
    },
  });
  runtime.collaboration.setRevision(runtime.remoteRevision);
  for (const entry of buffered) {
    runtime.collaboration.enqueueLocalMutations(entry.mutations, runtime.model.unitId, entry.operationId);
  }
}

function submitChangeset(
  runtime: SpreadsheetRuntime,
  operationId: string,
  mutations: MutationInfo[],
): void {
  if (!runtime.collaboration) {
    runtime.handlers.onSaveState?.(runtime.remoteSyncRequested ? 'offline' : 'saved');
    return;
  }
  const operation = runtime.collaboration.enqueueLocalMutations(mutations, runtime.model.unitId, operationId);
  scheduleOperation(runtime, operation);
}

function scheduleOperation(
  runtime: SpreadsheetRuntime,
  operation: import('@react-sheets/protocol').OperationEnvelope,
): void {
  if (!runtime.collaboration) return;
  runtime.ownOperationIds.add(operation.operationId);
  runtime.handlers.onSaveState?.('saving');
  // The operation is durable immediately. Only an open authenticated socket
  // may start a flush; disconnected edits remain in the journal.
  if (!runtime.localOnly && runtime.collab && runtime.collaboration.offlineQueue.getState() !== 'offline') {
    void runtime.collaboration.offlineQueue.flushAll().then(({ failed }) => {
      if (failed > 0) runtime.handlers.onNotice?.('Some offline changes could not be synced');
    });
  } else {
    runtime.handlers.onSaveState?.(runtime.remoteSyncRequested ? 'offline' : 'saved');
  }
}

export function rehydrateFormulaAfterRestore(runtime: SpreadsheetRuntime, revision?: number): void {
  runtime.formula.disposeCalculationTasks();
  runtime.formula = rebuildFormulaEngine(runtime.model);
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
  if (revision != null) {
    runtime.remoteRevision = revision;
    runtime.collaboration?.setRevision(revision);
  }
  runtime.pivotResults = {};
  void scheduleFormulaRecalculation(runtime);
}

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.primarySheetId });
  loadFormulaInputs(engine, workbook);
  return engine;
}

export function hydrateRuntime(runtime: SpreadsheetRuntime, response: SnapshotResponse): void {
  if (runtime.disposed) return;
  const workbook = WorkbookModel.fromSnapshot(response.snapshot);
  // Legacy block overlays are normalized exactly once at the snapshot boundary.
  // All runtime reads after this point require the canonical CellPatch carrier.
  for (const sheet of workbook.getSheets()) migrateDataRegionCellPatches(sheet);
  detachCoreListeners(runtime);
  runtime.formula.disposeCalculationTasks();
  runtime.model = workbook;
  runtime.commands = new CommandRuntime(workbook);
  registerSpreadsheetFeatures(runtime.commands, runtime.drawing);
  runtime.formula = rebuildFormulaEngine(workbook);
  runtime.formulaAudit.setFormula(runtime.formula);
  registerFormulaAuditCommands(runtime.commands.registry, runtime.formulaAudit);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  runtime.collaboration?.setRevision(response.revision);
  runtime.collaboration?.rebindCommands(runtime.commands);
  runtime.pivotResults = {};
  initializeDataContent(runtime);
  void scheduleFormulaRecalculation(runtime);
}

function initializeDataContent(runtime: SpreadsheetRuntime): void {
  for (const detach of runtime.dataContentDetachers) detach();
  runtime.dataContentDetachers = [];
  runtime.dataContent.clear();
  for (const manifest of runtime.model.dataModel.sources.values()) {
    const query = new DataSourceContentQuery(manifest, {
      get: async (reference) => {
        const ref = manifest.blocks.find((block) => block.id === reference.id && block.dataSourceId === reference.dataSourceId && block.checksum === reference.checksum);
        if (!ref) return null;
        const bytes = await runtime.dataBlocks.get(ref);
        return { sourceId: ref.dataSourceId, blockId: ref.id, checksum: ref.checksum, bytes };
      },
    });
    runtime.dataContentDetachers.push(query.subscribe(() => {
      if (!runtime.disposed) runtime.handlers.onMutationsApplied?.();
    }));
    runtime.dataContent.set(manifest.id, query);
  }
}

/** Replay durable local intent on top of the authoritative server snapshot. */
export function replayPendingOperations(
  runtime: SpreadsheetRuntime,
  operations = runtime.collaboration?.getPendingOperations() ?? [],
): number {
  const pending = operations;
  let applied = 0;
  for (const operation of pending) {
    const items = operation.mutations.map((mutation) => {
      const metadata = runtime.commands.registry.getMutationMetadata(mutation.id);
      let affectedRanges: MutationInfo['affectedRanges'] = [];
      try {
        const resolved = metadata?.affectedRanges?.resolve(mutation.params as never);
        if (Array.isArray(resolved)) affectedRanges = [...resolved];
      } catch {
        affectedRanges = [];
      }
      return {
        id: mutation.id,
        unitId: operation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges,
      } satisfies MutationInfo;
    });
    runtime.commands.applyRemoteMutations(items);
    applied += 1;
  }
  return applied;
}

async function loadHistoryAndReplayPending(runtime: SpreadsheetRuntime): Promise<void> {
  try {
    const revisions = await runtime.api.listRevisions(runtime.model.unitId);
    runtime.collaboration?.loadCommittedHistory(revisions.map((record) => record.payload));
    runtime.handlers.onRemoteRevisions?.(revisions);
  } catch {
    runtime.handlers.onRemoteRevisions?.([]);
  }
  replayPendingOperations(runtime);
}

export function startCollaborationSession(
  runtime: SpreadsheetRuntime,
  getSelectionKey: () => string,
  authTokenProvider: AuthTokenProvider | undefined = runtime.authTokenProvider,
  shareTokenProvider: ShareTokenProvider | undefined = runtime.shareTokenProvider,
): () => void {
  runtime.disposed = false;
  if (typeof window === 'undefined') return () => undefined;

  let active = true;
  let disposeOpenSession: (() => void) | null = null;
  void runtime.persistenceReady.then(() => {
    if (!active || runtime.disposed || runtime.localOnly) {
      runtime.handlers.onCollabStatus?.('closed');
      return;
    }
    runtime.collaboration ??= new CollaborationSession(runtime.commands);
    runtime.collaboration.attachTransport(async (operation) => {
      if (runtime.disposed) throw new Error('Workbook runtime has been disposed');
      runtime.ownOperationIds.add(operation.operationId);
      try {
        const committed = await runtime.api.commitOperation(runtime.model.unitId, operation);
        const revision = committed.operation.revision;
        runtime.remoteRevision = Math.max(runtime.remoteRevision, revision);
        runtime.collaboration?.acknowledge(operation.operationId, revision);
        await runtime.checkpointWorkspace(false);
        runtime.handlers.onSaveState?.('saved');
        return revision;
      } catch (error) {
        runtime.ownOperationIds.delete(operation.operationId);
        runtime.collaboration?.reject(operation.operationId, error instanceof Error ? error : new Error(String(error)));
        runtime.handlers.onSaveState?.('conflict');
        runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Change could not be committed');
        throw error;
      }
    });
    runtime.collaboration.setRevision(runtime.remoteRevision);

    const protocol = window.location.protocol === 'https:' ? 'wss' : 'ws';
    const client = new CollabSocketClient(protocol + '://' + window.location.host + '/ws', {
      authTokenProvider,
      shareTokenProvider,
    });
    runtime.collab = client;

    const applyRemote = (message: OperationMessage) => {
      if (runtime.disposed) return;
      if (message.type === 'revision.created') {
        if (message.payload.unitId !== runtime.model.unitId) return;
        if (runtime.ownOperationIds.has(message.payload.operationId)) return;
        runtime.collaboration?.applyRemote(message.payload);
        runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
        runtime.collaboration?.setRevision(runtime.remoteRevision);
        void checkpointWorkspace(runtime, false);
        runtime.handlers.onMutationsApplied?.();
        void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
      } else if (message.type === 'cursor.broadcast' || message.type === 'presence.broadcast') {
        if (!message.unitId || message.unitId !== runtime.model.unitId) return;
        if (message.type === 'presence.broadcast' && (message.state as { status?: string } | null)?.status === 'offline') {
          runtime.handlers.onPeersChange?.([]);
          runtime.collaboration?.presence.removeUser(message.actorId);
          return;
        }
        const cursorState = message.state as { row?: number; column?: number; name?: string; sheetId?: string } | null;
        const peer = mapPeerCursor(message.actorId, cursorState, runtime.model.primarySheetId);
        runtime.collaboration?.presence.upsertUser({
          actorId: peer.actorId,
          displayName: peer.name,
          color: peer.color,
        });
        if (runtime.collaboration) updatePresenceFromPeer(runtime.collaboration, peer);
        runtime.handlers.onPeersChange?.([peer]);
      }
    };

    const detachMessage = client.onMessage(applyRemote);
    const detachStatus = client.onStatus((status: 'connecting' | 'open' | 'closed') => {
      if (runtime.disposed) return;
      runtime.handlers.onCollabStatus?.(status);
      runtime.remoteConnected = status !== 'closed';
      runtime.collaboration?.offlineQueue.setOnline(status === 'open');
      if (status === 'closed') runtime.collaboration?.transportClosed();
      if (status === 'closed') runtime.handlers.onSaveState?.(runtime.remoteSyncRequested ? 'offline' : 'saved');
      else if (status === 'connecting') runtime.handlers.onSaveState?.('syncing');
      if (status === 'open') {
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
      client.send({ type: 'cursor.updated', unitId: runtime.model.unitId, state });
    }, 400);

    disposeOpenSession = () => {
      window.clearInterval(broadcastTimer);
      detachMessage();
      detachStatus();
      client.close();
      runtime.collaboration?.attachTransport(undefined);
      runtime.collab = null;
    };
  }).catch(() => {
    runtime.handlers.onCollabStatus?.('closed');
  });

  const dispose = () => {
    active = false;
    disposeOpenSession?.();
    if (runtime.collabDispose === dispose) runtime.collabDispose = null;
  };
  runtime.collabDispose = dispose;
  return dispose;
}

export function startPersistenceSession(runtime: SpreadsheetRuntime): () => void {
  runtime.disposed = false;
  // `disposeSpreadsheetRuntime` detaches command listeners. Reattach them on
  // a real remount so StrictMode does not leave edits outside the journal.
  attachCoreListeners(runtime);
  let active = true;
  const initialization = initializePersistence(runtime, () => active);
  runtime.persistenceReady = initialization;
  const dispose = () => {
    active = false;
    if (runtime.bootstrapDispose === dispose) runtime.bootstrapDispose = null;
  };
  runtime.bootstrapDispose = dispose;
  return dispose;
}

export function disposeSpreadsheetRuntime(runtime: SpreadsheetRuntime): void {
  if (runtime.disposed) return;
  runtime.disposed = true;
  runtime.collabDispose?.();
  runtime.bootstrapDispose?.();
  detachCoreListeners(runtime);
  runtime.formula.disposeCalculationTasks();
  for (const detach of runtime.dataContentDetachers) detach();
  runtime.dataContentDetachers = [];
  runtime.dataContent.clear();
  runtime.collaboration?.attachTransport(undefined);
  runtime.collaboration = null;
  runtime.collab = null;
}

async function initializePersistence(runtime: SpreadsheetRuntime, isActive: () => boolean): Promise<void> {
  const localPendingBeforeLoad = runtime.collaboration?.getPendingOperations() ?? [];
  if (!runtime.localOnly && !(await hasValidRemoteBinding(runtime))) runtime.localOnly = true;
  let localRecord: WorkspaceRecord | null = null;
  try {
    localRecord = await runtime.workspacePersistence.load(runtime.model.unitId);
    const canDiscoverLocalDefault = runtime.model.unitId === 'wb-local-default'
      && (typeof window === 'undefined' || !/^\/workbooks\/[^/]+\/?$/.test(window.location.pathname));
    if (!localRecord && canDiscoverLocalDefault) {
      const summaries = await runtime.workspacePersistence.list();
      const first = summaries[0];
      if (first) localRecord = await runtime.workspacePersistence.load(first.unitId);
    }
  } catch {
    runtime.handlers.onNotice?.('Local IndexedDB workspace is unavailable');
  }

  if (!isActive()) return;

  if (localRecord) {
    runtime.workspaceRecord = localRecord;
    runtime.localRevision = localRecord.localRevision;
    runtime.remoteRevision = localRecord.serverRevision;
    runtime.localOnly = runtime.localOnly || localRecord.syncMode === 'local-only';
    runtime.remoteSyncRequested = runtime.remoteSyncRequested || localRecord.syncMode === 'remote';
    if (!isActive()) return;
    hydrateRuntime(runtime, {
      snapshot: localRecord.snapshot,
      revision: localRecord.serverRevision,
    });
    replaceCollaborationSession(runtime, localRecord);
    if (localPendingBeforeLoad.length > 0) replayPendingOperations(runtime, localPendingBeforeLoad);
    runtime.handlers.onNotice?.('Workbook restored from local IndexedDB');
  }

  if (runtime.localOnly) {
    runtime.remoteConnected = false;
    runtime.handlers.onAccessRole?.(null);
    // A StrictMode dispose may have detached the collaboration journal even
    // though the workbook remains local-only. Recreate the journal owner
    // before the first post-remount command so edits remain durable.
    if (!runtime.collaboration) replaceCollaborationSession(runtime, localRecord);
    const checkpointed = await checkpointStartupLocally(runtime);
    if (isActive()) {
      runtime.handlers.onSaveState?.(checkpointed ? (runtime.remoteSyncRequested ? 'offline' : 'saved') : 'error');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
      runtime.handlers.onMutationsApplied?.();
    }
    return;
  }

  try {
    const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
    const access = await runtime.api.getAccess(runtime.model.unitId);
    if (!isActive()) return;
    hydrateRuntime(runtime, snapshotResponse);
    runtime.remoteRevision = snapshotResponse.revision;
    runtime.localOnly = false;
    runtime.remoteSyncRequested = true;
    replaceCollaborationSession(runtime, localRecord);
    await loadHistoryAndReplayPending(runtime);
    if (!isActive()) return;
    runtime.workspaceRecord = await runtime.workspacePersistence.checkpoint(runtime.model.snapshot(), runtime.localRevision, runtime.remoteRevision, 'remote');
    runtime.remoteConnected = true;
    runtime.handlers.onAccessRole?.(access.role);
    if (isActive()) {
      runtime.handlers.onSaveState?.('saved');
      runtime.handlers.onNotice?.('Workbook restored from server');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
      runtime.handlers.onMutationsApplied?.();
      runtime.handlers.onWorkspacePersisted?.();
    }
  } catch (error) {
    // Authentication, authorization and an unknown shared workbook are
    // authoritative server decisions. They must never be disguised as a new
    // local workbook with the same URL. Only an actual unavailable service
    // leaves the user in offline local mode.
    if (isAuthoritativeRemoteFailure(error)) {
      runtime.remoteConnected = false;
      runtime.handlers.onAccessRole?.(null);
      if (isActive()) {
        runtime.handlers.onSaveState?.('error');
        runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Server access was rejected');
        runtime.handlers.onPhaseChange?.('error');
      }
      return;
    }

    runtime.localOnly = true;
    runtime.remoteConnected = false;
    runtime.remoteSyncRequested = true;
    runtime.handlers.onAccessRole?.(null);
    const checkpointed = await checkpointStartupLocally(runtime);
    if (isActive()) {
      runtime.handlers.onSaveState?.(checkpointed ? 'offline' : 'error');
      runtime.handlers.onNotice?.('Server unavailable; using local IndexedDB workspace');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onMutationsApplied?.();
    }
  }
}

/**
 * Startup may continue with an in-memory local workbook when IndexedDB is
 * unavailable. A server failure must never become a blank or permanently
 * loading frontend merely because local persistence also reports an error.
 */
async function checkpointStartupLocally(runtime: SpreadsheetRuntime): Promise<boolean> {
  try {
    runtime.workspaceRecord = await runtime.workspacePersistence.checkpoint(
      runtime.model.snapshot(),
      runtime.localRevision,
      runtime.remoteRevision,
      'local-only',
    );
    return true;
  } catch (error) {
    runtime.workspaceRecord = null;
    runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Local workspace persistence is unavailable');
    return false;
  }
}

function isAuthoritativeRemoteFailure(error: unknown): boolean {
  return error instanceof ApiRequestError
    && (error.status === 401 || error.status === 403 || error.status === 404);
}

async function hasValidRemoteBinding(runtime: SpreadsheetRuntime): Promise<boolean> {
  if (!runtime.authTokenProvider && !runtime.shareTokenProvider) return false;
  try {
    const token = await runtime.authTokenProvider?.();
    if (token?.trim()) return true;
    const shareToken = await runtime.shareTokenProvider?.();
    return Boolean(shareToken?.trim());
  } catch {
    return false;
  }
}

export type { HistoryEntry };
