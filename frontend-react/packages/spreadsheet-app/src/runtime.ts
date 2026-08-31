import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry, type MutationInfo } from '@react-sheets/command-runtime';
import { canonicalExcelDateFromUtcDate, FormulaEngine, type CalculationSessionPort, type CanonicalExcelDateParts, type CellAddressInput, type ExcelDateSystem } from '@react-sheets/formula-engine';
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
import { FindIndex } from '@react-sheets/sheet-features';
import { activateSpreadsheetFeatures, createSpreadsheetFeatureRuntime, registerSpreadsheetFeatures, type SpreadsheetFeatureRuntime } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { createDefaultConnectorRegistry, type ConnectorRegistry } from './features/query';
import { FormulaAuditController, registerFormulaAuditCommands } from './features/formula-audit';
import { DataSourceContentQuery } from './features/data-source';
import { CollaborationSession } from './collaboration/collaboration-session';
import { createWorkbookRowVisibilityResolver, type WorkbookRowVisibilityResolver } from './formula-visibility';
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
  LocalAssetStore,
  type AssetStore,
  type WorkspacePersistenceOptions,
  type WorkspaceRecord,
  WorkspaceStorageError,
} from './features/persistence';
import { migrateLegacyImageAssets } from './features/persistence/asset-migration';
import { isAssetRef, type AssetRef } from '@react-sheets/core-model';
import type { WorkbookResolution } from './features/workbook-catalog';
import type { NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';

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
  onDataSourceContentChanged?: (sourceId: string) => void;
  onRuntimeFailure?: (failure: RuntimeFailure) => void;
}

export type RuntimeFailureCode = 'HISTORY_LOAD_FAILED' | 'HISTORY_GAP' | 'REMOTE_WORKBOOK_UNAVAILABLE' | 'FEATURE_LIFECYCLE_FAILED';

export interface RuntimeFailure {
  readonly code: RuntimeFailureCode;
  readonly message: string;
  readonly recovery: string;
  readonly cause?: unknown;
}

export interface SpreadsheetRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  rowVisibilityResolver: WorkbookRowVisibilityResolver;
  formulaAudit: FormulaAuditController;
  dateSystem: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Node/SSR callers must supply an explicit persistent calculation session. */
  calculationSessionPort?: CalculationSessionPort;
  collaborationUrl?: string;
  model: WorkbookModel;
  commands: CommandRuntime;
  drawing: DrawingRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
  pendingMutations: MutationInfo[];
  /** Local-durable geometry changed without producing a remote operation. */
  pendingLocalCheckpoint: boolean;
  /** Mutation facts are drained by the refresh coordinator after each apply. */
  pendingPivotMutations: MutationInfo[];
  drainPivotMutations: () => MutationInfo[];
  detachers: Array<() => void>;
  handlers: RuntimeHandlers;
  ownOperationIds: Set<string>;
  nextClientSequence: number;
  pivotResults: Record<string, import('@react-sheets/core-model').PivotResultTree>;
  pivotErrors: Record<string, import('./features/pivot/task-protocol').PivotTaskError>;
  collab: CollabSocketClient | null;
  collabDispose: (() => void) | null;
  broadcastPresence: (state: unknown) => boolean;
  collaboration: CollaborationSession | null;
  bootstrapDispose: (() => void) | null;
  operationJournal: OperationJournalStore;
  workspacePersistence: WorkspacePersistence;
  dataBlocks: DataBlockSynchronizer;
  assetStore: AssetStore;
  dataContent: Map<string, DataSourceContentQuery>;
  dataContentDetachers: Array<() => void>;
  workspaceRecord: WorkspaceRecord | null;
  localRevision: number;
  localOnly: boolean;
  remoteSyncRequested: boolean;
  formulaCalculation: Promise<void>;
  persistenceReady: Promise<void>;
  pendingLocalOperations: Array<{ operationId: string; mutations: MutationInfo[] }>;
  checkpointWorkspace: (advanceLocalRevision?: boolean, artifact?: NativeDocumentArtifact) => Promise<void>;
  connectors: ConnectorRegistry;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  resolution?: WorkbookResolution;
  /** Feature lifecycle is workbook-instance owned; it is never a process singleton. */
  featureRuntime: SpreadsheetFeatureRuntime;
  /** Sparse content index shared by Find/Replace and selection commands. */
  findIndex: FindIndex;
  /** Runtime lifecycle is explicit so late Worker callbacks cannot
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
  persistence?: WorkspacePersistenceOptions;
  workspacePersistence?: WorkspacePersistence;
  assetStore?: AssetStore;
  resolution?: WorkbookResolution;
  dateSystem?: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Node/SSR callers must supply an explicit persistent calculation session. */
  calculationSessionPort?: CalculationSessionPort;
  collaborationUrl?: string;
} = {}): SpreadsheetRuntime {
  const unitId = options.unitId ?? options.resolution?.unitId ?? resolveUnitId();
  if (options.resolution && options.resolution.unitId !== unitId) throw new Error('Workbook resolution unitId does not match runtime unitId');
  const model = new WorkbookModel(unitId, 'Untitled workbook');
  const dateSystem = options.dateSystem ?? '1900';
  const canonicalReferenceDate = options.canonicalReferenceDate
    ? structuredClone(options.canonicalReferenceDate)
    : canonicalExcelDateFromUtcDate(new Date(), dateSystem);
  const commands = new CommandRuntime(model);
  const drawing = new DrawingRuntime();
  const featureRuntime = createSpreadsheetFeatureRuntime();
  const connectors = createDefaultConnectorRegistry();
  let formula: FormulaEngine | undefined;
  const rowVisibilityResolver = createWorkbookRowVisibilityResolver(model, dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return formula?.getCellResult(address)?.value;
  });
  formula = new FormulaEngine({ defaultSheetId: 'sheet-1', dateSystem, canonicalReferenceDate, collationContext: model.collationContext, calculationSettings: model.calculationSettings, rowVisibilityResolver, calculationSessionPort: options.calculationSessionPort });
  const findIndex = new FindIndex(model, (sheet, row, column) => formula?.getCellValue({ sheetId: sheet.id, row, column }));
  const formulaAudit = new FormulaAuditController(formula);
  registerSpreadsheetFeatures(commands, drawing, featureRuntime);
  activateSpreadsheetFeatures(featureRuntime, { documentType: 'spreadsheet', environment: typeof window === 'undefined' ? 'worker' : 'browser' });
  featureRuntime.advance('ready');
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
  const assetStore = options.assetStore ?? new LocalAssetStore(model.unitId, workspacePersistence.coordinator);
  runtime = {
    api,
    formula: formula as FormulaEngine,
    rowVisibilityResolver,
    formulaAudit,
    dateSystem,
    canonicalReferenceDate,
    collaborationUrl: options.collaborationUrl,
    model,
    commands,
    drawing,
    remoteConnected: false,
    remoteRevision: 0,
    pendingMutations: [],
    pendingLocalCheckpoint: false,
    pendingPivotMutations: [],
    drainPivotMutations: () => {
      const pending = runtime.pendingPivotMutations;
      runtime.pendingPivotMutations = [];
      return pending;
    },
    detachers: [],
    handlers: {},
    ownOperationIds: new Set(),
    nextClientSequence: 0,
    pivotResults: {},
    pivotErrors: {},
    collab: null,
    collabDispose: null,
    broadcastPresence: () => false,
    collaboration: null,
    bootstrapDispose: null,
    operationJournal,
    workspacePersistence,
    dataBlocks,
    assetStore,
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
    resolution: options.resolution,
    featureRuntime,
    calculationSessionPort: options.calculationSessionPort,
    findIndex,
    disposed: false,
  };
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  // The offline journal records operation intent and its client sequence.
  // The same memory transaction also checkpoints the canonical local
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
      void enqueuePersistenceWrite(runtime, async () => {
        const current = await runtime.workspacePersistence.load(runtime.model.unitId);
        const storageRevision = await runtime.workspacePersistence.commitOperationJournal(
          runtime.model.unitId,
          operations,
          nextClientSequence,
          current?.storageRevision,
        );
        if (runtime.workspaceRecord) runtime.workspaceRecord.storageRevision = storageRevision;
        runtime.handlers.onWorkspacePersisted?.();
      }).catch((error: unknown) => publishPersistenceFailure(runtime, error));
    },
  });
  runtime.checkpointWorkspace = (advanceLocalRevision = true, artifact) => checkpointWorkspace(runtime, advanceLocalRevision, artifact);
  installCommandCellValueResolver(runtime);
  attachCoreListeners(runtime);
  return runtime;
}

/** Keep command-side sorting on the same formula/spill value authority as the canvas. */
function installCommandCellValueResolver(runtime: SpreadsheetRuntime): void {
  runtime.commands.setCellValueResolver((sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    const spillValue = runtime.formula.getSpillValueAt(sheet.id, row, column);
    if (spillValue !== undefined) return spillValue;
    const cell = sheet.cells.get(row, column);
    if (cell?.formula !== undefined) {
      const result = runtime.formula.getCellResult(address);
      if (result === undefined) throw new Error(`Sort formula result unavailable at ${sheet.id}!${row}:${column}`);
      return result.value;
    }
    if (runtime.formula.getCellResult(address) !== undefined) {
      return runtime.formula.getCellValue(address);
    }
    return cell?.formulaValue ?? cell?.value ?? null;
  });
}

const FORMULA_SYNC_MUTATIONS = new Set([
  'cell.set',
  'cell.restore',
  'range.set',
  'fill.applied',
  'fill.restored',
  'flashFill.applied',
  'flashFill.restored',
  'range.clear',
  'range.paste',
  'style.preset.set',
  'dataRegion.materialize.commit',
  'dataRegion.materialize.restore',
  'query.load.range',
  'query.load.sheet-table',
  'query.load.pivot-source',
  'query.load.workbook-table',
  'cells.inserted',
  'cells.deleted',
  'cells.inserted.restore',
  'rows.permuted',
  'cells.deleted.restore',
  'rows.inserted',
  'rows.deleted',
  'columns.inserted',
  'columns.deleted',
  'sheet.rename',
  'sheet.remove',
  'sheet.restore',
  'sheet.duplicated',
  'sheetTable.add',
  'sheetTable.remove',
  'sheetTable.update',
  'table.add',
  'table.remove',
  'name.set',
  'name.remove',
  'workbook.calculation.mode.set',
  'row.hidden',
  'row.unhidden',
  'rows.unhidden.all',
  'rows.hidden.restore',
  'sheet.rows.visibility.set',
  'sheet.rows.unhide.all',
  'autoFilter.set',
  'autoFilter.remove',
  'sheet.autoFilter.set',
  'sheet.autoFilter.remove',
  'outline.group.toggle',
  'outline.showLevel',
]);

const VISIBILITY_MUTATIONS = new Set([
  'row.hidden', 'row.unhidden', 'rows.unhidden.all', 'rows.hidden.restore',
  'sheet.rows.visibility.set', 'sheet.rows.unhide.all',
  'autoFilter.set', 'autoFilter.remove', 'sheet.autoFilter.set', 'sheet.autoFilter.remove',
  'outline.group.toggle', 'outline.showLevel',
]);

const DIRECT_CELL_WRITE_MUTATIONS = new Set([
  'cell.set',
  'cell.restore',
  'range.set',
  'fill.applied',
  'fill.restored',
  'flashFill.applied',
  'flashFill.restored',
  'range.clear',
  'range.paste',
  'cells.inserted',
  'cells.deleted',
  'cells.inserted.restore',
  'cells.deleted.restore',
]);

const FORMULA_METADATA_MUTATIONS = new Set([
  'sheetTable.add', 'sheetTable.remove', 'sheetTable.update',
  'table.add', 'table.remove',
  'name.set', 'name.remove',
  'cells.inserted', 'cells.deleted', 'cells.inserted.restore', 'cells.deleted.restore',
  'rows.inserted', 'rows.deleted', 'columns.inserted', 'columns.deleted',
  'sheet.remove', 'sheet.restore', 'sheet.duplicated', 'sheet.rename',
]);

const FIND_INDEX_MUTATIONS = new Set([
  ...DIRECT_CELL_WRITE_MUTATIONS,
  'sheet.add', 'sheet.remove', 'sheet.restore', 'sheet.duplicated',
  'find.replaced',
  'note.set', 'note.remove', 'note.visibility',
  'comment.add', 'comment.update', 'comment.remove', 'comment.reply', 'comment.reply.remove', 'comment.resolve',
]);

function synchronizeManualCellMutation(engine: FormulaEngine, workbook: WorkbookModel, mutation: MutationInfo): boolean {
  let formulaInputChanged = false;
  for (const range of mutation.affectedRanges) {
    const sheet = workbook.getSheet(range.sheetId);
    configureFormulaSpillEnvironment(engine, sheet);
    const expected = new Set<string>();
    const previousFormulaKeys = new Set(
      engine.listFormulaCells()
        .filter((address) => address.sheetId === sheet.id
          && address.row >= range.startRow && address.row <= range.endRow
          && address.column >= range.startColumn && address.column <= range.endColumn)
        .map((address) => `${address.row}:${address.column}`),
    );
    sheet.cells.forEach((cell, row, column) => {
      if (row < range.startRow || row > range.endRow || column < range.startColumn || column > range.endColumn) return;
      const key = `${row}:${column}`;
      expected.add(key);
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) {
        formulaInputChanged = true;
        engine.applyInputDelta({ kind: 'set-formula', address, formula: cell.formula });
      }
      else if (cell.formula === undefined && cell.value != null) engine.applyInputDelta({ kind: 'set-value', address, value: cell.value as never });
      else {
        formulaInputChanged ||= previousFormulaKeys.has(key);
        engine.applyInputDelta({ kind: 'clear', address });
      }
    });
    for (const address of engine.listInputCells()) {
      if (address.sheetId !== sheet.id || address.row < range.startRow || address.row > range.endRow || address.column < range.startColumn || address.column > range.endColumn) continue;
      if (!expected.has(`${address.row}:${address.column}`)) engine.applyInputDelta({ kind: 'clear', address });
    }
  }
  return formulaInputChanged;
}

/**
 * Load only formula inputs from the canonical workbook. The actual evaluation
 * is intentionally scheduled separately through FormulaEngine.recalculateAsync
 * so browser calculation stays in its Worker.
 */
const loadedFormulaCounts = new WeakMap<FormulaEngine, number>();

function loadFormulaInputs(engine: FormulaEngine, workbook: WorkbookModel): number {
  const initialized = loadedFormulaCounts.get(engine);
  if (initialized !== undefined) return engine.listFormulaCells().length;
  const mode = engine.getRecalculationMode();
  engine.cancelCalculation();
  engine.reset();
  engine.setRecalculationMode('manual');
  engine.setDefinedNameModels(workbook.definedNameModels);
  configureWorkbookSpillEnvironments(engine, workbook);
  syncWorkbookSheetTables(engine, workbook);
  let formulaCount = 0;
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell) => {
      if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) formulaCount += 1;
    });
  }
  // Seed authored values before formulas so a formula's initial host-side
  // result never observes an absent precedent. The Worker still owns the
  // canonical recalculation; this ordering only establishes its input graph.
  if (formulaCount > 0) {
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        if (cell.formula !== undefined || cell.value == null) return;
        engine.setValue({ sheetId: sheet.id, row, column }, cell.value as never);
      });
    }
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        if (cell.formula === undefined || cell.formulaMetadata?.preservedOnly) return;
        engine.setFormula({ sheetId: sheet.id, row, column }, cell.formula);
      });
    }
  }
  // A value-only workbook has no formula dependency graph. Keeping tens of
  // thousands of ordinary cells in FormulaEngine duplicates CellMatrix and
  // makes native-document open proportional to every imported value for no calculation
  // benefit. Formula workbooks retain the complete existing input contract.
  engine.setRecalculationMode(mode);
  loadedFormulaCounts.set(engine, formulaCount);
  return formulaCount;
}

interface FormulaQueueState {
  tail: Promise<void>;
  scheduled: boolean;
  epoch: number;
  force: boolean;
  roots?: readonly CellAddressInput[];
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
export function scheduleFormulaRecalculation(runtime: SpreadsheetRuntime, force = false, roots?: readonly CellAddressInput[]): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  const state = formulaQueueStates.get(runtime) ?? {
    tail: Promise.resolve(),
    scheduled: false,
    epoch: 0,
    force: false,
    roots: undefined,
  } satisfies FormulaQueueState;
  formulaQueueStates.set(runtime, state);
  state.epoch += 1;
  state.force ||= force;
  if (roots !== undefined) state.roots = [...roots];
  else state.roots = undefined;
  runtime.formula.cancelCalculation();
  if (state.scheduled) return runtime.formulaCalculation;

  state.scheduled = true;
  state.tail = state.tail
    .then(async () => {
      if (runtime.disposed) return;
      state.scheduled = false;
      const epoch = state.epoch;
      const forceCalculation = state.force;
      const calculationRoots = state.roots;
      state.force = false;
      state.roots = undefined;
      const engine = runtime.formula;
      const workbook = runtime.model;
      const formulaCount = loadFormulaInputs(engine, workbook);
      if (formulaCount === 0) {
        runtime.handlers.onSaveState?.(localFormulaIdleState(runtime));
        return;
      }
      if (engine.getRecalculationMode() !== 'automatic' && !forceCalculation) {
        if (!runtime.disposed && epoch === state.epoch && runtime.formula === engine && runtime.model === workbook) {
          runtime.handlers.onMutationsApplied?.();
        }
        return;
      }

      runtime.handlers.onSaveState?.('calculating');
      try {
        const report = await engine.recalculateAsync(calculationRoots);
        if (runtime.disposed || epoch !== state.epoch || runtime.formula !== engine || runtime.model !== workbook) return;
        syncWorkbookSpills(engine, workbook);
        for (const sheetId of new Set(report.recalculated.map((address) => address.sheetId))) runtime.findIndex.rebuildSheet(sheetId);
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
const persistenceWriteChains = new WeakMap<SpreadsheetRuntime, Promise<void>>();
const persistenceWriteFailures = new WeakMap<SpreadsheetRuntime, unknown>();

function enqueuePersistenceWrite<T>(runtime: SpreadsheetRuntime, operation: () => Promise<T>): Promise<T> {
  const previous = persistenceWriteChains.get(runtime) ?? Promise.resolve();
  const next = previous.then(async () => {
    // Commands can arrive before WorkbookSession.start() finishes its first
    // persistence pass (unit tests and headless callers also intentionally do
    // not start the UI lifecycle). Establish the local workspace head before
    // journaling the first mutation, otherwise commitOperationJournal sees an
    // unknown workbook and permanently trips the write latch.
    await runtime.persistenceReady;
    if (runtime.localOnly && !runtime.workspaceRecord) {
      await checkpointStartupLocally(runtime);
    }
    const failure = persistenceWriteFailures.get(runtime);
    if (failure !== undefined) throw new WorkspaceStorageError({
      code: 'STORAGE_TRANSACTION_FAILED',
      operation: 'checkpoint',
      message: 'Workbook persistence is blocked after a previous transaction failure.',
      recovery: 'Reopen the workbook or invoke the explicit storage recovery flow before saving again.',
      cause: failure,
    });
    try {
      return await operation();
    } catch (error) {
      persistenceWriteFailures.set(runtime, error);
      throw error;
    }
  });
  persistenceWriteChains.set(runtime, next.then(() => undefined, () => undefined));
  return next;
}

function checkpointWorkspace(runtime: SpreadsheetRuntime, advanceLocalRevision = true, artifact?: NativeDocumentArtifact): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  if (advanceLocalRevision) runtime.localRevision += 1;
  const snapshot = runtime.model.snapshot();
  const localRevision = runtime.localRevision;
  const serverRevision = runtime.remoteRevision;
  const resolution = runtime.resolution;
  const syncMode = resolution?.binding.syncMode ?? (runtime.localOnly ? 'local-only' as const : 'remote' as const);
  const metadata = resolution ? {
    location: resolution.binding.location,
    lifecycle: resolution.lifecycle,
    source: runtime.workspaceRecord?.metadata.source ?? 'native' as const,
    role: resolution.access?.role ?? runtime.workspaceRecord?.metadata.role ?? 'viewer' as const,
  } : undefined;
  const pendingJournal = runtime.operationJournal.read(runtime.model.unitId);
  const previous = checkpointChains.get(runtime) ?? Promise.resolve();
  const next = previous.then(() => enqueuePersistenceWrite(runtime, async () => {
      if (runtime.disposed) return;
      const record = artifact
        ? await runtime.workspacePersistence.checkpointWithArtifact(snapshot, localRevision, serverRevision, syncMode, artifact, pendingJournal, metadata)
        : await runtime.workspacePersistence.checkpoint(snapshot, localRevision, serverRevision, syncMode, pendingJournal, metadata);
      if (runtime.disposed) return;
      runtime.workspaceRecord = record;
      await runtime.assetStore.reconcile(collectAssetReferences(snapshot, [
        ...(pendingJournal?.operations ?? []),
        ...runtime.commands.getUndoEntries(),
        ...runtime.commands.getRedoEntries(),
      ]));
      if (runtime.disposed) return;
      runtime.handlers.onWorkspacePersisted?.();
    }));
  checkpointChains.set(runtime, next.then(() => undefined, () => undefined));
  return next;
}

function collectAssetReferences(snapshot: unknown, pending: readonly unknown[]): AssetRef[] {
  const references: AssetRef[] = [];
  const visit = (value: unknown): void => {
    if (isAssetRef(value)) {
      references.push(value);
      return;
    }
    if (Array.isArray(value)) {
      for (const entry of value) visit(entry);
      return;
    }
    if (value && typeof value === 'object') for (const entry of Object.values(value)) visit(entry);
  };
  visit(snapshot);
  visit(pending);
  return references;
}

export function attachCoreListeners(runtime: SpreadsheetRuntime): void {
  detachCoreListeners(runtime);

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (runtime.disposed) return;
      if (VISIBILITY_MUTATIONS.has(mutation.id)) {
        runtime.rowVisibilityResolver.invalidate();
        runtime.formula.notifyVisibilityChanged();
      }
      if (mutation.id === 'workbook.calculation.mode.set') {
        const mode = (mutation.params as { mode?: unknown } | undefined)?.mode;
        if (mode !== 'automatic' && mode !== 'manual' && mode !== 'partial') throw new Error('Workbook calculation mode mutation is invalid');
        runtime.formula.setCalculationSettings({ mode });
      }
      // CommandRuntime invokes listeners after the mutation handler.  Throwing
      // here still causes the command transaction to run its inverse, so a
      // direct write into a dynamic-array child cannot leave partial model or
      // formula state behind.  Undo/redo replay is allowed to restore the
      // exact prior snapshot.
      if (source === 'command' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
        assertNoSpillChildWrite(runtime.model, mutation);
      }

      runtime.pendingPivotMutations.push(structuredClone(mutation));
      if (mutation.id === 'dataSource.add' || mutation.id === 'dataSource.update' || mutation.id === 'dataSource.remove'
        || mutation.id === 'dataRegion.add' || mutation.id === 'dataRegion.remove'
        || mutation.id === 'dataRegion.materialize.commit' || mutation.id === 'dataRegion.materialize.restore'
        || mutation.id === 'query.load.range' || mutation.id === 'query.load.sheet-table'
        || mutation.id === 'query.load.pivot-source' || mutation.id === 'query.load.workbook-table') {
        initializeDataContent(runtime);
      }
      if (FORMULA_SYNC_MUTATIONS.has(mutation.id)) {
        const formulaInputChanged = DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)
          ? synchronizeManualCellMutation(runtime.formula, runtime.model, mutation)
          : false;
        if (FORMULA_METADATA_MUTATIONS.has(mutation.id)) {
          if (mutation.id === 'name.set' || mutation.id === 'name.remove') runtime.formula.setDefinedNameModels(runtime.model.definedNameModels);
          syncWorkbookSheetTables(runtime.formula, runtime.model);
          const touchedSheets = new Set(mutation.affectedRanges.map((range) => range.sheetId));
          for (const sheetId of touchedSheets) configureFormulaSpillEnvironment(runtime.formula, runtime.model.getSheet(sheetId));
        }
        void scheduleFormulaRecalculation(runtime, VISIBILITY_MUTATIONS.has(mutation.id) || formulaInputChanged);
      }
      if (FIND_INDEX_MUTATIONS.has(mutation.id)) {
        if ((mutation.id === 'sheet.add' || mutation.id === 'sheet.restore' || mutation.id === 'sheet.duplicated') && mutation.sheetId) runtime.findIndex.rebuildSheet(mutation.sheetId);
        else if (mutation.id === 'sheet.remove' && mutation.sheetId) runtime.findIndex.removeSheet(mutation.sheetId);
        else for (const sheetId of new Set(mutation.affectedRanges.map((range) => range.sheetId))) runtime.findIndex.rebuildSheet(sheetId);
      }
    }),
  );

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (runtime.disposed) return;
      if (source !== 'command') return;
      const durability = mutationCapability(mutation.id)?.durability;
      if (durability === 'transient') return;
      if (durability === 'local') {
        runtime.pendingLocalCheckpoint = true;
        return;
      }
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
      const localCheckpoint = runtime.pendingLocalCheckpoint;
      runtime.pendingLocalCheckpoint = false;
      if (batch.length === 0 && !localCheckpoint) return;
      runtime.handlers.onMutationsApplied?.();
      const history = runtime.commands.getUndoEntries().find((entry) => entry.operationId === result.operationId);
      if (history) {
        runtime.collaboration?.recordLocalUndo({
          operationId: result.operationId,
          undoMutations: history.inversePlan,
        });
      }
      if (batch.length > 0) {
        if (runtime.collaboration) submitChangeset(runtime, result.operationId, batch);
        else runtime.pendingLocalOperations.push({ operationId: result.operationId, mutations: batch });
      }
      void runtime.checkpointWorkspace();
    }),
  );

  runtime.detachers.push(
    runtime.commands.onHistoryReplay((source, entry) => {
      if (runtime.disposed) return;
      // Undo/redo has no command-listener completion callback. Publish the
      // replayed mutation facts through the same session coordinator boundary
      // used by local and remote command application.
      runtime.handlers.onMutationsApplied?.();
      if (!runtime.collaboration || entry.inversePlan.length === 0) return;
      const operation = source === 'undo'
        ? runtime.collaboration.enqueueCompensatingMutations(
          runtime.collaboration.undoOwnLast() ?? entry.inversePlan,
          runtime.model.unitId,
          entry.operationId,
          entry.baseRevision,
        )
        : runtime.collaboration.enqueueLocalMutations(entry.forwardMutations, runtime.model.unitId);
      if (source === 'redo') {
        runtime.collaboration.recordLocalUndo({ operationId: entry.operationId, undoMutations: entry.inversePlan });
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
  runtime.pendingLocalCheckpoint = false;
  runtime.pendingPivotMutations = [];
}

function replaceCollaborationSession(runtime: SpreadsheetRuntime, record: WorkspaceRecord | null, options: { deferRevision?: boolean } = {}): void {
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
      void enqueuePersistenceWrite(runtime, async () => {
        const current = await runtime.workspacePersistence.load(runtime.model.unitId);
        const storageRevision = await runtime.workspacePersistence.commitOperationJournal(
          runtime.model.unitId,
          operations,
          sequence,
          current?.storageRevision,
        );
        if (runtime.workspaceRecord) runtime.workspaceRecord.storageRevision = storageRevision;
        runtime.handlers.onWorkspacePersisted?.();
      }).catch((error: unknown) => publishPersistenceFailure(runtime, error));
    },
  });
  if (!options.deferRevision) runtime.collaboration.setRevision(runtime.remoteRevision);
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
  runtime.formula = rebuildFormulaEngine(runtime.model, runtime.dateSystem, runtime.canonicalReferenceDate, runtime.rowVisibilityResolver, runtime.calculationSessionPort);
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
  if (revision != null) {
    runtime.remoteRevision = revision;
    runtime.collaboration?.setRevision(revision);
  }
  runtime.pivotResults = {};
  void scheduleFormulaRecalculation(runtime);
}

export function setRuntimeDateContext(runtime: SpreadsheetRuntime, dateSystem: ExcelDateSystem, canonicalReferenceDate?: CanonicalExcelDateParts): void {
  runtime.dateSystem = dateSystem;
  runtime.canonicalReferenceDate = canonicalReferenceDate ? structuredClone(canonicalReferenceDate) : runtime.canonicalReferenceDate;
  runtime.formula.disposeCalculationTasks();
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(runtime.model, runtime.dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return runtime.formula?.getCellResult(address)?.value;
  });
  runtime.formula = rebuildFormulaEngine(runtime.model, runtime.dateSystem, runtime.canonicalReferenceDate, runtime.rowVisibilityResolver, runtime.calculationSessionPort);
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
  void scheduleFormulaRecalculation(runtime);
}

function rebuildFormulaEngine(workbook: WorkbookModel, dateSystem: ExcelDateSystem = '1900', canonicalReferenceDate?: CanonicalExcelDateParts, rowVisibilityResolver?: WorkbookRowVisibilityResolver, calculationSessionPort?: CalculationSessionPort): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.primarySheetId, dateSystem, canonicalReferenceDate, collationContext: workbook.collationContext, calculationSettings: workbook.calculationSettings, rowVisibilityResolver, calculationSessionPort });
  loadFormulaInputs(engine, workbook);
  return engine;
}

export function hydrateRuntime(runtime: SpreadsheetRuntime, response: SnapshotResponse, options: { deferCollaborationRevision?: boolean } = {}): void {
  if (runtime.disposed) return;
  const workbook = WorkbookModel.fromSnapshot(response.snapshot);
  detachCoreListeners(runtime);
  runtime.formula.disposeCalculationTasks();
  runtime.model = workbook;
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(workbook, runtime.dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return runtime.formula?.getCellResult(address)?.value;
  });
  runtime.commands = new CommandRuntime(workbook);
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  runtime.featureRuntime.dispose();
  registerSpreadsheetFeatures(runtime.commands, runtime.drawing, runtime.featureRuntime);
  activateSpreadsheetFeatures(runtime.featureRuntime, { documentType: 'spreadsheet', environment: typeof window === 'undefined' ? 'worker' : 'browser' });
  runtime.featureRuntime.advance('ready');
  runtime.formula = rebuildFormulaEngine(workbook, runtime.dateSystem, runtime.canonicalReferenceDate, runtime.rowVisibilityResolver, runtime.calculationSessionPort);
  runtime.findIndex = new FindIndex(workbook, (sheet, row, column) => runtime.formula.getCellValue({ sheetId: sheet.id, row, column }));
  installCommandCellValueResolver(runtime);
  runtime.formulaAudit.setFormula(runtime.formula);
  registerFormulaAuditCommands(runtime.commands.registry, runtime.formulaAudit);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  if (!options.deferCollaborationRevision) runtime.collaboration?.setRevision(response.revision);
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
      if (!runtime.disposed) {
        runtime.handlers.onDataSourceContentChanged?.(manifest.id);
        runtime.handlers.onMutationsApplied?.();
      }
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
  const prepared: Array<{ operationId: string; baseRevision: number; revision?: number; items: MutationInfo[] }> = [];
  // Resolve every operation before applying the first one. A malformed
  // pending envelope must not leave a prefix of the journal applied.
  let applied = 0;
  for (const operation of pending) {
    if (operation.unitId !== runtime.model.unitId) throw new Error(`PENDING_OPERATION_UNIT_MISMATCH: ${operation.operationId}`);
    const items = operation.mutations.map((mutation) => {
      const metadata = runtime.commands.registry.getMutationMetadata(mutation.id);
      if (!metadata?.affectedRanges) throw new Error(`PENDING_OPERATION_CONTRACT_MISSING: ${mutation.id}`);
      const resolved = metadata.affectedRanges.resolve(mutation.params as never);
      if (!Array.isArray(resolved)) throw new Error(`PENDING_OPERATION_RANGE_INVALID: ${mutation.id}`);
      const affectedRanges = [...resolved];
      return {
        id: mutation.id,
        unitId: operation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges,
      } satisfies MutationInfo;
    });
    prepared.push({ operationId: operation.operationId, baseRevision: operation.baseRevision, items });
  }
  for (const operation of prepared) {
    if (operation.items.length === 0) throw new Error(`PENDING_OPERATION_EMPTY: ${operation.operationId}`);
    runtime.commands.applyRemoteMutations(operation.items, { operationId: operation.operationId, baseRevision: operation.baseRevision });
    applied += 1;
  }
  return applied;
}

async function loadHistoryAndReplayPending(runtime: SpreadsheetRuntime): Promise<void> {
  let revisions: import('@react-sheets/protocol').RevisionRecord[];
  try {
    revisions = await runtime.api.listRevisions(runtime.model.unitId);
  } catch (cause) {
    throw new Error(`HISTORY_LOAD_FAILED: unable to load committed workbook history before replay (${cause instanceof Error ? cause.message : 'request failed'})`);
  }
  const ordered = [...revisions].sort((left, right) => left.revision - right.revision);
  const duplicate = ordered.some((record, index) => index > 0 && record.revision === ordered[index - 1]!.revision);
  const highest = ordered.at(-1)?.revision ?? 0;
  const expected = runtime.remoteRevision;
  if (duplicate || (ordered.length > 0 && ordered[0]!.revision !== 1) || (expected > 0 && highest < expected)) {
    throw new Error(`HISTORY_GAP: committed history ends at revision ${highest}, but snapshot requires ${expected}`);
  }
  for (let index = 1; index < ordered.length; index += 1) {
    if (ordered[index]!.revision !== ordered[index - 1]!.revision + 1) {
      throw new Error(`HISTORY_GAP: missing committed revision between ${ordered[index - 1]!.revision} and ${ordered[index]!.revision}`);
    }
  }
  try {
    runtime.collaboration?.loadCommittedHistory(ordered.map((record) => record.payload));
  } catch (cause) {
    throw new Error(`HISTORY_LOAD_FAILED: committed history cannot be classified for rebase (${cause instanceof Error ? cause.message : 'invalid history'})`);
  }
  runtime.handlers.onRemoteRevisions?.(ordered);
  // Pending operations are replayed only after the complete, validated history
  // has been installed. A history error leaves them durable and unapplied.
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
    const parsedCollaborationUrl = new URL(runtime.collaborationUrl ?? (protocol + '://' + window.location.host + '/ws'));
    if (!['ws:', 'wss:'].includes(parsedCollaborationUrl.protocol) || parsedCollaborationUrl.username || parsedCollaborationUrl.password) {
      throw new Error('Collaboration URL must be an uncredentialed ws:// or wss:// URL');
    }
    const client = new CollabSocketClient(parsedCollaborationUrl.toString(), {
      authTokenProvider,
      shareTokenProvider,
    });
    runtime.collab = client;
    runtime.broadcastPresence = (state) => !runtime.disposed && client.send({ type: 'presence.updated', unitId: runtime.model.unitId, state });

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
        void runtime.api.listRevisions(runtime.model.unitId)
          .then((revs) => runtime.handlers.onRemoteRevisions?.(revs))
          .catch((error: unknown) => publishRuntimeFailure(runtime, {
            code: 'HISTORY_LOAD_FAILED',
            message: error instanceof Error ? error.message : 'Committed history refresh failed',
            recovery: 'Retry the history refresh before submitting structural pending operations.',
            cause: error,
          }));
      } else if (message.type === 'cursor.broadcast' || message.type === 'presence.broadcast') {
        if (!message.unitId || message.unitId !== runtime.model.unitId) return;
        if (message.type === 'presence.broadcast' && (message.state as { status?: string } | null)?.status === 'offline') {
          runtime.handlers.onPeersChange?.([]);
          runtime.collaboration?.presence.removeUser(message.actorId);
          return;
        }
        const cursorState = message.state as { row?: number; column?: number; name?: string; sheetId?: string; edit?: { sheetId: string; row: number; column: number; status: 'enter' | 'edit' | 'point'; surface?: 'grid' | 'formula-bar' | 'formula-panel' } | null } | null;
        const peer = mapPeerCursor(message.actorId, cursorState, runtime.model.primarySheetId);
        runtime.collaboration?.presence.upsertUser({
          actorId: peer.actorId,
          displayName: peer.name,
          color: peer.color,
        });
        if (runtime.collaboration) updatePresenceFromPeer(runtime.collaboration, peer);
        if (cursorState?.edit) runtime.collaboration?.presence.updateEditSession({ actorId: message.actorId, ...cursorState.edit });
        else runtime.collaboration?.presence.clearEditSession(message.actorId);
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
      runtime.broadcastPresence = () => false;
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
  runtime.featureRuntime.dispose();
  for (const detach of runtime.dataContentDetachers) detach();
  runtime.dataContentDetachers = [];
  runtime.dataContent.clear();
  runtime.collaboration?.attachTransport(undefined);
  runtime.collaboration = null;
  runtime.collab = null;
}

async function initializePersistence(runtime: SpreadsheetRuntime, isActive: () => boolean): Promise<void> {
  const resolution = runtime.resolution;
  const localPendingBeforeLoad = runtime.collaboration?.getPendingOperations() ?? [];
  let localRecord: WorkspaceRecord | null = null;
  let resolvedRemote: SnapshotResponse | null = null;
  if (resolution) {
    if (resolution.unitId !== runtime.model.unitId) throw new Error('Workbook resolution unitId does not match runtime model');
    localRecord = resolution.localRecord ?? null;
    if (localRecord) runtime.operationJournal.hydrate(localRecord);
    if (resolution.mode === 'remote') {
      runtime.localOnly = false;
      runtime.remoteSyncRequested = true;
      resolvedRemote = { snapshot: structuredClone(resolution.snapshot), revision: resolution.revision };
    } else {
      if (!localRecord) throw new Error('Local workbook resolution is missing its WorkspaceRecord');
      runtime.localOnly = true;
      runtime.remoteSyncRequested = false;
    }
  } else {
    if (!runtime.localOnly && !(await hasValidRemoteBinding(runtime))) runtime.localOnly = true;
    try {
      localRecord = await runtime.workspacePersistence.load(runtime.model.unitId, runtime.assetStore);
      const canDiscoverLocalDefault = runtime.model.unitId === 'wb-local-default'
        && (typeof window === 'undefined' || !/^\/workbooks\/[^/]+\/?$/.test(window.location.pathname));
      if (!localRecord && canDiscoverLocalDefault) {
        const summaries = await runtime.workspacePersistence.list();
        const first = summaries[0];
        if (first) localRecord = await runtime.workspacePersistence.load(first.unitId, runtime.assetStore);
      }
    } catch (error) {
      publishPersistenceFailure(runtime, error);
      return;
    }
  }

  if (!isActive()) return;

  if (localRecord) {
    runtime.workspaceRecord = localRecord;
    runtime.localRevision = localRecord.localRevision;
    runtime.remoteRevision = resolution?.revision ?? localRecord.serverRevision;
    runtime.localOnly = runtime.localOnly || localRecord.syncMode === 'local-only';
    runtime.remoteSyncRequested = runtime.remoteSyncRequested || localRecord.syncMode === 'remote';
    if (resolution?.mode !== 'remote') {
      if (!isActive()) return;
      hydrateRuntime(runtime, {
        snapshot: resolution?.snapshot ?? localRecord.snapshot,
        revision: resolution?.revision ?? localRecord.serverRevision,
      });
      replaceCollaborationSession(runtime, localRecord);
      if (localPendingBeforeLoad.length > 0) replayPendingOperations(runtime, localPendingBeforeLoad);
      runtime.handlers.onNotice?.('Workbook restored from the current memory session');
    }
  }

  if (runtime.localOnly) {
    runtime.remoteConnected = false;
    runtime.handlers.onAccessRole?.(null);
    // A StrictMode dispose may have detached the collaboration journal even
    // though the workbook remains local-only. Recreate the journal owner
    // before the first post-remount command so edits remain durable.
    if (!runtime.collaboration) replaceCollaborationSession(runtime, localRecord);
    if (!runtime.workspaceRecord) {
      try {
        await checkpointStartupLocally(runtime);
      } catch (error) {
        publishPersistenceFailure(runtime, error);
        return;
      }
    }
    if (isActive()) {
      runtime.handlers.onSaveState?.(runtime.remoteSyncRequested ? 'offline' : 'saved');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
    }
    return;
  }

  try {
    const snapshotResponse = resolvedRemote ?? await runtime.api.getSnapshot(runtime.model.unitId);
    const access = resolution?.mode === 'remote' ? resolution.access : await runtime.api.getAccess(runtime.model.unitId);
    if (!access) throw new Error('Remote workbook resolution is missing access metadata');
    if (!isActive()) return;
    hydrateRuntime(runtime, { ...snapshotResponse, snapshot: await migrateLegacyImageAssets(snapshotResponse.snapshot, runtime.assetStore) }, { deferCollaborationRevision: true });
    runtime.remoteRevision = snapshotResponse.revision;
    runtime.localOnly = false;
    runtime.remoteSyncRequested = true;
    replaceCollaborationSession(runtime, localRecord, { deferRevision: true });
    await loadHistoryAndReplayPending(runtime);
    runtime.collaboration?.setRevision(runtime.remoteRevision);
    if (!isActive()) return;
    runtime.workspaceRecord = await runtime.workspacePersistence.checkpoint(
      runtime.model.snapshot(),
      runtime.localRevision,
      runtime.remoteRevision,
      'remote',
      undefined,
      resolution ? {
        location: resolution.binding.location,
        lifecycle: resolution.lifecycle,
        source: localRecord?.metadata.source ?? 'native',
        role: resolution.access?.role ?? localRecord?.metadata.role ?? 'viewer',
      } : undefined,
    );
    runtime.remoteConnected = true;
    runtime.handlers.onAccessRole?.(access.role);
    if (isActive()) {
      runtime.handlers.onSaveState?.('saved');
      runtime.handlers.onNotice?.('Workbook restored from server');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
      runtime.handlers.onWorkspacePersisted?.();
    }
  } catch (error) {
    // Authentication, authorization, unknown workbooks, transport failures
    // and history gaps are authoritative remote-session outcomes. None may be
    // disguised as a new local workbook with the same URL.
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
    // A remote workbook cannot silently become a local memory workbook. A
    // separate local-only copy is an explicit catalog operation; this runtime
    // never turns an authoritative remote failure into a ready local session.
    runtime.remoteConnected = false;
    runtime.handlers.onAccessRole?.(null);
    if (!isActive()) return;
    const failureCode: RuntimeFailureCode = error instanceof Error && error.message.startsWith('HISTORY_GAP')
      ? 'HISTORY_GAP'
      : error instanceof Error && error.message.startsWith('HISTORY_LOAD_FAILED')
        ? 'HISTORY_LOAD_FAILED'
        : 'REMOTE_WORKBOOK_UNAVAILABLE';
    publishRuntimeFailure(runtime, {
      code: failureCode,
      message: error instanceof Error ? error.message : 'Remote workbook initialization failed',
      recovery: 'Retry the authoritative server operation; create a separate local-only copy explicitly if offline work is required.',
      cause: error,
    });
  }
}

async function checkpointStartupLocally(runtime: SpreadsheetRuntime): Promise<void> {
  const localResolution = runtime.resolution?.mode === 'local' ? runtime.resolution : null;
  runtime.workspaceRecord = await runtime.workspacePersistence.checkpoint(
    runtime.model.snapshot(),
    runtime.localRevision,
    runtime.remoteRevision,
    'local-only',
    undefined,
    localResolution ? {
      location: localResolution.binding.location,
      lifecycle: localResolution.lifecycle,
      source: runtime.workspaceRecord?.metadata.source ?? 'native',
      role: runtime.workspaceRecord?.metadata.role ?? 'viewer',
    } : undefined,
  );
}

function publishPersistenceFailure(runtime: SpreadsheetRuntime, error: unknown): void {
  runtime.workspaceRecord = null;
  runtime.remoteConnected = false;
  runtime.handlers.onAccessRole?.(null);
  runtime.handlers.onSaveState?.('error');
  runtime.handlers.onPhaseChange?.('error');
  if (error instanceof WorkspaceStorageError) {
    runtime.handlers.onNotice?.(`${error.code}: ${error.message} ${error.recovery}`);
    return;
  }
  runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'STORAGE_TRANSACTION_FAILED: 本地工作簿持久化失败。');
}

function publishRuntimeFailure(runtime: SpreadsheetRuntime, failure: RuntimeFailure): void {
  if (runtime.disposed) return;
  runtime.handlers.onRuntimeFailure?.(failure);
  runtime.handlers.onSaveState?.('error');
  runtime.handlers.onPhaseChange?.('error');
  runtime.handlers.onNotice?.(`${failure.code}: ${failure.message} ${failure.recovery}`);
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
