import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry, type MutationInfo } from '@react-sheets/command-runtime';
import { canonicalExcelDateFromUtcDate, FormulaEngine, type CanonicalExcelDateParts, type CellAddressInput, type ExcelDateSystem } from '@react-sheets/formula-engine';
import {
  ApiRequestError,
  WorkbookApiClient,
  type AuthTokenProvider,
  type ShareTokenProvider,
  type WorkbookAclRole,
  type OperationMessage,
  mutationCapability,
} from '@react-sheets/protocol';
import { CollabSocketClient } from '@react-sheets/protocol';
import { FindIndex } from '@react-sheets/sheet-features';
import { activateSpreadsheetFeatures, createSpreadsheetFeatureRuntime, registerSpreadsheetFeatures, type SpreadsheetFeatureRuntime } from './feature-registry';
import { DrawingRuntime } from './features/drawing';
import { createDefaultConnectorRegistry, type ConnectorRegistry } from './features/query';
import { FormulaAuditController, registerFormulaAuditCommands } from './features/formula-audit';
import { DataSourceContentQuery } from './features/data-source';
import type { AnalyticsExecutor } from './features/query';
import { CollaborationSession } from './collaboration/collaboration-session';
import { createKernelWorkbookVisibilityResolver, createWorkbookRowVisibilityResolver, type WorkbookRowVisibilityResolver } from './formula-visibility';
import type { ResolvedVisibility } from '@react-sheets/sheet-features';
import { mapPeerCursor, updatePresenceFromPeer } from './collaboration';
import {
  WorkspacePersistence,
  DataBlockSynchronizer,
  RemoteAssetStore,
  type AssetStore,
  type WorkspacePersistenceOptions,
  type WorkspaceRecord,
  WorkspaceStorageError,
} from './features/persistence';
import { initializeKernel, kernelInvoke } from '../../kernel-client/src/index';
import type { WorkbookOpenResponse } from '@react-sheets/protocol';
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
  resolveVisibility: (sheet: import('@react-sheets/core-model').WorksheetModel) => ResolvedVisibility;
  formulaAudit: FormulaAuditController;
  dateSystem: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Node/SSR callers must supply an explicit persistent calculation session. */
  collaborationUrl?: string;
  model: WorkbookModel;
  commands: CommandRuntime;
  drawing: DrawingRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
  /** Local-durable geometry changed without producing a remote operation. */
  /** Mutation facts are drained by the refresh coordinator after each apply. */
  pendingPivotMutations: MutationInfo[];
  drainPivotMutations: () => MutationInfo[];
  detachers: Array<() => void>;
  handlers: RuntimeHandlers;
  ownOperationIds: Set<string>;
  nextClientSequence: number;
  kernelClientSequences: Map<string, number>;
  pivotResults: Record<string, import('@react-sheets/core-model').PivotResultTree>;
  pivotErrors: Record<string, import('./features/pivot/task-protocol').PivotTaskError>;
  collab: CollabSocketClient | null;
  collabDispose: (() => void) | null;
  broadcastPresence: (state: unknown) => boolean;
  collaboration: CollaborationSession | null;
  bootstrapDispose: (() => void) | null;
  workspacePersistence: WorkspacePersistence;
  dataBlocks: DataBlockSynchronizer;
  assetStore: AssetStore;
  dataContent: Map<string, DataSourceContentQuery>;
  dataContentDetachers: Array<() => void>;
  workspaceRecord: WorkspaceRecord | null;
  localRevision: number;
  /** Cloud authority is mandatory; this flag only tracks whether a save is awaiting server ack. */
  remoteSyncRequested: boolean;
  formulaCalculation: Promise<void>;
  persistenceReady: Promise<void>;
  checkpointWorkspace: () => Promise<void>;
  connectors: ConnectorRegistry;
  /** Revision-pinned Rust analytics boundary used by query/pivot/filter features. */
  analytics: AnalyticsExecutor;
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
  persistence?: WorkspacePersistenceOptions;
  workspacePersistence?: WorkspacePersistence;
  assetStore?: AssetStore;
  resolution?: WorkbookResolution;
  dateSystem?: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
  /** Node/SSR callers must supply an explicit persistent calculation session. */
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
  const analytics: AnalyticsExecutor = {
    execute: async ({ unitId: requestUnitId, revision, request }) => kernelInvoke('analytics.execute', { unitId: requestUnitId, revision, request }),
  };
  let formula: FormulaEngine | undefined;
  const resolveVisibility = createKernelWorkbookVisibilityResolver(model, () => model.revision);
  const rowVisibilityResolver = createWorkbookRowVisibilityResolver(model, resolveVisibility);
  formula = new FormulaEngine({ unitId: model.unitId, revision: () => model.revision, defaultSheetId: 'sheet-1' });
  const findIndex = new FindIndex(model, (sheet, row, column) => formula?.getCellValue({ sheetId: sheet.id, row, column }));
  const formulaAudit = new FormulaAuditController(formula);
  registerSpreadsheetFeatures(commands, drawing, featureRuntime);
  activateSpreadsheetFeatures(featureRuntime, { documentType: 'spreadsheet', environment: typeof window === 'undefined' ? 'worker' : 'browser' });
  featureRuntime.advance('ready');
  registerFormulaAuditCommands(commands.registry, formulaAudit);
  const workspacePersistence = options.workspacePersistence ?? new WorkspacePersistence({
    ...options.persistence,
    unitId: () => runtime?.model.unitId ?? model.unitId,
  });
  const api = options.api ?? new WorkbookApiClient({ authTokenProvider: options.authTokenProvider, shareTokenProvider: options.shareTokenProvider });
  let runtime!: SpreadsheetRuntime;
  const dataBlocks = new DataBlockSynchronizer(workspacePersistence.dataBlocks, api, {
    unitId: () => runtime.model.unitId,
    isRemoteAvailable: () => runtime.remoteConnected,
  });
  const assetStore = options.assetStore ?? new RemoteAssetStore(model.unitId, api);
  runtime = {
    api,
    formula: formula as FormulaEngine,
    rowVisibilityResolver,
    resolveVisibility,
    formulaAudit,
    dateSystem,
    canonicalReferenceDate,
    collaborationUrl: options.collaborationUrl,
    model,
    commands,
    drawing,
    remoteConnected: false,
    remoteRevision: 0,
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
    kernelClientSequences: new Map(),
    pivotResults: {},
    pivotErrors: {},
    collab: null,
    collabDispose: null,
    broadcastPresence: () => false,
    collaboration: null,
    bootstrapDispose: null,
    workspacePersistence,
    dataBlocks,
    assetStore,
    dataContent: new Map(),
    dataContentDetachers: [],
    workspaceRecord: null,
    localRevision: 0,
    remoteSyncRequested: true,
    formulaCalculation: Promise.resolve(),
    persistenceReady: Promise.resolve(),
    checkpointWorkspace: () => checkpointWorkspace(runtime),
    connectors,
    analytics,
    authTokenProvider: options.authTokenProvider,
    shareTokenProvider: options.shareTokenProvider,
    resolution: options.resolution,
    featureRuntime,
    findIndex,
    disposed: false,
  };
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  bindKernelCommitPort(runtime);
  // Draft operations remain in the active command session until the cloud ack.
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    loadPending: () => null,
    persistPending: () => {
      runtime.handlers.onSaveState?.('syncing');
    },
  });
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

const FIND_INDEX_MUTATIONS = new Set([
  ...DIRECT_CELL_WRITE_MUTATIONS,
  'sheet.add', 'sheet.remove', 'sheet.restore', 'sheet.duplicated',
  'find.replaced',
  'note.set', 'note.remove', 'note.visibility',
  'comment.add', 'comment.update', 'comment.remove', 'comment.reply', 'comment.reply.remove', 'comment.resolve',
]);

/** Formula state is owned by the revision-pinned kernel. Runtime only schedules a read. */
function synchronizeManualCellMutation(_engine: FormulaEngine, _workbook: WorkbookModel, _mutation: MutationInfo): boolean {
  return true;
}

/** The only command commit boundary: mutations are applied by the cloud kernel. */
function bindKernelCommitPort(runtime: SpreadsheetRuntime): void {
  runtime.commands.setCommitPort(async (request) => {
    const clientSequence = runtime.kernelClientSequences.get(request.operationId) ?? runtime.nextClientSequence + 1;
    runtime.kernelClientSequences.set(request.operationId, clientSequence);
    runtime.nextClientSequence = Math.max(runtime.nextClientSequence, clientSequence);
    const changeSet = await runtime.api.commitKernelOperation(runtime.model.unitId, { ...request, clientSequence });
    if (!Number.isSafeInteger(changeSet.revision) || changeSet.revision <= runtime.remoteRevision) {
      throw new Error('KERNEL_COMMIT_REVISION_INVALID: server acknowledgement is not newer than the current revision');
    }
    return changeSet;
  });
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
  if (!runtime.remoteConnected) return 'offline';
  return runtime.commands.activeDepth > 0 ? 'syncing' : 'saved';
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
      if (engine.getRecalculationMode() !== 'automatic' && !forceCalculation) return;

      runtime.handlers.onSaveState?.('calculating');
      try {
        await engine.recalculateAsync();
        if (runtime.disposed || epoch !== state.epoch || runtime.formula !== engine || runtime.model !== workbook) return;
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

async function checkpointWorkspace(runtime: SpreadsheetRuntime): Promise<void> {
  if (runtime.disposed || !runtime.remoteConnected) throw new Error('CLOUD_CONNECTION_REQUIRED');
  await runtime.commands.whenIdle();
  const checkpoint = await runtime.api.checkpointWorkbook(runtime.model.unitId);
  if (checkpoint.workbook.revision !== runtime.model.revision) throw new Error('KERNEL_CHECKPOINT_REVISION_MISMATCH');
}

export function attachCoreListeners(runtime: SpreadsheetRuntime): void {
  detachCoreListeners(runtime);

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (runtime.disposed) return;
      runtime.remoteRevision = runtime.model.revision;
      if (VISIBILITY_MUTATIONS.has(mutation.id)) {
        runtime.rowVisibilityResolver.invalidate();
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
        void scheduleFormulaRecalculation(runtime, VISIBILITY_MUTATIONS.has(mutation.id) || formulaInputChanged);
      }
      if (FIND_INDEX_MUTATIONS.has(mutation.id)) {
        if ((mutation.id === 'sheet.add' || mutation.id === 'sheet.restore' || mutation.id === 'sheet.duplicated') && mutation.sheetId) runtime.findIndex.rebuildSheet(mutation.sheetId);
        else if (mutation.id === 'sheet.remove' && mutation.sheetId) runtime.findIndex.removeSheet(mutation.sheetId);
        else for (const sheetId of new Set(mutation.affectedRanges.map((range) => range.sheetId))) runtime.findIndex.rebuildSheet(sheetId);
      }
    }),
  );

  runtime.detachers.push(runtime.commands.onCommand(() => {
    if (runtime.disposed) return;
    runtime.remoteRevision = runtime.model.revision;
    runtime.handlers.onMutationsApplied?.();
    runtime.handlers.onSaveState?.('saved');
  }));
  runtime.detachers.push(runtime.commands.onHistoryReplay(() => {
    if (runtime.disposed) return;
    runtime.remoteRevision = runtime.model.revision;
    runtime.handlers.onMutationsApplied?.();
    runtime.handlers.onSaveState?.('saved');
  }));
}

function detachCoreListeners(runtime: SpreadsheetRuntime): void {
  for (const detach of runtime.detachers) detach();
  runtime.detachers = [];
  runtime.pendingPivotMutations = [];
}

function replaceCollaborationSession(runtime: SpreadsheetRuntime, record: WorkspaceRecord | null, options: { deferRevision?: boolean } = {}): void {
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    loadPending: () => null,
    persistPending: () => {
      runtime.handlers.onSaveState?.('syncing');
    },
  });
  if (!options.deferRevision) runtime.collaboration.setRevision(runtime.remoteRevision);
}

export function rehydrateFormulaAfterRestore(runtime: SpreadsheetRuntime, revision?: number): void {
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

export function setRuntimeDateContext(runtime: SpreadsheetRuntime, dateSystem: ExcelDateSystem, canonicalReferenceDate?: CanonicalExcelDateParts): void {
  runtime.dateSystem = dateSystem;
  runtime.canonicalReferenceDate = canonicalReferenceDate ? structuredClone(canonicalReferenceDate) : runtime.canonicalReferenceDate;
  runtime.resolveVisibility = createKernelWorkbookVisibilityResolver(runtime.model, () => runtime.model.revision);
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(runtime.model, runtime.resolveVisibility);
  runtime.formulaAudit.refresh();
}

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  return new FormulaEngine({ unitId: workbook.unitId, revision: () => workbook.revision, defaultSheetId: workbook.primarySheetId });
}

export function hydrateRuntime(runtime: SpreadsheetRuntime, response: WorkbookOpenResponse, options: { deferCollaborationRevision?: boolean } = {}): void {
  if (runtime.disposed) return;
  const workbook = runtime.model;
  workbook.applyCommittedManifest(response.manifest, response.pages);
  detachCoreListeners(runtime);
  runtime.model = workbook;
  runtime.resolveVisibility = createKernelWorkbookVisibilityResolver(workbook, () => workbook.revision);
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(workbook, runtime.resolveVisibility);
  runtime.commands = new CommandRuntime(workbook);
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  bindKernelCommitPort(runtime);
  runtime.featureRuntime.dispose();
  registerSpreadsheetFeatures(runtime.commands, runtime.drawing, runtime.featureRuntime);
  activateSpreadsheetFeatures(runtime.featureRuntime, { documentType: 'spreadsheet', environment: typeof window === 'undefined' ? 'worker' : 'browser' });
  runtime.featureRuntime.advance('ready');
  runtime.formula = rebuildFormulaEngine(workbook);
  runtime.dateSystem = runtime.formula.getDateSystem();
  runtime.canonicalReferenceDate = runtime.formula.getCanonicalReferenceDate();
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
  void scheduleFormulaRecalculation(runtime, true);
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
  void runtime; void operations;
  return 0;
}

async function loadHistoryAndReplayPending(runtime: SpreadsheetRuntime): Promise<void> {
  void runtime;
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
    if (!active || runtime.disposed) {
      runtime.handlers.onCollabStatus?.('closed');
      return;
    }
    runtime.collaboration ??= new CollaborationSession(runtime.commands);
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
      if (message.type === 'cursor.broadcast' || message.type === 'presence.broadcast') {
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
      if (status === 'closed') runtime.handlers.onSaveState?.('offline');
      else if (status === 'connecting') runtime.handlers.onSaveState?.('syncing');
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
  runtime.featureRuntime.dispose();
  for (const detach of runtime.dataContentDetachers) detach();
  runtime.dataContentDetachers = [];
  runtime.dataContent.clear();
  runtime.collaboration?.attachTransport(undefined);
  runtime.collaboration = null;
  runtime.collab = null;
}

async function initializePersistence(runtime: SpreadsheetRuntime, isActive: () => boolean): Promise<void> {
  await initializeKernel();
  if (!isActive()) return;
  const resolution = runtime.resolution;
  let localRecord: WorkspaceRecord | null = null;
  let resolvedRemote: WorkbookOpenResponse | null = null;
  if (resolution) {
    if (resolution.unitId !== runtime.model.unitId) throw new Error('Workbook resolution unitId does not match runtime model');
    localRecord = resolution.localRecord ?? null;
    const manifest = await runtime.api.getManifest(runtime.model.unitId, resolution.revision);
    resolvedRemote = { unitId: manifest.unitId, manifest, pages: [], revision: manifest.revision };
  } else {
    try {
      localRecord = await runtime.workspacePersistence.load(runtime.model.unitId, runtime.assetStore);
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
    runtime.remoteSyncRequested = true;
  }

  if (!resolvedRemote) {
    const manifest = await runtime.api.getManifest(runtime.model.unitId);
    resolvedRemote = { unitId: manifest.unitId, manifest, pages: [], revision: manifest.revision };
  }

  try {
    const snapshotResponse = resolvedRemote;
    const access = resolution?.mode === 'remote' ? resolution.access : await runtime.api.getAccess(runtime.model.unitId);
    if (!access) throw new Error('Remote workbook resolution is missing access metadata');
    if (!isActive()) return;
    hydrateRuntime(runtime, snapshotResponse, { deferCollaborationRevision: true });
    runtime.remoteRevision = snapshotResponse.manifest.revision;
    runtime.remoteSyncRequested = true;
    replaceCollaborationSession(runtime, localRecord, { deferRevision: true });
    runtime.collaboration?.setRevision(runtime.remoteRevision);
    if (!isActive()) return;
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
    // A remote workbook cannot silently become an in-memory authoritative session.
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
      recovery: 'Retry the authoritative server operation before editing or saving again.',
      cause: error,
    });
  }
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
  runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'STORAGE_TRANSACTION_FAILED: 云端工作簿持久化失败。');
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

export type { HistoryEntry };
