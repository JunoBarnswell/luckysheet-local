import { RecoveryJournal } from './features/persistence/recovery-journal';
import { CheckpointCoordinator } from './features/persistence/checkpoint-coordinator';
import { WorkbookModel, isWorkbookCalculationContextEffect, type CellData, type DataSourceManifest, type StructuralTransformResult } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry, type MutationInfo } from '@react-sheets/command-runtime';
import { canonicalExcelDateFromUtcDate, FormulaEngine, type CanonicalExcelDateParts, type CellAddressInput, type ExcelDateSystem, type CalculationInputUpdate } from '@react-sheets/formula-engine';
import {
  ApiRequestError,
  assertOperationResultMatches,
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
import { createWorkbookRowVisibilityResolver, type WorkbookRowVisibilityResolver } from './formula-visibility';
import { mapPeerCursor, updatePresenceFromPeer } from './collaboration';
import {
  configureFormulaSpillEnvironment,
  configureWorkbookSpillEnvironments,
  syncFormulaSpillsToSheet,
  syncWorkbookSheetTables,
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
  onCalculationApplied?: (addresses: readonly { readonly sheetId: string; readonly row: number; readonly column: number }[]) => void;
  onPhaseChange?: (phase: import('./types').AppPhase) => void;
  onActiveSheetChange?: (sheetId: string) => void;
  onRemoteRevisions?: (revisions: import('@react-sheets/protocol').RevisionRecord[]) => void;
  onCollabStatus?: (status: 'connecting' | 'open' | 'closed') => void;
  onAccessRole?: (role: WorkbookAclRole | null) => void;
  onPeersChange?: (peers: import('./types').PeerCursor[]) => void;
  onWorkspacePersisted?: () => void;
  onDataSourceContentChanged?: (sourceId: string) => void;
}

export interface SpreadsheetRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  rowVisibilityResolver: WorkbookRowVisibilityResolver;
  formulaAudit: FormulaAuditController;
  dateSystem: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
  collaborationUrl?: string;
  model: WorkbookModel;
  commands: CommandRuntime;
  drawing: DrawingRuntime;
  remoteConnected: boolean;
  /** REST data blocks remain readable after an authorized snapshot while WebSocket sync is connecting. */
  remoteDataAvailable: boolean;
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
  /** Set only when a new authoritative workbook model invalidates derived Pivot results. */
  pivotRehydrationPending: boolean;
  collab: CollabSocketClient | null;
  collabDispose: (() => void) | null;
  broadcastPresence: (state: unknown) => boolean;
  collaboration: CollaborationSession | null;
  bootstrapDispose: (() => void) | null;
  operationJournal: OperationJournalStore;
  workspacePersistence: WorkspacePersistence;
  recoveryJournal: RecoveryJournal | null;
  dataBlocks: DataBlockSynchronizer;
  assetStore: AssetStore;
  dataContent: Map<string, DataSourceContentQuery>;
  dataContentSubscriptions: Map<string, { manifest: DataSourceManifest; manifestRef: { current: DataSourceManifest }; unsubscribe: () => void }>;
  workspaceRecord: WorkspaceRecord | null;
  localRevision: number;
  localOnly: boolean;
  remoteSyncRequested: boolean;
  formulaCalculation: Promise<void>;
  persistenceReady: Promise<void>;
  pendingLocalOperations: Array<{ operationId: string; mutations: MutationInfo[] }>;
  checkpointWorkspace: (advanceLocalRevision?: boolean, artifact?: NativeDocumentArtifact) => Promise<void>;
  flushCheckpoint: () => Promise<void>;
  connectors: ConnectorRegistry;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  resolution?: WorkbookResolution;
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
  recoverySubject?: string;
  assetStore?: AssetStore;
  resolution?: WorkbookResolution;
  dateSystem?: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
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
  const connectors = createDefaultConnectorRegistry();
  let formula: FormulaEngine | undefined;
  const rowVisibilityResolver = createWorkbookRowVisibilityResolver(model, dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return formula?.getCellResult(address)?.value;
  });
  formula = new FormulaEngine({ defaultSheetId: 'sheet-1', sheetOrder: model.sheetOrder.map((id) => ({ id, name: model.getSheet(id).name })), dateSystem, canonicalReferenceDate, collationContext: model.collationContext, calculationSettings: model.calculationSettings, rowVisibilityResolver });
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
    isRemoteAvailable: () => !runtime.localOnly && (runtime.remoteDataAvailable || runtime.remoteConnected),
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
    remoteDataAvailable: false,
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
    pivotRehydrationPending: false,
    collab: null,
    collabDispose: null,
    broadcastPresence: () => false,
    collaboration: null,
    bootstrapDispose: null,
    operationJournal,
    workspacePersistence,
    recoveryJournal: options.recoverySubject && typeof window !== 'undefined' ? new RecoveryJournal(window.location.origin, options.recoverySubject, unitId) : null,
    dataBlocks,
    assetStore,
    dataContent: new Map(),
    dataContentSubscriptions: new Map(),
    workspaceRecord: null,
    localRevision: 0,
    localOnly: options.localOnly ?? (!options.authTokenProvider && !options.shareTokenProvider),
    remoteSyncRequested: Boolean(options.authTokenProvider || options.shareTokenProvider),
    formulaCalculation: Promise.resolve(),
    persistenceReady: Promise.resolve(),
    pendingLocalOperations: [],
    checkpointWorkspace: () => Promise.resolve(),
    flushCheckpoint: () => Promise.resolve(),
    connectors,
    authTokenProvider: options.authTokenProvider,
    shareTokenProvider: options.shareTokenProvider,
    resolution: options.resolution,
    disposed: false,
  };
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  runtime.commands.setStructuralReferenceOwnersProvider(() => runtime.formula.dependencies);
  // The initial runtime has no workbook snapshot boundary yet, but it still
  // needs a live spill environment so the first authored dynamic-array formula
  // can resolve without rebuilding the whole calculation engine.
  configureWorkbookSpillEnvironments(runtime.formula, runtime.model);
  // Recovery intent is written before HTTP submission. Java owns durable
  // workbook state; reconnection reconciles each original operation ID.
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    clientSessionId: runtime.recoveryJournal?.clientSessionId,
    loadPending: () => {
      const journal = operationJournal.read(runtime.model.unitId);
      return journal
        ? { operations: journal.operations, nextClientSequence: journal.nextClientSequence }
        : null;
    },
    persistPending: (operations, nextClientSequence) => {
      operationJournal.write(runtime.model.unitId, operations, nextClientSequence);
      if (runtime.recoveryJournal) void runtime.recoveryJournal.persist(operations).catch((error: Error) => publishPersistenceFailure(runtime, error));
    },
  });
  runtime.checkpointWorkspace = (advanceLocalRevision = true, artifact) => checkpointWorkspace(runtime, advanceLocalRevision, artifact);
  runtime.flushCheckpoint = () => serverCheckpoint(runtime).flush();
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
  'range.move',
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
  'sheetTable.autoFilter.set',
  'outline.group.toggle',
  'outline.showLevel',
]);

const VISIBILITY_MUTATIONS = new Set([
  'row.hidden', 'row.unhidden', 'rows.unhidden.all', 'rows.hidden.restore',
  'sheet.rows.visibility.set', 'sheet.rows.unhide.all',
  'autoFilter.set', 'autoFilter.remove', 'sheet.autoFilter.set', 'sheet.autoFilter.remove',
  'sheetTable.autoFilter.set', 'sheetTable.add', 'sheetTable.remove', 'sheetTable.update',
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
  'range.move',
  'cells.inserted',
  'cells.deleted',
  'cells.inserted.restore',
  'cells.deleted.restore',
  'dataRegion.materialize.commit',
  'dataRegion.materialize.restore',
  'query.load.range',
  'query.load.sheet-table',
  'query.load.pivot-source',
  'query.load.workbook-table',
]);

function calculationInputUpdate(
  sheetId: string,
  row: number,
  column: number,
  cell: CellData | undefined,
): CalculationInputUpdate {
  const address = { sheetId, row, column };
  const input = !cell || (cell.formula === undefined && cell.value == null)
    ? null
    : cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly
      ? { kind: 'formula' as const, formula: cell.formula }
      : { kind: 'value' as const, value: (cell.value ?? null) as never };
  return { address, input };
}

const STRUCTURAL_FORMULA_SOURCE_IDS = {
  preservedFormula: 'structural:preserved-formula',
  formulaProvenance: 'structural:formula-provenance',
  barcode: 'structural:barcode',
} as const;

function indexAuxiliaryFormulaOwners(engine: FormulaEngine, address: CellAddressInput, cell: CellData): void {
  if (cell.formulaMetadata?.preservedOnly && cell.formula !== undefined) {
    engine.setStructuralFormulaReference(address, STRUCTURAL_FORMULA_SOURCE_IDS.preservedFormula, cell.formula);
  }
  if (cell.formulaMetadata?.sourceFormula !== undefined) {
    engine.setStructuralFormulaReference(address, STRUCTURAL_FORMULA_SOURCE_IDS.formulaProvenance, cell.formulaMetadata.sourceFormula);
  }
  if (cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula') {
    engine.setStructuralFormulaReference(address, STRUCTURAL_FORMULA_SOURCE_IDS.barcode, cell.presentation.source.formula);
  }
}

function removeAuxiliaryFormulaOwnersInRange(
  engine: FormulaEngine,
  range: StructuralTransformResult['clearInputRanges'][number],
): void {
  for (const owner of engine.dependencies.getStructuralReferenceOwnersInRange(range.sheetId, range)) {
    engine.removeStructuralFormulaReference(owner.address, owner.sourceId);
  }
}

function rangeContainsOwner(
  range: StructuralTransformResult['clearInputRanges'][number],
  owner: StructuralTransformResult['rewrittenFormulaOwners'][number],
): boolean {
  return range.sheetId === owner.sheetId
    && range.startRow <= owner.row && range.endRow >= owner.row
    && range.startColumn <= owner.column && range.endColumn >= owner.column;
}

function reindexAuxiliaryFormulaOwnerAt(
  engine: FormulaEngine,
  workbook: WorkbookModel,
  owner: StructuralTransformResult['rewrittenFormulaOwners'][number],
): void {
  const point = {
    sheetId: owner.sheetId,
    startRow: owner.row,
    endRow: owner.row,
    startColumn: owner.column,
    endColumn: owner.column,
  };
  removeAuxiliaryFormulaOwnersInRange(engine, point);
  const cell = workbook.getSheet(owner.sheetId).cells.get(owner.row, owner.column);
  if (cell) indexAuxiliaryFormulaOwners(engine, owner, cell);
}

function synchronizeCellMutation(engine: FormulaEngine, workbook: WorkbookModel, mutation: MutationInfo): readonly CellAddressInput[] {
  const hadFormulaInputs = engine.getFormulaCount() > 0;
  const cleared = new Map<string, CalculationInputUpdate>();
  const populated = new Map<string, CalculationInputUpdate>();
  for (const range of mutation.affectedRanges) {
    const sheet = workbook.getSheet(range.sheetId);
    removeAuxiliaryFormulaOwnersInRange(engine, range);
    for (const address of engine.getInputAddressesInRange(range)) {
      cleared.set(`${address.sheetId}:${address.row}:${address.column}`, { address, input: null });
    }
    sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (cell, row, column) => {
      indexAuxiliaryFormulaOwners(engine, { sheetId: sheet.id, row, column }, cell);
      const update = calculationInputUpdate(sheet.id, row, column, cell);
      if (update.input !== null) populated.set(`${sheet.id}:${row}:${column}`, update);
    });
  }
  const updates = [...cleared.values(), ...populated.values()];
  const roots = [...engine.synchronizeInputs(updates)];
  if (!hadFormulaInputs && engine.getFormulaCount() > 0) {
    const ordinaryValues: CalculationInputUpdate[] = [];
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        const update = calculationInputUpdate(sheet.id, row, column, cell);
        if (update.input?.kind === 'value') ordinaryValues.push(update);
      });
    }
    roots.push(...engine.synchronizeInputs(ordinaryValues));
  }
  return [...new Map(roots.map((address) => [`${address.sheetId}:${address.row}:${address.column}`, address])).values()];
}

function isStructuralTransformResult(effect: unknown): effect is StructuralTransformResult {
  if (typeof effect !== 'object' || effect === null) return false;
  const result = effect as Partial<StructuralTransformResult>;
  return result.kind === 'structural-transform'
    && Array.isArray(result.clearInputRanges)
    && Array.isArray(result.populateInputRanges)
    && Array.isArray(result.rewrittenFormulaOwners);
}

function mutationTouchesFilterCriteria(workbook: WorkbookModel, ranges: MutationInfo['affectedRanges']): boolean {
  for (const range of ranges) {
    const sheet = workbook.getSheet(range.sheetId);
    const filters = [
      ...(sheet.autoFilter ? [sheet.autoFilter] : []),
      ...sheet.sheetTables.flatMap((table) => table.autoFilter ? [table.autoFilter] : []),
    ];
    for (const filter of filters) {
      if (range.startRow > filter.range.endRow || range.endRow < filter.range.startRow) continue;
      if (Object.values(filter.columns).some(({ column, criterion }) => criterion !== undefined
        && column >= range.startColumn && column <= range.endColumn)) return true;
    }
  }
  return false;
}

function synchronizeStructuralMutation(
  runtime: SpreadsheetRuntime,
  mutation: MutationInfo,
  effect: StructuralTransformResult,
): readonly CellAddressInput[] {
  const engine = runtime.formula;
  const workbook = runtime.model;
  const cleared = new Map<string, CalculationInputUpdate>();
  const populated = new Map<string, CalculationInputUpdate>();
  const keyOf = (address: { readonly sheetId: string; readonly row: number; readonly column: number }): string =>
    `${address.sheetId}:${address.row}:${address.column}`;

  for (const range of effect.clearInputRanges) {
    removeAuxiliaryFormulaOwnersInRange(engine, range);
    for (const address of engine.getInputAddressesInRange(range)) {
      cleared.set(keyOf(address), { address, input: null });
    }
  }
  for (const range of effect.populateInputRanges) {
    const sheet = workbook.getSheet(range.sheetId);
    sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (cell, row, column) => {
      indexAuxiliaryFormulaOwners(engine, { sheetId: sheet.id, row, column }, cell);
      const update = calculationInputUpdate(sheet.id, row, column, cell);
      if (update.input !== null) populated.set(keyOf(update.address), update);
    });
  }
  for (const owner of effect.rewrittenFormulaOwners) {
    const ownerKey = keyOf(owner);
    if (effect.populateInputRanges.some((range) => rangeContainsOwner(range, owner))) continue;
    const sheet = workbook.getSheet(owner.sheetId);
    const cell = sheet.cells.get(owner.row, owner.column);
    if (!cell) throw new Error(`STRUCTURAL_PATCH_INVARIANT: rewritten reference owner ${owner.sheetId}!${owner.row}:${owner.column} is not a live cell`);
    const update = calculationInputUpdate(owner.sheetId, owner.row, owner.column, cell);
    if (update.input !== null && !populated.has(ownerKey)) populated.set(ownerKey, update);
    reindexAuxiliaryFormulaOwnerAt(engine, workbook, owner);
  }

  const hadFormulaInputs = engine.getFormulaCount() > 0;
  const roots = [...engine.synchronizeInputs([...cleared.values(), ...populated.values()])];
  if (effect.definedNameOwnerDeltas !== undefined) {
    engine.applyDefinedNameModelDeltas(effect.definedNameOwnerDeltas, false);
  } else {
    engine.setDefinedNameModels(workbook.definedNameModels, false);
  }
  syncWorkbookSheetTables(engine, workbook, false);
  configureFormulaSpillEnvironment(engine, workbook.getSheet(mutation.sheetId));
  if (!hadFormulaInputs && engine.getFormulaCount() > 0) {
    const ordinaryValues: CalculationInputUpdate[] = [];
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        const update = calculationInputUpdate(sheet.id, row, column, cell);
        if (update.input?.kind === 'value') ordinaryValues.push(update);
      });
    }
    roots.push(...engine.synchronizeInputs(ordinaryValues));
  }
  roots.push(...engine.getPendingRecalculationRoots());
  runtime.formulaAudit.refresh();
  return [...new Map(roots.map((address) => [`${address.sheetId}:${address.row}:${address.column}`, address])).values()];
}

/**
 * Load only formula inputs from the canonical workbook. The actual evaluation
 * is intentionally scheduled separately through FormulaEngine.recalculateAsync
 * so browser calculation stays in its Worker.
 */
function loadFormulaInputs(engine: FormulaEngine, workbook: WorkbookModel): number {
  const mode = engine.getRecalculationMode();
  engine.cancelCalculation();
  engine.reset();
  engine.setRecalculationMode('manual');
  engine.setDefinedNameModels(workbook.definedNameModels);
  configureWorkbookSpillEnvironments(engine, workbook);
  syncWorkbookSheetTables(engine, workbook);
  let formulaCount = 0;
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEachFormulaOwner((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined) {
        if (cell.formulaMetadata?.preservedOnly) {
          engine.setStructuralFormulaReference(address, STRUCTURAL_FORMULA_SOURCE_IDS.preservedFormula, cell.formula);
        } else {
          formulaCount += 1;
          engine.setFormula(address, cell.formula);
        }
      }
      indexAuxiliaryFormulaOwners(engine, address, cell);
    });
  }
  // A value-only workbook has no formula dependency graph. Keeping tens of
  // thousands of ordinary cells in FormulaEngine duplicates CellMatrix and
  // makes native-document open proportional to every imported value for no calculation
  // benefit. Formula workbooks retain the complete existing input contract.
  if (formulaCount > 0) {
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        const update = calculationInputUpdate(sheet.id, row, column, cell);
        if (update.input?.kind === 'value') engine.setValue(update.address, update.input.value);
      });
    }
  }
  engine.setRecalculationMode(mode);
  return formulaCount;
}

interface FormulaQueueState {
  tail: Promise<void>;
  scheduled: boolean;
  epoch: number;
  force: boolean;
  full: boolean;
  roots?: readonly CellAddressInput[];
}

const formulaQueueStates = new WeakMap<SpreadsheetRuntime, FormulaQueueState>();

function localFormulaIdleState(runtime: SpreadsheetRuntime): import('./types').SaveState {
  if (runtime.localOnly) return runtime.remoteSyncRequested ? 'offline' : 'saved';
  if (runtime.collaboration?.offlineQueue.getState() === 'error') return 'conflict';
  if (!runtime.remoteConnected) return 'offline';
  return runtime.collaboration?.offlineQueue.getPendingCount() ? 'syncing' : 'saved';
}

/**
 * Coalesce formula input changes into one Worker task. A new mutation cancels
 * the active task and advances the epoch, so a late worker result cannot
 * mutate spills or render projections for an older workbook state.
 */
export function scheduleFormulaRecalculation(runtime: SpreadsheetRuntime, force = false, roots?: readonly CellAddressInput[], full = false): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  const state = formulaQueueStates.get(runtime) ?? {
    tail: Promise.resolve(),
    scheduled: false,
    epoch: 0,
    force: false,
    full: false,
    roots: undefined,
  } satisfies FormulaQueueState;
  formulaQueueStates.set(runtime, state);
  state.epoch += 1;
  state.force ||= force;
  state.full ||= full;
  if (roots !== undefined) {
    const merged = new Map<string, CellAddressInput>();
    for (const root of state.roots ?? []) merged.set(typeof root === 'string' ? root : `${root.sheetId}:${root.row}:${root.column}`, root);
    for (const root of roots) merged.set(typeof root === 'string' ? root : `${root.sheetId}:${root.row}:${root.column}`, root);
    state.roots = [...merged.values()];
  } else {
    state.roots = undefined;
  }
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
      const fullCalculation = state.full;
      const requestedRoots = state.roots;
      state.force = false;
      state.full = false;
      state.roots = undefined;
      const engine = runtime.formula;
      const workbook = runtime.model;
      const calculationRoots = requestedRoots;
      const formulaCount = engine.getFormulaCount();
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
        const report = await engine.recalculateAsync(calculationRoots, undefined, fullCalculation);
        if (runtime.disposed || epoch !== state.epoch || runtime.formula !== engine || runtime.model !== workbook) return;
        const changedAddresses = report.changedAddresses ?? [];
        const affectedSheetIds = new Set(changedAddresses.map((address) => address.sheetId));
        for (const sheetId of affectedSheetIds) {
          syncFormulaSpillsToSheet(engine, workbook.getSheet(sheetId));
        }
        void checkpointWorkspace(runtime, false).catch((error: unknown) => {
          runtime.handlers.onSaveState?.('error');
          runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Local formula checkpoint failed');
        });
        runtime.handlers.onCalculationApplied?.(changedAddresses);
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
      if (spill.state !== 'ok') continue;
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
const serverCheckpoints = new WeakMap<SpreadsheetRuntime, CheckpointCoordinator>();
function serverCheckpoint(runtime: SpreadsheetRuntime): CheckpointCoordinator {
  let coordinator = serverCheckpoints.get(runtime);
  if (!coordinator) {
    coordinator = new CheckpointCoordinator(
      () => runtime.api.checkpointWorkbook(runtime.model.unitId),
      async revision => { await runtime.recoveryJournal?.coverCheckpoint(revision); },
      error => {
        runtime.handlers.onSaveState?.('error');
        runtime.handlers.onNotice?.(`CHECKPOINT_FAILED: ${error.message}；已提交的数据仍在服务器，恢复日志已保留。请使用保存重试。`);
      },
    );
    serverCheckpoints.set(runtime, coordinator);
  }
  return coordinator;
}
const persistenceWriteChains = new WeakMap<SpreadsheetRuntime, Promise<void>>();

interface LocalPersistenceState {
  journalChecksum: string | null;
  queuedJournalChecksum: string | null;
  storageRevision: number | null;
  snapshotLocalRevision: number;
  snapshotTimer: ReturnType<typeof setTimeout> | null;
}

const localPersistenceStates = new WeakMap<SpreadsheetRuntime, LocalPersistenceState>();
const LOCAL_SNAPSHOT_CHECKPOINT_REVISION_LIMIT = 50;

export function isLocalSnapshotCheckpointDue(localRevision: number, snapshotLocalRevision: number): boolean {
  return localRevision - snapshotLocalRevision >= LOCAL_SNAPSHOT_CHECKPOINT_REVISION_LIMIT;
}

function localPersistenceState(runtime: SpreadsheetRuntime): LocalPersistenceState {
  const existing = localPersistenceStates.get(runtime);
  if (existing) return existing;
  const created: LocalPersistenceState = {
    journalChecksum: null,
    queuedJournalChecksum: null,
    storageRevision: runtime.workspaceRecord?.storageRevision ?? null,
    snapshotLocalRevision: runtime.workspaceRecord?.pending.snapshotRevision ?? runtime.localRevision,
    snapshotTimer: null,
  };
  localPersistenceStates.set(runtime, created);
  return created;
}

function adoptLocalSnapshotBaseline(runtime: SpreadsheetRuntime, record: WorkspaceRecord): void {
  const state = localPersistenceState(runtime);
  state.storageRevision = record.storageRevision;
  state.snapshotLocalRevision = record.pending.snapshotRevision;
  state.journalChecksum = record.pending.checksum;
  if (state.queuedJournalChecksum === record.pending.checksum) state.queuedJournalChecksum = null;
}

function enqueuePersistenceWrite<T>(runtime: SpreadsheetRuntime, operation: () => Promise<T>): Promise<T> {
  const previous = persistenceWriteChains.get(runtime) ?? Promise.resolve();
  const next = previous.catch(() => undefined).then(operation);
  persistenceWriteChains.set(runtime, next.then(() => undefined, () => undefined));
  return next;
}

function scheduleLocalSnapshotCheckpoint(runtime: SpreadsheetRuntime): void {
  if (runtime.disposed || !runtime.localOnly) return;
  const state = localPersistenceState(runtime);
  // The journal for this revision is durable; full workbook snapshots are periodic compaction, not per-edit durability.
  if (!isLocalSnapshotCheckpointDue(runtime.localRevision, state.snapshotLocalRevision)) return;
  if (state.snapshotTimer !== null) return;
  state.snapshotTimer = setTimeout(() => {
    state.snapshotTimer = null;
    void writeLocalSnapshotCheckpoint(runtime).catch((error: unknown) => {
      if (runtime.disposed) return;
      runtime.handlers.onSaveState?.('error');
      runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Local snapshot checkpoint failed');
    });
  }, 1000);
}

function cancelLocalSnapshotCheckpoint(runtime: SpreadsheetRuntime): void {
  const state = localPersistenceState(runtime);
  if (state.snapshotTimer === null) return;
  clearTimeout(state.snapshotTimer);
  state.snapshotTimer = null;
}

function commitLocalOperationJournal(runtime: SpreadsheetRuntime, pendingJournal: NonNullable<ReturnType<OperationJournalStore['read']>>): Promise<void> {
  const state = localPersistenceState(runtime);
  if (state.journalChecksum === pendingJournal.checksum || state.queuedJournalChecksum === pendingJournal.checksum) {
    scheduleLocalSnapshotCheckpoint(runtime);
    return Promise.resolve();
  }
  state.queuedJournalChecksum = pendingJournal.checksum;
  const localRevision = runtime.localRevision;
  const previous = checkpointChains.get(runtime) ?? Promise.resolve();
  const next = previous
    .catch(() => undefined)
    .then(() => enqueuePersistenceWrite(runtime, async () => {
      // A root operation already committed to this queue must survive a runtime unmount.
      const expectedStorageRevision = state.storageRevision ?? runtime.workspaceRecord?.storageRevision;
      const storageRevision = await runtime.workspacePersistence.commitOperationJournal(
        runtime.model.unitId,
        pendingJournal.operations,
        pendingJournal.nextClientSequence,
        expectedStorageRevision,
        localRevision,
      );
      state.storageRevision = storageRevision;
      state.journalChecksum = pendingJournal.checksum;
      if (state.queuedJournalChecksum === pendingJournal.checksum) state.queuedJournalChecksum = null;
      if (runtime.workspaceRecord) {
        runtime.workspaceRecord = {
          ...runtime.workspaceRecord,
          localRevision,
          storageRevision,
          pending: structuredClone(pendingJournal),
          updatedAt: new Date().toISOString(),
        };
      }
      if (!runtime.disposed) {
        scheduleLocalSnapshotCheckpoint(runtime);
        runtime.handlers.onWorkspacePersisted?.();
      }
    }));
  checkpointChains.set(runtime, next);
  void next.catch(() => {
    if (state.queuedJournalChecksum === pendingJournal.checksum) state.queuedJournalChecksum = null;
  });
  return next;
}

function writeLocalSnapshotCheckpoint(runtime: SpreadsheetRuntime, artifact?: NativeDocumentArtifact): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  cancelLocalSnapshotCheckpoint(runtime);
  const snapshot = runtime.model.snapshot();
  const localRevision = runtime.localRevision;
  const serverRevision = runtime.remoteRevision;
  const resolution = runtime.resolution;
  const syncMode = resolution?.binding.syncMode ?? 'local-only';
  const metadata = resolution ? {
    location: resolution.binding.location,
    lifecycle: resolution.lifecycle,
    source: runtime.workspaceRecord?.metadata.source ?? 'native' as const,
    role: resolution.access?.role ?? runtime.workspaceRecord?.metadata.role ?? 'viewer' as const,
  } : undefined;
  const pendingJournal = runtime.operationJournal.read(runtime.model.unitId);
  const previous = checkpointChains.get(runtime) ?? Promise.resolve();
  const next = previous
    .catch(() => undefined)
    .then(() => enqueuePersistenceWrite(runtime, async () => {
      if (runtime.disposed) return;
      const record = artifact
        ? await runtime.workspacePersistence.checkpointWithArtifact(snapshot, localRevision, serverRevision, syncMode, artifact, pendingJournal, metadata)
        : await runtime.workspacePersistence.checkpoint(snapshot, localRevision, serverRevision, syncMode, pendingJournal, metadata);
      runtime.operationJournal.adoptSnapshotCheckpoint(runtime.model.unitId, record.pending);
      runtime.workspaceRecord = record;
      adoptLocalSnapshotBaseline(runtime, record);
      if (runtime.disposed) return;
      await runtime.assetStore.reconcile(collectAssetReferences(snapshot, [
        ...record.pending.operations,
        ...runtime.commands.getUndoEntries(),
        ...runtime.commands.getRedoEntries(),
      ]));
      if (runtime.disposed) return;
      runtime.handlers.onWorkspacePersisted?.();
    }));
  checkpointChains.set(runtime, next);
  return next;
}

function checkpointWorkspace(runtime: SpreadsheetRuntime, advanceLocalRevision = true, artifact?: NativeDocumentArtifact): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  if (advanceLocalRevision) runtime.localRevision += 1;
  if (!runtime.localOnly) {
    return (async () => {
      await runtime.recoveryJournal?.flushed();
      if (artifact) throw new Error('ARTIFACT_SAVE_OWNER_REQUIRED: 原生文件必须通过版本校验的保存命令提交');
      runtime.handlers.onWorkspacePersisted?.();
    })();
  }
  const pendingJournal = runtime.operationJournal.read(runtime.model.unitId);
  if (!artifact && runtime.workspaceRecord && pendingJournal && pendingJournal.operations.length > 0) {
    return commitLocalOperationJournal(runtime, pendingJournal);
  }
  return writeLocalSnapshotCheckpoint(runtime, artifact);
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
    runtime.commands.onMutation((mutation, source, appliedEffect) => {
      if (runtime.disposed) return;
      let structuralRoots: readonly CellAddressInput[] | undefined;
      const structuralEffect = isStructuralTransformResult(appliedEffect) ? appliedEffect : undefined;
      const calculationContextEffect = isWorkbookCalculationContextEffect(appliedEffect)
        ? appliedEffect
        : structuralEffect?.calculationContextEffect;
      const rebuildsCalculationContext = calculationContextEffect?.action === 'rebuild';
      const changesVisibilityProjection = VISIBILITY_MUTATIONS.has(mutation.id)
        || rebuildsCalculationContext
        || structuralEffect !== undefined
        || mutationTouchesFilterCriteria(runtime.model, mutation.affectedRanges);
      if (changesVisibilityProjection) runtime.rowVisibilityResolver.invalidate();
      if (mutation.id === 'workbook.calculation.mode.set') {
        const mode = (mutation.params as { mode?: unknown } | undefined)?.mode;
        if (mode !== 'automatic' && mode !== 'manual' && mode !== 'partial') throw new Error('Workbook calculation mode mutation is invalid');
        runtime.formula.setRecalculationMode(mode);
      }
      if (rebuildsCalculationContext) {
        // Add/remove/reorder and ambiguous rename references need a full
        // address-space refresh; exact rename/structural deltas stay incremental.
        rebuildFormulaCalculation(runtime);
        runtime.formula.notifyVisibilityChanged();
      } else if (structuralEffect) {
        if (mutation.id === 'sheet.rename') {
          runtime.formula.updateSheetNames(
            runtime.model.sheetOrder.map((id) => ({ id, name: runtime.model.getSheet(id).name })),
          );
        }
        structuralRoots = synchronizeStructuralMutation(runtime, mutation, structuralEffect);
        runtime.formula.notifyVisibilityChanged();
        structuralRoots = [...new Map([
          ...structuralRoots,
          ...runtime.formula.getPendingRecalculationRoots(),
        ].map((address) => [typeof address === 'string' ? address : `${address.sheetId}:${address.row}:${address.column}`, address])).values()];
      } else if (calculationContextEffect?.action === 'sync-defined-names') {
        runtime.formula.setDefinedNameModels(runtime.model.definedNameModels, false);
      } else if (calculationContextEffect?.action === 'sync-tables') {
        syncWorkbookSheetTables(runtime.formula, runtime.model, false);
      }
      if (changesVisibilityProjection && !rebuildsCalculationContext && !structuralEffect) {
        runtime.formula.notifyVisibilityChanged();
      }
      // CommandRuntime invokes listeners after the mutation handler.  Throwing
      // here still causes the command transaction to run its inverse, so a
      // direct write into a dynamic-array child cannot leave partial model or
      // formula state behind.  Undo/redo replay is allowed to restore the
      // exact prior snapshot.
      if (source === 'command' && DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id)) {
        assertNoSpillChildWrite(runtime.model, mutation);
      }

      const formulaOwnerDeltas = structuralEffect?.formulaOwnerDeltas ?? [];
      const projectionMutation = formulaOwnerDeltas.length > 0
        ? { ...mutation, structuralFormulaOwnerDeltas: [...formulaOwnerDeltas] }
        : mutation;
      runtime.pendingPivotMutations.push(structuredClone(projectionMutation));
      if (mutation.id === 'dataSource.add' || mutation.id === 'dataSource.update' || mutation.id === 'dataSource.remove'
        || mutation.id === 'dataRegion.add' || mutation.id === 'dataRegion.remove'
        || mutation.id === 'dataRegion.materialize.commit' || mutation.id === 'dataRegion.materialize.restore'
        || mutation.id === 'query.load.range' || mutation.id === 'query.load.sheet-table'
        || mutation.id === 'query.load.pivot-source' || mutation.id === 'query.load.workbook-table'
        || mutation.id === 'pivot.drilldown.add' || mutation.id === 'pivot.drilldown.remove') {
        initializeDataContent(runtime);
      }
      if (FORMULA_SYNC_MUTATIONS.has(mutation.id) || calculationContextEffect !== undefined) {
        const isDirectCellWrite = DIRECT_CELL_WRITE_MUTATIONS.has(mutation.id);
        const roots = rebuildsCalculationContext
          ? undefined
          : structuralRoots
            ?? (isDirectCellWrite ? synchronizeCellMutation(runtime.formula, runtime.model, mutation) : undefined)
            ?? (changesVisibilityProjection ? runtime.formula.getPendingRecalculationRoots() : undefined);
        const automatic = runtime.formula.getRecalculationMode() === 'automatic';
        if (automatic || changesVisibilityProjection || !isDirectCellWrite || rebuildsCalculationContext) {
          void scheduleFormulaRecalculation(
            runtime,
            VISIBILITY_MUTATIONS.has(mutation.id),
            isDirectCellWrite ? undefined : roots,
            false,
          );
        }
      } else if (changesVisibilityProjection) {
        void scheduleFormulaRecalculation(
          runtime,
          VISIBILITY_MUTATIONS.has(mutation.id),
          runtime.formula.getPendingRecalculationRoots(),
        );
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
        ...(mutation.structuralImpactRanges
          ? { structuralImpactRanges: [...mutation.structuralImpactRanges] }
          : {}),
      });
    }),
  );

  runtime.detachers.push(
    runtime.commands.onCommandAbort(() => {
      runtime.pendingMutations = [];
      runtime.pendingLocalCheckpoint = false;
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
      if (!runtime.collaboration) return;
      const replayMutations = source === 'undo' ? entry.inversePlan : entry.forwardMutations;
      if (replayMutations.length === 0) return;
      const operation = source === 'undo'
        ? runtime.collaboration.enqueueCompensatingMutations(
          replayMutations,
          runtime.model.unitId,
          entry.operationId,
          entry.baseRevision,
        )
        : runtime.collaboration.enqueueLocalMutations(replayMutations, runtime.model.unitId);
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

function replaceCollaborationSession(runtime: SpreadsheetRuntime, record: WorkspaceRecord | null): void {
  const existingPending = runtime.collaboration?.getPendingOperations() ?? [];
  const buffered = runtime.pendingLocalOperations.splice(0);
  const byId = new Map<string, import('@react-sheets/protocol').OperationEnvelope>();
  for (const operation of runtime.operationJournal.read(runtime.model.unitId)?.operations ?? record?.pending.operations ?? []) byId.set(operation.operationId, operation);
  for (const operation of existingPending) byId.set(operation.operationId, operation);
  const pending = [...byId.values()];
  const nextClientSequence = Math.max(
    record?.pending.nextClientSequence ?? 0,
    ...pending.map((operation) => operation.clientSequence),
  );
  runtime.operationJournal.write(runtime.model.unitId, pending, nextClientSequence, record?.pending.snapshotRevision);
  runtime.collaboration = new CollaborationSession(runtime.commands, {
    clientSessionId: runtime.recoveryJournal?.clientSessionId,
    loadPending: () => {
      const journal = runtime.operationJournal.read(runtime.model.unitId);
      return journal ? { operations: journal.operations, nextClientSequence: journal.nextClientSequence } : null;
    },
    persistPending: (operations, sequence) => {
      runtime.operationJournal.write(runtime.model.unitId, operations, sequence);
      if (runtime.recoveryJournal) void runtime.recoveryJournal.persist(operations).catch((error: Error) => publishPersistenceFailure(runtime, error));
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
  runtime.formula = rebuildFormulaEngine(runtime.model, runtime.dateSystem, runtime.canonicalReferenceDate);
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
  if (revision != null) {
    runtime.remoteRevision = revision;
    runtime.collaboration?.setRevision(revision);
  }
  runtime.pivotResults = {};
  runtime.pivotRehydrationPending = true;
  void scheduleFormulaRecalculation(runtime);
}

/** Explicit Ctrl+Alt+Shift+F9 boundary: rebuild the dependency context once. */
export function rebuildFormulaCalculation(runtime: SpreadsheetRuntime): void {
  runtime.formula.disposeCalculationTasks();
  runtime.formula = rebuildFormulaEngine(
    runtime.model,
    runtime.dateSystem,
    runtime.canonicalReferenceDate,
    runtime.rowVisibilityResolver,
  );
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
}

export function setRuntimeDateContext(runtime: SpreadsheetRuntime, dateSystem: ExcelDateSystem, canonicalReferenceDate?: CanonicalExcelDateParts): void {
  runtime.dateSystem = dateSystem;
  runtime.canonicalReferenceDate = canonicalReferenceDate ? structuredClone(canonicalReferenceDate) : runtime.canonicalReferenceDate;
  runtime.formula.disposeCalculationTasks();
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(runtime.model, runtime.dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return runtime.formula?.getCellResult(address)?.value;
  });
  runtime.formula = rebuildFormulaEngine(runtime.model, runtime.dateSystem, runtime.canonicalReferenceDate, runtime.rowVisibilityResolver);
  runtime.formulaAudit.setFormula(runtime.formula);
  runtime.formulaAudit.refresh();
  void scheduleFormulaRecalculation(runtime);
}

function rebuildFormulaEngine(workbook: WorkbookModel, dateSystem: ExcelDateSystem = '1900', canonicalReferenceDate?: CanonicalExcelDateParts, rowVisibilityResolver?: WorkbookRowVisibilityResolver): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.primarySheetId, sheetOrder: workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name })), dateSystem, canonicalReferenceDate, collationContext: workbook.collationContext, calculationSettings: workbook.calculationSettings, rowVisibilityResolver });
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
  runtime.rowVisibilityResolver = createWorkbookRowVisibilityResolver(workbook, runtime.dateSystem, (sheet, row, column) => {
    const address = { sheetId: sheet.id, row, column };
    return runtime.formula?.getCellResult(address)?.value;
  });
  const mutationGuard = runtime.commands.getMutationGuard();
  runtime.commands = new CommandRuntime(workbook);
  runtime.commands.setMutationGuard(mutationGuard);
  runtime.commands.setRevisionProvider(() => runtime.remoteRevision);
  registerSpreadsheetFeatures(runtime.commands, runtime.drawing);
  runtime.formula = rebuildFormulaEngine(workbook, runtime.dateSystem, runtime.canonicalReferenceDate, runtime.rowVisibilityResolver);
  runtime.commands.setStructuralReferenceOwnersProvider(() => runtime.formula.dependencies);
  installCommandCellValueResolver(runtime);
  runtime.formulaAudit.setFormula(runtime.formula);
  registerFormulaAuditCommands(runtime.commands.registry, runtime.formulaAudit);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  runtime.collaboration?.setRevision(response.revision);
  runtime.collaboration?.rebindCommands(runtime.commands);
  runtime.pivotResults = {};
  runtime.pivotRehydrationPending = true;
  initializeDataContent(runtime);
  void scheduleFormulaRecalculation(runtime);
}

function initializeDataContent(runtime: SpreadsheetRuntime): void {
  // Canonical mutations replace the affected manifest object. Keep readers
  // and decoded blocks for every unchanged source, including inactive sheets.
  for (const [sourceId, subscription] of runtime.dataContentSubscriptions) {
    const current = runtime.model.dataModel.sources.get(sourceId);
    if (current === subscription.manifest) continue;
    const query = runtime.dataContent.get(sourceId);
    if (current !== undefined && query?.rebindManifest(current)) {
      subscription.manifest = current;
      subscription.manifestRef.current = current;
      continue;
    }
    subscription.unsubscribe();
    runtime.dataContentSubscriptions.delete(sourceId);
    runtime.dataContent.delete(sourceId);
  }
  for (const manifest of runtime.model.dataModel.sources.values()) {
    if (runtime.dataContent.has(manifest.id)) continue;
    const manifestRef = { current: manifest };
    const query = new DataSourceContentQuery(manifest, {
      get: async (reference) => {
        const ref = manifestRef.current.blocks.find((block) => block.id === reference.id && block.dataSourceId === reference.dataSourceId && block.checksum === reference.checksum);
        if (!ref) return null;
        const bytes = await runtime.dataBlocks.get(ref);
        return { sourceId: ref.dataSourceId, blockId: ref.id, checksum: ref.checksum, bytes };
      },
    });
    let notificationScheduled = false;
    const notifyContentChanged = (): void => {
      if (notificationScheduled) return;
      notificationScheduled = true;
      // Block reads can be initiated by a synchronous render projection.  Do
      // not publish a React-facing refresh from that render stack; coalesce
      // loading/ready transitions from the same fetch wave into one task.
      setTimeout(() => {
        notificationScheduled = false;
        if (runtime.disposed || runtime.dataContent.get(manifest.id) !== query) return;
        runtime.handlers.onDataSourceContentChanged?.(manifest.id);
        runtime.handlers.onMutationsApplied?.();
      }, 0);
    };
    runtime.dataContentSubscriptions.set(manifest.id, { manifest, manifestRef, unsubscribe: query.subscribe(notifyContentChanged) });
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
      if (!metadata) throw new Error(`RECOVERY_MUTATION_UNSUPPORTED: ${mutation.id}`);
      const resolved = metadata.affectedRanges?.resolve(mutation.params as never);
      if (!Array.isArray(resolved)) throw new Error(`RECOVERY_RANGE_INVALID: ${mutation.id}`);
      const affectedRanges: MutationInfo['affectedRanges'] = [...resolved];
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
  const pending = runtime.collaboration?.getPendingOperations() ?? [];
  for (const operation of pending) {
    const result = await runtime.api.getOperationResult(runtime.model.unitId, operation.operationId);
    if (result) {
      assertOperationResultMatches(operation, result.operation);
      await runtime.recoveryJournal?.confirm(operation, result.operation.revision);
      runtime.collaboration?.acknowledge(operation.operationId, result.operation.revision);
      await runtime.recoveryJournal?.flushed();
      serverCheckpoint(runtime).request(result.operation.revision);
    }
    else if (operation.baseRevision !== runtime.remoteRevision) {
      throw new Error(`RECOVERY_REVISION_CONFLICT: ${operation.operationId}，恢复草稿保留，请核对服务器版本后处理`);
    }
  }
  await runtime.flushCheckpoint();
  const unresolvedSessions = new Set(runtime.collaboration?.getPendingOperations().map(operation => operation.clientSessionId));
  if (unresolvedSessions.size > 1) {
    throw new Error('RECOVERY_SESSION_CONFLICT: 多个已关闭页面存在未确认修改；草稿已保留，请分别核对恢复分支，不能自动混合提交');
  }
  const revisions = await runtime.api.listRevisions(runtime.model.unitId);
  runtime.collaboration?.loadCommittedHistory(revisions.map(record => record.payload));
  runtime.handlers.onRemoteRevisions?.(revisions);
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
        await runtime.recoveryJournal?.flushed();
        const existing = runtime.collaboration?.offlineQueue.requiresResultLookup(operation.operationId)
          ? await runtime.api.getOperationResult(runtime.model.unitId, operation.operationId) : null;
        const committed = existing ?? await runtime.api.commitOperation(runtime.model.unitId, operation);
        assertOperationResultMatches(operation, committed.operation);
        const revision = committed.operation.revision;
        await runtime.recoveryJournal?.confirm(operation, revision);
        runtime.remoteRevision = Math.max(runtime.remoteRevision, revision);
        runtime.collaboration?.acknowledge(operation.operationId, revision);
        await runtime.checkpointWorkspace(false);
        serverCheckpoint(runtime).request(revision);
        runtime.handlers.onSaveState?.(serverCheckpoint(runtime).hasFailure ? 'error' : runtime.collaboration?.getPendingOperations().length ? 'saving' : 'saved');
        return revision;
      } catch (error) {
        runtime.ownOperationIds.delete(operation.operationId);
        runtime.collaboration?.reject(operation.operationId, error instanceof Error ? error : new Error(String(error)));
        runtime.remoteConnected = false;
        runtime.handlers.onPhaseChange?.('error');
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

    let synchronizing = false;
    let synchronizationFailed = false;
    const deferredMessages: OperationMessage[] = [];
    const applyRemote = (message: OperationMessage) => {
      if (synchronizing) { deferredMessages.push(message); return; }
      if (runtime.disposed) return;
      if (message.type === 'revision.created') {
        if (message.payload.unitId !== runtime.model.unitId || message.revision <= runtime.remoteRevision) return;
        if (runtime.ownOperationIds.has(message.payload.operationId)) return;
        try { runtime.collaboration?.applyRemote(message.payload); }
        catch (error) {
          synchronizationFailed = true;
          runtime.remoteConnected = false;
          runtime.collaboration?.offlineQueue.setOnline(false);
          runtime.handlers.onSaveState?.('conflict');
          runtime.handlers.onNotice?.(error instanceof Error ? error.message : '远端变更与本地草稿冲突');
          return;
        }
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
      runtime.remoteConnected = false;
      runtime.collaboration?.offlineQueue.setOnline(false);
      runtime.handlers.onCollabStatus?.(status);
      if (status === 'closed') {
        runtime.collaboration?.transportClosed();
        runtime.handlers.onSaveState?.('offline');
      } else runtime.handlers.onSaveState?.('syncing');
      if (status !== 'open') return;
      synchronizing = true;
      synchronizationFailed = false;
      client.send({ type: 'cursor.updated', unitId: runtime.model.unitId, state: { sheetId: runtime.model.primarySheetId, row: 0, column: 0 } });
      void (async () => {
        const [snapshot, access] = await Promise.all([runtime.api.getSnapshot(runtime.model.unitId), runtime.api.getAccess(runtime.model.unitId)]);
        if (!active || runtime.disposed) return;
        runtime.handlers.onAccessRole?.(access.role);
        hydrateRuntime(runtime, snapshot);
        runtime.collaboration?.setRevision(snapshot.revision);
        await loadHistoryAndReplayPending(runtime);
        if (!active || runtime.disposed) return;
        synchronizing = false;
        for (const message of deferredMessages.splice(0)) applyRemote(message);
        if (synchronizationFailed) return;
        runtime.remoteConnected = true;
        runtime.collaboration?.offlineQueue.setOnline(true);
        runtime.handlers.onMutationsApplied?.();
        runtime.handlers.onSaveState?.(runtime.collaboration?.getPendingOperations().length ? 'saving' : 'saved');
      })().catch((error: Error) => {
        synchronizing = false;
        runtime.remoteConnected = false;
        runtime.handlers.onSaveState?.('conflict');
        runtime.handlers.onNotice?.(error.message);
      });
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
  runtime.remoteDataAvailable = false;
  serverCheckpoints.get(runtime)?.dispose();
  serverCheckpoints.delete(runtime);
  runtime.recoveryJournal?.release();
  runtime.collabDispose?.();
  runtime.bootstrapDispose?.();
  detachCoreListeners(runtime);
  runtime.formula.disposeCalculationTasks();
  for (const subscription of runtime.dataContentSubscriptions.values()) subscription.unsubscribe();
  runtime.dataContentSubscriptions.clear();
  runtime.dataContent.clear();
  runtime.collaboration?.attachTransport(undefined);
  runtime.collaboration = null;
  runtime.collab = null;
  cancelLocalSnapshotCheckpoint(runtime);
}

async function initializePersistence(runtime: SpreadsheetRuntime, isActive: () => boolean): Promise<void> {
  const resolution = runtime.resolution;
  if (runtime.recoveryJournal) {
    const pending = await runtime.recoveryJournal.load();
    runtime.operationJournal.write(runtime.model.unitId, pending, Math.max(0, ...pending.map(operation => operation.clientSequence)));
  }
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
      runtime.remoteSyncRequested = resolution.mode === 'offline';
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
    adoptLocalSnapshotBaseline(runtime, localRecord);
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
      const pendingLocalOperations = runtime.collaboration?.getPendingOperations() ?? [];
      const journalBaseRevision = runtime.operationJournal.read(runtime.model.unitId)?.snapshotRevision ?? runtime.localRevision;
      if (journalBaseRevision < runtime.localRevision) replayPendingOperations(runtime, pendingLocalOperations);
      else if (localPendingBeforeLoad.length > 0) replayPendingOperations(runtime, localPendingBeforeLoad);
      runtime.handlers.onNotice?.('Workbook restored from the current memory session');
    }
  }

  if (runtime.localOnly) {
    runtime.remoteDataAvailable = false;
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
      runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
      // Publish ready only after the session observes the authoritative
      // primary sheet.  Otherwise a user can click another tab between these
      // callbacks and have that selection overwritten by the late bootstrap
      // callback, leaving its lazy derived state unrequested.
      runtime.handlers.onPhaseChange?.('ready');
    }
    return;
  }

  try {
    const snapshotResponse = resolvedRemote ?? await runtime.api.getSnapshot(runtime.model.unitId);
    const access = resolution?.mode === 'remote' ? resolution.access : await runtime.api.getAccess(runtime.model.unitId);
    if (!access) throw new Error('Remote workbook resolution is missing access metadata');
    if (!isActive()) return;
    runtime.handlers.onAccessRole?.(access.role);
    hydrateRuntime(runtime, { ...snapshotResponse, snapshot: await migrateLegacyImageAssets(snapshotResponse.snapshot, runtime.assetStore) });
    runtime.remoteDataAvailable = true;
    runtime.remoteRevision = snapshotResponse.revision;
    runtime.localOnly = false;
    runtime.remoteSyncRequested = true;
    replaceCollaborationSession(runtime, localRecord);
    await loadHistoryAndReplayPending(runtime);
    if (!isActive()) return;
      runtime.remoteConnected = false;
      if (isActive()) {
        runtime.handlers.onSaveState?.('saved');
        runtime.handlers.onNotice?.('Workbook restored from server');
        runtime.handlers.onActiveSheetChange?.(runtime.model.primarySheetId);
        // Keep the visible ready boundary after the authoritative active-sheet
        // callback so immediate tab interaction cannot be reverted by startup.
        runtime.handlers.onPhaseChange?.('ready');
        runtime.handlers.onWorkspacePersisted?.();
      }
  } catch (error) {
    // Authentication, authorization and an unknown shared workbook are
    // authoritative server decisions. They must never be disguised as a new
    // local workbook with the same URL. Only an actual unavailable service
    // leaves the user in offline local mode.
    runtime.remoteDataAvailable = false;
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

    runtime.remoteConnected = false;
    runtime.handlers.onAccessRole?.(null);
    runtime.handlers.onSaveState?.('error');
    runtime.handlers.onPhaseChange?.('error');
    throw error;
  }
}

async function checkpointStartupLocally(runtime: SpreadsheetRuntime): Promise<void> {
  const offlineResolution = runtime.resolution?.mode === 'offline' ? runtime.resolution : null;
  runtime.workspaceRecord = await runtime.workspacePersistence.checkpoint(
    runtime.model.snapshot(),
    runtime.localRevision,
    runtime.remoteRevision,
    offlineResolution?.binding.syncMode ?? 'local-only',
    undefined,
    offlineResolution ? {
      location: offlineResolution.binding.location,
      lifecycle: offlineResolution.lifecycle,
      source: runtime.workspaceRecord?.metadata.source ?? 'native',
      role: runtime.workspaceRecord?.metadata.role ?? 'viewer',
    } : undefined,
  );
  adoptLocalSnapshotBaseline(runtime, runtime.workspaceRecord);
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
