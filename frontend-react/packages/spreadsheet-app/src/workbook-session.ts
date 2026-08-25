import type {
  CellData,
  CellEditorConfig,
  CellStyle,
  CellStyleTemplate,
  BarcodeSymbology,
  CameraDrawingPayload,
  ChartDrawingPayload,
  ConditionalFormatRule,
  DataBlockRef,
  DataSourceManifest,
  DefinedNameModel,
  DataValidationRule,
  FilterCriterion,
  DrawingObject,
  DrawingTransform,
  DataChartDrawingPayload,
  DataChartPlotType,
  FormControlDrawingPayload,
  FormControlType,
  ImageDrawingPayload,
  PivotFieldDefinition,
  PivotLayout,
  PivotMemberKey,
  PivotModel,
  PivotSourceRowPath,
  PivotAggregateFunction,
  PivotDefinition,
  RangeRef,
  ShapeDrawingPayload,
  SheetTableModel,
  SheetKind,
  SheetSnapshot,
  SheetDataRegion,
  SparklineModel,
  SparklineGroup,
  TableSheetDefinition,
  WorkbookTableModel,
  WorkbookSnapshot,
  WorksheetModel,
} from '@react-sheets/core-model';
import type { HistoryEntry, MutationInfo, CommandDescriptor, CommandResult } from '@react-sheets/command-runtime';
import type { AuthTokenProvider, GuestShareRole, RevisionRecord, ServerQueryRequest, ShareTokenProvider } from '@react-sheets/protocol';
import type { WorkbookApiClient } from '@react-sheets/protocol';
import type { NativePackageState } from '@react-sheets/exchange-excel-ooxml';
import { computePivotResult, computePivotResultFromBlockSource, getPivotFieldCatalog as buildPivotFieldCatalog, normalizePivotDefinition } from './features/pivot/engine';
import {
  copyRangeToClipboardData,
  planSheetTableCreation,
  findSheetTableAt,
  findValidationRule,
  groupsWithinRange,
  buildRowOutlineGroup,
  buildColumnOutlineGroup,
  normalizeRangeRef,
  parseTsv,
  resolveActiveAutoFilter,
  resolveFilterOwner,
  validationList,
  type ClipboardPayload,
  type PasteSpecialSpec,
  createPasteSpecialSpec,
  type DataRegionMaterializeParams,
  type GoToSpecialKind,
} from '@react-sheets/sheet-features';
import { isSpillChild, type RecalculationMode } from '@react-sheets/formula-engine';
import { EditSession } from './edit-session';
import {
  buildCollaborationSnapshot,
  type CollaborationSnapshot,
} from './collaboration';
import { PermissionService, type PermissionCapabilities, type ShareRole } from './permission-service';
import {
  canExecuteCommand,
  findProtectionRuleCoveringRange,
} from './features/permission';
import {
  createSpreadsheetRuntime,
  disposeSpreadsheetRuntime,
  hydrateRuntime,
  rehydrateFormulaAfterRestore,
  resolveActorId,
  resolveShareToken,
  resolveUnitId,
  scheduleFormulaRecalculation,
  startCollaborationSession,
  startPersistenceSession,
  type SpreadsheetRuntime,
} from './runtime';
import { createInitialSelection, SelectionService, parseRangeReference, type SelectionState } from './selection-service';
import { cellAddress, columnLabel } from './address';
import { writeSystemClipboard, type SystemClipboardWriteOutcome } from './clipboard-browser';
import { buildCanvasSheetSnapshot, type CanvasSheetSnapshot } from './ui-snapshot';
import {
  buildDrawingAdd,
  findDrawingByPayloadId,
  resolveDrawingMoveTransform,
} from './features/drawing';
import {
  buildChartInsertParams,
} from './features/drawing';
import {
  buildPivotModel,
  connectedPivotIdsForSource,
  readPivotBlockSource,
} from './features/pivot';
import {
  buildPivotSlicerDrawing,
  buildPivotTimelineDrawing,
  listPivotControlsForPivot,
  type PivotControlRecord,
} from './features/pivot-controls';
import {
  prepareDataRegionMaterialization,
  createWorkbookCellResolver,
} from './features/data-source';
import type { TableRowsResponse, WorkbookCellResolver } from './features/data-source';
import {
  buildCellNote,
  buildCommentReply,
  buildCommentThread,
  findCommentThreadAt,
  getCellHyperlink,
  parseUrlHyperlink,
} from './features/review';
import {
  buildSparklineDataLocationParams,
  buildSparklineGroup,
  buildSparklineInsertParams,
  resolveQuickSparklinePlacement,
} from './features/sparkline';
import {
  buildRestoreParams,
  revisionToHistoryMeta,
} from './features/history';
import {
  buildPersistenceMeta,
  type WorkspacePersistence,
  type PersistenceSnapshotMeta,
} from './features/persistence';
import type { FormulaAuditProjection } from './features/formula-audit';
import { exchangeExportXlsx, summarizeCompatibilityReport } from './features/xlsx';
import {
  buildPrintSnapshot,
  getPrintDocument,
  pageSetupToPrintLayout,
  summarizePrintSnapshot,
  type PageSetup,
  type PrintPageBreak,
  type PrintPageSnapshot,
  type PrintSnapshot,
} from './features/print';
import { browserPrintHook, PdfExportService, type PrintLayout } from './features/print';
import type { LoadTarget, QueryDefinition } from './features/query/query-steps';
import {
  buildQueryResultSnapshot,
  deserializeQueryDefinition,
  executeQueryDefinition,
  resolveLoadTarget,
  summarizeQueryResult,
  type QueryResultSnapshot,
  type QuerySessionEntry,
} from './features/query';
import {
  createCommandRecorder,
  runAutomationScriptAsync,
  SAMPLE_AUTOMATION_SCRIPT,
  summarizeScriptResult,
  type AutomationSnapshot,
} from './features/automation';
import type { AutomationWorkerFactory } from './features/automation';
import { CommandRecorder } from './features/automation/command-recorder';
import type { ScriptRunResult } from './features/automation';
import type {
  DataTableParams,
  DataTableResult,
  GoalSeekParams,
  GoalSeekResult,
  ScenarioDefinition,
  ScenarioResult,
} from './features/extended/what-if';
import type { WhatIfPlan } from './features/extended';
import type { CompatibilityReport } from './features/xlsx';
import {
  HistoryPreviewSession,
  type HistoryEntryMeta,
} from './features/history';
import { currentRegionOfSheet, inferTableFieldType, nextId, usedRangeOfSheet } from './application-helpers';
import type {
  ActiveContext,
  AppPhase,
  BackstagePanel,
  BackstageState,
  ClipboardState,
  CellShiftOperation,
  DesignerState,
  EditSession as DesignerEditSession,
  FocusState,
  HomeRibbonState,
  HomeStyleKey,
  InputMode,
  PanelState,
  PeerCursor,
  RibbonTabId,
  SaveState,
  SidebarPanelId,
  SheetDialogState,
  UiSessionIntent,
  DialogState,
  UndoRedoState,
} from './types';

export interface WorkbookSessionOptions {
  unitId?: string;
  api?: WorkbookApiClient;
  workspacePersistence?: WorkspacePersistence;
  initialPhase?: AppPhase;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  automationWorkerFactory?: AutomationWorkerFactory;
  /** Only Node/unit harnesses may opt into the inline exchange implementation. */
  xlsxExecution?: 'worker' | 'inline-test';
}

export type DispatchErrorCode = 'WORKBOOK_NOT_READY' | 'COMMAND_REJECTED' | 'MATERIALIZATION_FAILED';

export class CommandDispatchError extends Error {
  constructor(
    readonly code: DispatchErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'CommandDispatchError';
  }
}

export type DispatchOutcome =
  | { status: 'committed'; result: CommandResult }
  | { status: 'rejected'; error: CommandDispatchError };

export type ClipboardExecutionOutcome = SystemClipboardWriteOutcome & {
  privatePayloadStored: boolean;
};

export interface UiSnapshot extends DesignerState {
  unitId: string;
  workbookName: string;
  phase: AppPhase;
  saveState: SaveState;
  notice: string;
  selection: SelectionState;
  homeRibbon: HomeRibbonState;
  activeCell: string;
  activeSheetId: string;
  panels: PanelState;
  ribbonTab: RibbonTabId;
  formulaDraft: string;
  editingCell: { row: number; column: number } | null;
  zoom: number;
  sheets: CanvasSheetSnapshot[];
  selectedSheet: CanvasSheetSnapshot;
  selectedFloatingId: string | null;
  selectedDrawingIds: readonly string[];
  drawingSelectionMode: boolean;
  activeContext: ActiveContext;
  peers: PeerCursor[];
  collabStatus: 'connecting' | 'open' | 'closed';
  collabRevision: number;
  pendingChangeSetCount: number;
  pendingCommandCount: number;
  offlineQueueState: string;
  actorId: string;
  shareRole: ShareRole | null;
  permissions: PermissionCapabilities;
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  historyPreviewRevision: number | null;
  hasPendingOperations: boolean;
  persistenceChecksum: string;
  compatibilityReport: CompatibilityReport | null;
  tables: readonly WorkbookTableModel[];
  relationships: readonly import('@react-sheets/core-model').DataRelationship[];
  dataSources: readonly DataSourceManifest[];
  definedNameModels: readonly DefinedNameModel[];
  cellStyleTemplates: readonly CellStyleTemplate[];
  dialogs: DialogState;
  inputMode: InputMode;
  focus: FocusState;
  formatPainter: 'once' | 'locked' | null;
  printLayout: PrintLayout;
  printPages: readonly PrintPageSnapshot[];
  printPageCount: number;
  printArea: RangeRef | null;
  lastQueryResult: QueryResultSnapshot | null;
  queryConnectors: readonly string[];
  loadedQueries: readonly QueryResultSnapshot[];
  automationRecording: boolean;
  recordedScript: string;
  lastScriptResult: ScriptRunResult | null;
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
  formulaAudit: FormulaAuditProjection;
  version: number;
  /** Canonical DesignerState fields; render-specific projections remain read-only views of the same session. */
  workbook: DesignerState['workbook'];
  editSession: DesignerEditSession | null;
  activeObject: DesignerState['activeObject'];
  ribbon: DesignerState['ribbon'];
  clipboard: ClipboardState;
  undoRedo: UndoRedoState;
}

const HOME_STYLE_KEYS: readonly HomeStyleKey[] = [
  'fontFamily',
  'fontSizePx',
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'textColor',
  'background',
  'horizontalAlignment',
  'verticalAlignment',
  'indent',
  'wrapText',
  'numberFormat',
  'borders',
];
const HOME_STYLE_SCAN_LIMIT = 10_000;

function sameRange(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startColumn === right.startColumn
    && left.endColumn === right.endColumn;
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn
    && left.endColumn >= right.startColumn;
}

function compareStyleValue(value: unknown): string {
  return value === undefined ? '__unset__' : JSON.stringify(value);
}

export interface ExtendedSnapshot {
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
}

export interface CreatePivotTableParams {
  source?: PivotDefinition['source'];
  destination: { kind: 'new-sheet' } | { kind: 'existing-sheet'; sheetId: string; anchor: { row: number; column: number } };
}

export interface DefinedNameCommandInput {
  name: string;
  formula: string;
  scope?: DefinedNameModel['scope'];
  sheetId?: string;
  hidden?: boolean;
  comment?: string;
}

export function getInitialSessionPhase(): AppPhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'error' || queryPhase === 'empty' ? queryPhase : 'loading';
}

export class WorkbookSession {
  private readonly runtime: SpreadsheetRuntime;
  private readonly cellResolver: WorkbookCellResolver;
  private readonly permission: PermissionService;
  private readonly editSession = new EditSession();
  private readonly listeners = new Set<() => void>();
  private readonly actorId: string;
  private readonly automationWorkerFactory?: AutomationWorkerFactory;
  private readonly xlsxExecution: 'worker' | 'inline-test';

  private phase: AppPhase;
  private saveState: SaveState = 'saved';
  private notice = 'Workbook engine ready';
  private version = 0;
  private activeSheetId: string;
  private panels: PanelState = { active: 'inspector', open: false, dock: 'right' };
  private barcodeDraftSymbology: BarcodeSymbology = 'qr';
  private backstage: BackstageState = { open: false, panel: 'info' };
  private ribbonTab: RibbonTabId = 'home';
  private formulaDraft = '';
  private zoom = 100;
  /** DrawingRuntime is the sole owner of transient object selection. */
  private get selectedDrawingIds(): readonly string[] {
    return this.runtime.drawing.getSelection(this.activeSheetId);
  }
  /** Canvas needs a single focus rectangle; derive it from canonical multi-select. */
  private get selectedFloatingId(): string | null {
    return this.selectedDrawingIds[0] ?? null;
  }
  private drawingSelectionMode = false;
  private activeContext: ActiveContext = { kind: 'none' };
  private peers: PeerCursor[] = [];
  private collabStatus: 'connecting' | 'open' | 'closed' = 'closed';
  private remoteRevisions: RevisionRecord[] = [];
  private historyPreview: HistoryPreviewSession | null = null;
  private hasPendingOperations = false;
  private persistenceChecksum = '';
  private compatibilityReport: CompatibilityReport | null = null;
  /** The sole native package baseline paired with this workbook snapshot. */
  private nativePackage: NativePackageState | undefined;
  private dialogs: DialogState = { active: null, findQuery: '', mergeDiscardCount: 0, columnWidth: null, sheet: null, cellShiftOperation: 'insert' };
  private pendingMergeRange: RangeRef | null = null;
  private inputMode: InputMode = 'grid';
  private focus: FocusState = { mode: 'grid', target: 'grid' };
  private formatPainter: { mode: 'once' | 'locked'; sourceRange: RangeRef } | null = null;
  private printLayout: PrintLayout = {
    paper: 'A4',
    orientation: 'portrait',
    margin: { top: 20, right: 20, bottom: 20, left: 20 },
  };
  private printSnapshot: PrintSnapshot | null = null;
  private querySessions = new Map<string, QuerySessionEntry>();
  private lastQueryResult: QueryResultSnapshot | null = null;
  private readonly commandRecorder: CommandRecorder = createCommandRecorder();
  private recorderDetach: (() => void) | null = null;
  private automationRecording = false;
  private recordedScript = '';
  private lastScriptResult: ScriptRunResult | null = null;
  private lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null = null;
  private lastRepeatableCommand: CommandDescriptor | null = null;
  private readonly pivotTaskGeneration = new Map<string, number>();

  private selectionService: SelectionService;
  private collabDispose: (() => void) | null = null;
  private persistenceDispose: (() => void) | null = null;
  private started = false;
  private disposed = false;
  private lifecycleGeneration = 0;
  private overrideTarget: { row: number; column: number } | null = null;
  private clipboardData: ClipboardPayload | null = null;
  private clipboardSystemStatus: 'unknown' | 'published' | 'reduced' | 'failed' = 'unknown';
  private clipboardSystemFormats: readonly string[] = [];
  private readonly materializingDataRegions = new Map<string, Promise<void>>();
  private pendingCommandCount = 0;
  private snapshotGeneration = 0;
  private cachedUiSnapshot: UiSnapshot | null = null;
  private cachedUiSnapshotGeneration = -1;
  private projectionGeneration = 0;
  private readonly sheetProjectionCache = new Map<string, { generation: number; snapshot: CanvasSheetSnapshot }>();
  private persistenceMetaDirty = true;

  constructor({ unitId, api, workspacePersistence, initialPhase = 'ready', authTokenProvider, shareTokenProvider, automationWorkerFactory, xlsxExecution = 'worker' }: WorkbookSessionOptions = {}) {
    const routeShareToken = shareTokenProvider ? null : resolveShareToken();
    this.runtime = createSpreadsheetRuntime({
      unitId,
      api,
      workspacePersistence,
      authTokenProvider,
      shareTokenProvider: shareTokenProvider ?? (routeShareToken ? () => routeShareToken : undefined),
    });
    this.cellResolver = createWorkbookCellResolver(this.runtime.dataContent);
    this.permission = new PermissionService();
    this.automationWorkerFactory = automationWorkerFactory;
    this.xlsxExecution = xlsxExecution;
    this.permission.setOnline(!this.runtime.localOnly);
    this.actorId = resolveActorId();
    this.phase = initialPhase;
    this.activeSheetId = this.runtime.model.primarySheetId;
    this.selectionService = new SelectionService(
      this.runtime.model.unitId,
      () => this.activeSheetId,
      () => {
        const sheet = this.runtime.model.getSheet(this.activeSheetId);
        return { rowCount: sheet.rowCount, columnCount: sheet.columnCount };
      },
      createInitialSelection(this.activeSheetId),
    );
    this.wireRuntimeHandlers();
    this.syncPersistenceMeta();
  }

  private ensureActiveSheetSession(): void {
    if (this.runtime.model.sheets.has(this.activeSheetId)) return;
    this.activeSheetId = this.runtime.model.primarySheetId;
    this.selectionService.resetForSheet(this.activeSheetId);
    this.editSession.cancel();
    this.formulaDraft = '';
  }

  /**
   * Reconcile every transient drawing projection after a canonical workbook
   * change. Worksheet drawing collections are the sole existence authority;
   * selection, active context, and pointer gestures must never outlive them.
   */
  private reconcileDrawingSessionState(): void {
    const sheets = this.runtime.model.getSheets();
    const liveSheetIds = sheets.map((sheet) => sheet.id);
    for (const sheet of sheets) {
      this.runtime.drawing.reconcile(sheet.id, sheet.drawings.map((drawing) => drawing.id));
    }
    this.runtime.drawing.clearMissingSheets(liveSheetIds);

    const previousContext = this.activeContext;
    const activeSelection = this.runtime.drawing.getSelection(this.activeSheetId);
    let contextRemoved = false;
    let nextContext = this.activeContext;
    const context = this.activeContext;
    if (context.kind === 'drawing') {
      const contextSheet = sheets.find((sheet) => sheet.id === context.sheetId);
      const contextExists = contextSheet?.drawings.some((drawing) => drawing.id === context.drawingId) ?? false;
      if (!contextExists) {
        contextRemoved = true;
        const fallbackId = context.sheetId === this.activeSheetId ? activeSelection[0] : undefined;
        nextContext = fallbackId
          ? { kind: 'drawing', sheetId: this.activeSheetId, drawingId: fallbackId }
          : { kind: 'none' };
      }
    }

    if (activeSelection.length === 0 && this.drawingSelectionMode) {
      this.drawingSelectionMode = false;
      this.inputMode = 'grid';
      this.focus = { mode: 'grid', target: 'grid' };
    }
    if (contextRemoved && (this.panels.active === 'chart' || this.panels.active === 'dataChart' || this.panels.active === 'barcode' || this.panels.active === 'shape' || this.panels.active === 'picture')) {
      this.panels = { ...this.panels, open: false };
    }
    if (contextRemoved && this.ribbonTab === 'pictureFormat') this.ribbonTab = 'home';
    if (JSON.stringify(nextContext) !== JSON.stringify(previousContext)) {
      this.activeContext = nextContext;
    }
  }

  private syncTableContextFromSelection(): void {
    if (this.drawingSelectionMode) return;
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    if (sheet.kind !== 'worksheet') return;
    const active = this.selectionService.getState().activeCell;
    const sparkline = sheet.sparklines.find((entry) => entry.anchor.row === active.row && entry.anchor.column === active.column);
    const table = findSheetTableAt(sheet, active.row, active.column);
    const next: ActiveContext = sparkline
      ? { kind: 'sparkline', sheetId: sheet.id, sparklineId: sparkline.id }
      : table
      ? { kind: 'table', sheetId: sheet.id, tableId: table.id }
      : { kind: 'none' };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    if (this.activeContext.kind !== 'none' && !['table', 'sparkline'].includes(this.activeContext.kind)) return;
    this.activeContext = next;
    if (next.kind === 'sparkline') {
      this.panels = { ...this.panels, active: 'sparkline', open: true };
      this.ribbonTab = 'sparklineDesign';
    } else if (next.kind === 'table') {
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'tableDesign';
    } else if (this.ribbonTab === 'tableDesign' || this.ribbonTab === 'sparklineDesign') {
      this.ribbonTab = 'home';
    }
  }

  private wireRuntimeHandlers(): void {
    this.runtime.handlers.onSaveState = (state) => {
      this.saveState = state;
      this.emit();
    };
    this.runtime.handlers.onNotice = (message) => {
      this.notice = message;
      this.emit();
    };
    this.runtime.handlers.onMutationsApplied = () => {
      this.projectionGeneration += 1;
      this.persistenceMetaDirty = true;
      this.restorePersistedQuerySessions();
      this.ensureActiveSheetSession();
      this.reconcileDrawingSessionState();
      this.syncTableContextFromSelection();
      this.runtime.formulaAudit.refresh();
      this.refresh();
    };
    this.runtime.handlers.onPhaseChange = (phase) => {
      this.phase = phase;
      this.emit();
    };
    this.runtime.handlers.onActiveSheetChange = (sheetId) => {
      this.activeSheetId = sheetId;
      this.reconcileDrawingSessionState();
      this.emit();
    };
    this.runtime.handlers.onRemoteRevisions = (revisions) => {
      this.remoteRevisions = revisions;
      this.emit();
    };
    this.runtime.handlers.onWorkspacePersisted = () => {
      this.persistenceMetaDirty = true;
      this.syncPersistenceMeta();
    };
    this.runtime.handlers.onCollabStatus = (status) => {
      this.collabStatus = status;
      if (status === 'closed') this.permission.setOnline(false);
      else if (this.permission.getShareRole()) this.permission.setOnline(true);
      this.emit();
    };
    this.runtime.handlers.onAccessRole = (role) => {
      if (role) {
        this.permission.applyServerAccess(role);
        this.permission.setOnline(true);
      } else {
        this.permission.clearServerAccess();
        this.permission.setOnline(false);
      }
      this.emit();
    };
    this.runtime.handlers.onPeersChange = (peer) => {
      if (peer.length === 0) {
        this.peers = [];
      } else {
        const incoming = peer[0]!;
        this.peers = [...this.peers.filter((p) => p.actorId !== incoming.actorId), incoming];
      }
      this.emit();
    };
  }

  start(): void {
    if (this.started) return;
    this.disposed = false;
    this.started = true;
    const generation = ++this.lifecycleGeneration;
    this.persistenceDispose = startPersistenceSession(this.runtime);
    void this.runtime.persistenceReady.then(async () => {
      if (this.disposed || generation !== this.lifecycleGeneration) return;
      const artifact = await this.runtime.workspacePersistence.nativePackages.load(this.runtime.model.unitId);
      if (!this.disposed && generation === this.lifecycleGeneration && artifact) {
        this.nativePackage = artifact;
        this.projectionGeneration += 1;
        this.emit();
      }
      if (!this.disposed && generation === this.lifecycleGeneration) this.restorePersistedQuerySessions();
      if (!this.disposed && generation === this.lifecycleGeneration) {
        this.collabDispose = startCollaborationSession(this.runtime, () =>
          `${this.activeSheetId}:${this.selectionService.getState().activeCell.row}:${this.selectionService.getState().activeCell.column}`,
          this.runtime.authTokenProvider,
          this.runtime.shareTokenProvider,
        );
      }
    }).catch((error: unknown) => {
      if (!this.disposed && generation === this.lifecycleGeneration) this.notify(error instanceof Error ? error.message : 'Workbook persistence initialization failed');
    });
  }

  dispose(): void {
    if (!this.started && this.disposed) return;
    this.disposed = true;
    this.started = false;
    this.lifecycleGeneration += 1;
    this.recorderDetach?.();
    this.recorderDetach = null;
    this.collabDispose?.();
    this.persistenceDispose?.();
    this.collabDispose = null;
    this.persistenceDispose = null;
    disposeSpreadsheetRuntime(this.runtime);
    this.sheetProjectionCache.clear();
    this.cachedUiSnapshot = null;
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  private emit(): void {
    if (this.disposed) return;
    this.snapshotGeneration += 1;
    for (const listener of this.listeners) listener();
  }

  private refresh(): void {
    this.syncPersistenceMeta();
    this.version += 1;
    this.emit();
  }

  private syncPersistenceMeta(): void {
    if (!this.persistenceMetaDirty) return;
    const meta = buildPersistenceMeta(
      this.runtime.model.snapshot(),
      this.runtime.remoteRevision,
      this.runtime.collaboration?.offlineQueue.getPendingCount() ?? 0,
      this.runtime.workspaceRecord,
    );
    this.hasPendingOperations = meta.hasPendingOperations;
    this.persistenceChecksum = meta.checksum;
    this.persistenceMetaDirty = false;
  }

  private restorePersistedQuerySessions(): void {
    const persisted = this.runtime.model.listQueryDefinitions();
    const current = new Set(persisted.map((definition) => definition.id));
    for (const queryId of this.querySessions.keys()) {
      if (!current.has(queryId)) this.querySessions.delete(queryId);
    }
    for (const definition of persisted) {
      if (!this.querySessions.has(definition.id)) {
        try {
          this.querySessions.set(definition.id, { definition: deserializeQueryDefinition(definition) });
        } catch (error) {
          // Keep the canonical persisted definition in the workbook so the
          // user can repair it; do not expose a falsely loaded query result.
          this.notify(error instanceof Error ? `Query ${definition.id} could not be restored: ${error.message}` : `Query ${definition.id} could not be restored`);
        }
      }
    }
  }

  /** The sole worksheet read path for session-level Home behavior. */
  private readResolvedCell(sheet: WorksheetModel, row: number, column: number): CellData | undefined {
    return this.cellResolver.resolve(sheet, row, column)?.cell;
  }

  /**
   * Derives the one Home-ribbon view of the current selection. This is kept
   * beside command dispatch so buttons, shortcuts and context menus all make
   * permission decisions against the same ranges.
   */
  private deriveHomeRibbonState(selection: SelectionState): HomeRibbonState {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const primary = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0]
      ?? normalizeRangeRef({ sheetId: this.activeSheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
    const seen = new Map<HomeStyleKey, { value: unknown; signature: string; mixed: boolean }>();
    let scanned = 0;
    let truncated = false;

    outer: for (const selectedRange of selection.ranges) {
      const range = normalizeRangeRef({ ...selectedRange, sheetId: this.activeSheetId });
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const cell = this.readResolvedCell(sheet, row, column);
          const style: Partial<CellStyle> = {
            ...(cell?.style ?? {}),
            ...(cell?.numberFormat === undefined ? {} : { numberFormat: cell.numberFormat }),
          };
          for (const key of HOME_STYLE_KEYS) {
            const value = style[key];
            const signature = compareStyleValue(value);
            const prior = seen.get(key);
            if (!prior) {
              seen.set(key, { value, signature, mixed: false });
            } else if (prior.signature !== signature) {
              prior.mixed = true;
            }
          }
          scanned += 1;
          if (scanned >= HOME_STYLE_SCAN_LIMIT) {
            truncated = true;
            break outer;
          }
        }
      }
    }

    const style: Partial<CellStyle> = {};
    const mixedStyleKeys: HomeStyleKey[] = [];
    for (const key of HOME_STYLE_KEYS) {
      const summary = seen.get(key);
      if (truncated || summary?.mixed) {
        mixedStyleKeys.push(key);
      } else if (summary?.value !== undefined) {
        style[key] = summary.value as never;
      }
    }

    const exactMerge = sheet.merges.some((merge) => sameRange(merge.range, primary));
    const intersectsMerge = sheet.merges.some((merge) => selection.ranges.some((range) => rangesIntersect(range, merge.range)));
    const activeAutoFilter = resolveActiveAutoFilter(sheet);
    return {
      sheetId: this.activeSheetId,
      ranges: selection.ranges.map((range) => structuredClone(range)),
      activeCell: { ...selection.activeCell },
      style,
      mixedStyleKeys,
      merge: exactMerge ? 'full' : intersectsMerge ? 'mixed' : 'none',
      canFormat: this.canExecute('sheet.style.set', { style: {} }),
      canEdit: this.canExecute('sheet.range.clear', { mode: 'contents' }),
      canStructure: this.canExecute('sheet.rows.insert', { count: 1 }),
      hasFilter: Boolean(activeAutoFilter),
      hasFilterCriteria: Object.values(activeAutoFilter?.columns ?? {}).some((column) => Boolean(column.criterion)),
    };
  }

  getUiSnapshot = (): UiSnapshot => {
    if (this.cachedUiSnapshot && this.cachedUiSnapshotGeneration === this.snapshotGeneration) {
      return this.cachedUiSnapshot;
    }
    const activeSheetIds = new Set(this.runtime.model.getSheets().map((sheet) => sheet.id));
    for (const sheetId of this.sheetProjectionCache.keys()) {
      if (!activeSheetIds.has(sheetId)) this.sheetProjectionCache.delete(sheetId);
    }
    const sheets = this.runtime.model.getSheets().map((sheet) => {
      const cached = this.sheetProjectionCache.get(sheet.id);
      if (cached?.generation === this.projectionGeneration) return cached.snapshot;
      const snapshot = buildCanvasSheetSnapshot(
        this.runtime.model,
        sheet,
        this.runtime.formula,
        true,
        this.runtime.pivotResults,
        this.runtime.dataContent,
        this.nativePackage?.dateSystem ?? '1900',
      );
      this.sheetProjectionCache.set(sheet.id, { generation: this.projectionGeneration, snapshot });
      return snapshot;
    });
    const selectedSheet = sheets.find((sheet) => sheet.id === this.activeSheetId) ?? sheets[0]!;
    const selection = this.selectionService.getState();
    const collaboration = this.getCollaborationSnapshot();
    const homeRibbon = this.deriveHomeRibbonState(selection);
    const undoEntries = this.runtime.commands.getUndoEntries();
    const redoEntries = this.runtime.commands.getRedoEntries();
    const activeEdit = this.editSession.active;
    const activeDrawing = this.selectedFloatingId
      ? selectedSheet.drawings.find((drawing) => drawing.id === this.selectedFloatingId)
      : undefined;
    const snapshot: UiSnapshot = {
      workbook: { unitId: this.runtime.model.unitId, name: this.runtime.model.name },
      unitId: this.runtime.model.unitId,
      workbookName: this.runtime.model.name,
      phase: this.phase,
      saveState: this.saveState,
      notice: this.notice,
      selection,
      homeRibbon,
      activeCell: this.selectionService.activeCell,
      activeSheetId: this.activeSheetId,
      panels: { ...this.panels },
      ribbonTab: this.ribbonTab,
      formulaDraft: this.formulaDraft,
      editingCell: this.editSession.editingCell,
      zoom: this.zoom,
      sheets,
      selectedSheet,
      selectedFloatingId: this.selectedFloatingId,
      selectedDrawingIds: [...this.selectedDrawingIds],
      drawingSelectionMode: this.drawingSelectionMode,
      activeContext: this.activeContext,
      peers: this.peers,
      collabStatus: this.collabStatus,
      collabRevision: collaboration.revision,
      pendingChangeSetCount: collaboration.pendingCount,
      pendingCommandCount: this.pendingCommandCount,
      offlineQueueState: collaboration.offlineQueueState,
      actorId: this.actorId,
      shareRole: this.getShareRole(),
      permissions: this.permission.getCapabilities(),
      historyEntries: this.runtime.commands.getUndoEntries(),
      remoteRevisions: this.remoteRevisions,
      historyPreviewRevision: this.historyPreview?.revision ?? null,
      hasPendingOperations: this.hasPendingOperations,
      persistenceChecksum: this.persistenceChecksum,
      compatibilityReport: this.compatibilityReport,
      tables: [...this.runtime.model.dataModel.tables.values()].map((table) => structuredClone(table)),
      relationships: [...this.runtime.model.dataModel.relationships.values()].map((relationship) => structuredClone(relationship)),
      dataSources: [...this.runtime.model.dataModel.sources.values()].map((source) => structuredClone(source)),
      definedNameModels: structuredClone(this.runtime.model.definedNameModels),
      cellStyleTemplates: this.runtime.model.listCellStyleTemplates(),
      dialogs: { ...this.dialogs },
      backstage: { ...this.backstage },
      inputMode: this.inputMode,
      focus: { ...this.focus },
      editSession: activeEdit ? {
        sheetId: activeEdit.sheetId,
        cell: { row: activeEdit.row, column: activeEdit.column },
        originalValue: activeEdit.originalValue?.value ?? null,
        ...(activeEdit.originalFormula !== undefined ? { originalFormula: activeEdit.originalFormula } : {}),
        draftText: activeEdit.currentDraft,
        mode: activeEdit.originalFormula || activeEdit.currentDraft.startsWith('=') ? 'formula' : 'value',
        source: this.focus.target === 'formula-bar' ? 'formulaBar' : 'cell',
      } : null,
      activeObject: activeDrawing ? { kind: activeDrawing.kind, id: activeDrawing.id } : null,
      ribbon: { activeTab: this.ribbonTab },
      clipboard: {
        hasContent: Boolean(this.clipboardData),
        mode: this.clipboardData ? (this.clipboardData.transfer === 'move' ? 'cut' : 'copy') : null,
        systemStatus: this.clipboardSystemStatus,
        systemFormats: this.clipboardSystemFormats,
      },
      undoRedo: { canUndo: undoEntries.length > 0, canRedo: redoEntries.length > 0, undoCount: undoEntries.length, redoCount: redoEntries.length },
      formatPainter: this.formatPainter?.mode ?? null,
      printLayout: this.printLayout,
      printPages: this.printSnapshot?.pageSnapshots ?? [],
      printPageCount: this.printSnapshot?.pageCount ?? 0,
      printArea: this.printSnapshot?.printArea ?? null,
      lastQueryResult: this.lastQueryResult,
      queryConnectors: this.runtime.connectors.list().map((connector) => connector.id),
      loadedQueries: [...this.querySessions.values()]
        .map((session) => session.lastResult)
        .filter((result): result is QueryResultSnapshot => Boolean(result)),
      automationRecording: this.automationRecording,
      recordedScript: this.recordedScript,
      lastScriptResult: this.lastScriptResult,
      lastWhatIfResult: this.lastWhatIfResult,
      formulaAudit: this.runtime.formulaAudit.getProjection(),
      version: this.version,
    };
    this.cachedUiSnapshotGeneration = this.snapshotGeneration;
    this.cachedUiSnapshot = snapshot;
    return snapshot;
  };

  /** Dispatch one registered domain descriptor through the sole command path. */
  dispatch(descriptor: CommandDescriptor): Promise<DispatchOutcome> {
    if (this.phase !== 'ready') {
      return this.rejectDispatch(
        new CommandDispatchError('WORKBOOK_NOT_READY', 'Workbook is not ready'),
      );
    }
    try {
      const resolved = this.resolveCommandContext(descriptor.commandId, descriptor.params);
      const regions = this.dataRegionsRequiredForCommand(descriptor.commandId, resolved);
      if (regions.length > 0) {
        return this.dispatchAfterMaterialization(descriptor.commandId, resolved, regions);
      }
      const result = this.runCommand(descriptor.commandId, resolved);
      return Promise.resolve({ status: 'committed', result });
    } catch (error) {
      return this.rejectDispatch(this.toDispatchError(error, 'COMMAND_REJECTED', 'Command was rejected'));
    }
  }

  /** Materialization is part of the dispatch transaction and is not observable as a scheduled terminal state. */
  private async dispatchAfterMaterialization(commandId: string, params: unknown, regions: readonly SheetDataRegion[]): Promise<DispatchOutcome> {
    try {
      await this.materializeDataRegions(regions);
    } catch (error) {
      return this.rejectDispatch(this.toDispatchError(error, 'MATERIALIZATION_FAILED', 'Data region could not be prepared for editing'));
    }
    try {
      const result = this.runCommand(commandId, params);
      return { status: 'committed', result };
    } catch (error) {
      return this.rejectDispatch(this.toDispatchError(error, 'COMMAND_REJECTED', 'Command was rejected'));
    }
  }

  private toDispatchError(error: unknown, code: DispatchErrorCode, fallback: string): CommandDispatchError {
    if (error instanceof CommandDispatchError) return error;
    return new CommandDispatchError(code, error instanceof Error ? error.message : fallback);
  }

  private rejectDispatch(error: CommandDispatchError): Promise<DispatchOutcome> {
    this.notify(error.message);
    return Promise.resolve({ status: 'rejected', error });
  }

  /**
   * Async command boundary for operations whose domain handler reads cells.
   * It is the only path used by data tools that may target block-backed
   * regions: materialization completes before the command transaction starts.
   */
  private async executeCommandAfterMaterialization(commandId: string, params: unknown): Promise<CommandResult> {
    if (this.phase !== 'ready') throw new Error('Workbook is not ready');
    const resolved = this.resolveCommandContext(commandId, params);
    const regions = this.dataRegionsRequiredForCommand(commandId, resolved);
    if (regions.length > 0) await this.materializeDataRegions(regions);
    return this.runCommand(commandId, resolved);
  }

  private async materializeDataRegions(regions: readonly SheetDataRegion[]): Promise<void> {
    for (const region of regions) {
      const key = `${region.range.sheetId}:${region.id}`;
      let pending = this.materializingDataRegions.get(key);
      if (!pending) {
        this.pendingCommandCount += 1;
        this.emit();
        pending = (async () => {
          const prepared = await prepareDataRegionMaterialization(
            this.runtime.model,
            region.range.sheetId,
            region.id,
            this.runtime.dataContent,
          );
          this.runCommand('dataRegion.materialize.commit', prepared satisfies DataRegionMaterializeParams);
        })();
        this.materializingDataRegions.set(key, pending);
      }
      try {
        await pending;
      } finally {
        if (this.materializingDataRegions.get(key) === pending) {
          this.materializingDataRegions.delete(key);
          this.pendingCommandCount = Math.max(0, this.pendingCommandCount - 1);
          this.emit();
        }
      }
    }
  }

  private dataRegionsIntersectingRanges(sheetId: string, ranges: readonly RangeRef[]): SheetDataRegion[] {
    const sheet = this.runtime.model.getSheet(sheetId);
    return sheet.dataRegions.filter((region) => ranges.some((range) => rangesIntersect(range, region.range)));
  }

  private dataRegionsRequiredForCommand(commandId: string, params: unknown): SheetDataRegion[] {
    if (!this.commandRequiresResolvedWrites(commandId) || commandId === 'dataRegion.materialize.commit') return [];
    const input = params && typeof params === 'object' && !Array.isArray(params) ? params as Record<string, unknown> : {};
    const sheetId = typeof input.sheetId === 'string' ? input.sheetId : this.activeSheetId;
    const sheet = this.runtime.model.getSheet(sheetId);
    const ranges: RangeRef[] = [];
    const appendRange = (candidate: unknown) => {
      if (!candidate || typeof candidate !== 'object') return;
      const range = candidate as RangeRef;
      if (typeof range.sheetId === 'string' && Number.isInteger(range.startRow) && Number.isInteger(range.endRow)
        && Number.isInteger(range.startColumn) && Number.isInteger(range.endColumn)) ranges.push(range);
    };
    appendRange(input.range);
    appendRange(input.sourceRange);
    appendRange(input.targetRange);
    if (input.targetOrigin && typeof input.targetOrigin === 'object' && !Array.isArray(input.targetOrigin)) {
      const origin = input.targetOrigin as { row?: unknown; column?: unknown };
      const clipboard = input.clipboard && typeof input.clipboard === 'object' && !Array.isArray(input.clipboard)
        ? input.clipboard as { values?: unknown[][]; range?: unknown }
        : undefined;
      appendRange(clipboard?.range);
      const originRow = typeof origin.row === 'number' && Number.isInteger(origin.row) ? origin.row : undefined;
      const originColumn = typeof origin.column === 'number' && Number.isInteger(origin.column) ? origin.column : undefined;
      if (originRow !== undefined && originColumn !== undefined && Array.isArray(clipboard?.values)) {
        const sourceRows = clipboard.values.length;
        const sourceColumns = clipboard.values.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
        const spec = input.spec && typeof input.spec === 'object' && !Array.isArray(input.spec)
          ? input.spec as { transpose?: unknown }
          : undefined;
        const rowCount = spec?.transpose === true ? sourceColumns : sourceRows;
        const columnCount = spec?.transpose === true ? sourceRows : sourceColumns;
        if (rowCount > 0 && columnCount > 0) {
          appendRange({
            sheetId,
            startRow: originRow,
            endRow: originRow + rowCount - 1,
            startColumn: originColumn,
            endColumn: originColumn + columnCount - 1,
          });
        }
      }
    }
    if (Array.isArray(input.ranges)) input.ranges.forEach(appendRange);
    if (input.filter && typeof input.filter === 'object') appendRange((input.filter as { range?: unknown }).range);
    if (input.rule && typeof input.rule === 'object' && Array.isArray((input.rule as { ranges?: unknown[] }).ranges)) {
      (input.rule as { ranges: unknown[] }).ranges.forEach(appendRange);
    }
    if (ranges.length === 0) ranges.push(this.getCurrentRegion());
    return sheet.dataRegions.filter((region) => ranges.some((range) => rangesIntersect(range, region.range)));
  }

  private commandRequiresResolvedWrites(commandId: string): boolean {
    return commandId === 'formula.autosum'
      || commandId === 'sheet.range.fill'
      || commandId === 'sheet.range.replace'
      || commandId === 'sheet.range.move'
      || commandId === 'sheet.range.paste'
      || commandId === 'sheet.range.clear'
      || commandId === 'sheet.range.clearContents'
      || commandId === 'sheet.cells.insert'
      || commandId === 'sheet.cells.delete'
      || commandId === 'sheet.style.set'
      || commandId === 'sheet.style.setMulti'
      || commandId === 'sheet.style.setMultiRange'
      || commandId === 'sheet.style.preset.apply'
      || commandId === 'sheet.cellTemplate.apply'
      || commandId === 'sheet.cellEditor.set'
      || commandId === 'sheet.format.set'
      || commandId === 'sheet.merge.set'
      || commandId === 'sheet.merge.remove'
      || commandId === 'sheet.merge.center'
      || commandId === 'sheet.merge.across'
      || commandId === 'sheet.autoFilter.toggle'
      || commandId === 'sheet.autoFilter.set'
      || commandId === 'sheet.autoFilter.sort'
      || commandId === 'sheet.autoFilter.clearCriteria'
      || commandId === 'sheet.autoFilter.reapply'
      || commandId === 'sheetTable.autoFilter.set'
      || commandId === 'sheet.cf.add'
      || commandId === 'sheet.cf.remove'
      || commandId === 'sheet.cf.clear'
      || commandId === 'sheet.dv.add'
      || commandId === 'sheet.dv.remove'
      || commandId === 'data.sort.quick'
      || commandId === 'data.sort.rows'
      || commandId === 'sheet.sort.multi'
      || commandId === 'data.sort.reapply'
      || commandId === 'data.splitColumn'
      || commandId === 'data.textToColumns'
      || commandId === 'data.subtotal'
      || commandId === 'data.removeDuplicates'
      || commandId === 'sheetTable.add'
      || commandId === 'sheetTable.toggleTotalRow'
      || commandId === 'selection.gotoSpecial'
      || commandId === 'sheet.rows.insert'
      || commandId === 'sheet.rows.delete'
      || commandId === 'sheet.columns.insert'
      || commandId === 'sheet.columns.delete';
  }

  /**
   * Catalog commands deliberately express intent, not duplicated UI state.
   * This is the single boundary that supplies the active worksheet/range for
   * commands whose domain contract operates on the current selection.
   */
  private resolveCommandContext(commandId: string, params?: unknown): unknown {
    const selectionScoped = new Set([
      'sheet.style.set',
      'sheet.merge.set',
      'sheet.merge.remove',
      'sheet.range.clear',
      'sheet.rows.insert',
      'sheet.rows.delete',
      'sheet.columns.insert',
      'sheet.columns.delete',
      'sheet.freeze.set',
    ]);
    if (!selectionScoped.has(commandId)) return params;
    if (params !== undefined && (typeof params !== 'object' || params === null || Array.isArray(params))) return params;
    const input = (params ?? {}) as Record<string, unknown>;
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    const sheetId = typeof input.sheetId === 'string' && input.sheetId.trim() ? input.sheetId : this.activeSheetId;
    if (commandId === 'sheet.style.set' || commandId === 'sheet.merge.set' || commandId === 'sheet.merge.remove') {
      return { ...input, sheetId, range: input.range ?? range };
    }
    if (commandId === 'sheet.range.clear') {
      return { ...input, sheetId, range: input.range ?? range, mode: input.mode ?? 'contents' };
    }
    if (commandId === 'sheet.rows.insert' || commandId === 'sheet.rows.delete') {
      return { ...input, sheetId, at: input.at ?? range.startRow, count: input.count ?? 1 };
    }
    if (commandId === 'sheet.columns.insert' || commandId === 'sheet.columns.delete') {
      return { ...input, sheetId, at: input.at ?? range.startColumn, count: input.count ?? 1 };
    }
    if (commandId === 'sheet.freeze.set') {
      return { ...input, sheetId };
    }
    return params;
  }

  /** Dispatches transient chrome state without touching WorkbookModel. */
  dispatchUiSessionIntent(intent: UiSessionIntent): void {
    switch (intent.type) {
      case 'panel.open':
        if (intent.panel === 'selectionPane') this.drawingSelectionMode = true;
        this.setActivePanel(intent.panel);
        if (intent.notice) this.notify(intent.notice);
        return;
      case 'dialog.open':
        this.openDialog(intent.dialog, intent.findQuery, intent.columnWidth, intent.sheet, intent.operation);
        return;
      case 'dialog.close':
        this.closeActiveDialog();
        return;
      case 'dialog.update':
        this.updateDialogDraft(intent.value);
        return;
      case 'command-palette.open':
        this.openCommandPalette();
        return;
      case 'command-palette.close':
        this.closeCommandPalette();
        return;
      case 'backstage.open':
        this.openBackstage(intent.panel);
        return;
      case 'zoom.set':
        this.setZoom(intent.value);
        return;
      case 'zoom.adjust':
        this.setZoom(intent.value ?? this.getZoom() + (intent.delta ?? 0));
        return;
      case 'notice':
        this.notify(intent.message);
        return;
    }
  }

  openBackstage(panel: BackstagePanel = 'info'): void {
    this.backstage = { open: true, panel };
    this.setFocusState('side-panel', 'side-panel');
    this.emit();
  }

  closeBackstage(): void {
    if (!this.backstage.open) return;
    this.backstage = { ...this.backstage, open: false };
    this.setFocusState('grid', 'grid');
    this.emit();
  }

  setBackstagePanel(panel: BackstagePanel): void {
    this.backstage = { open: true, panel };
    this.setFocusState('side-panel', 'side-panel');
    this.emit();
  }

  runCommand(commandId: string, params?: unknown): CommandResult {
    const resolvedParams = this.resolveCommandContext(commandId, params);
    if (!this.runtime.commands.registry.hasCommand(commandId)) {
      throw new Error(`Unknown command: ${commandId}`);
    }
    this.assertPermission(commandId, resolvedParams);
    const result = this.runtime.commands.execute(commandId, resolvedParams);
    if (commandId.startsWith('pivot.')) this.recomputeAllPivotResults();
    if (result.mutationCount > 0 && !commandId.startsWith('history.') && commandId !== 'pivot.refresh') {
      this.lastRepeatableCommand = { commandId, ...(resolvedParams === undefined ? {} : { params: structuredClone(resolvedParams) }) };
    }
    if (commandId === 'history.restore') {
      const restoreParams = resolvedParams as { targetRevision?: number };
      rehydrateFormulaAfterRestore(this.runtime, restoreParams.targetRevision);
      this.activeSheetId = this.runtime.model.primarySheetId;
      this.selectionService.resetForSheet(this.activeSheetId);
      this.clearHistoryPreview();
    }
    this.applySelectionFromCommand(commandId, resolvedParams, result);
    this.syncTableContextFromSelection();
    this.refresh();
    return result;
  }

  canRepeatLastCommand(): boolean {
    return this.lastRepeatableCommand !== null
      && this.canExecute(this.lastRepeatableCommand.commandId, this.lastRepeatableCommand.params);
  }

  repeatLastCommand(): void {
    if (!this.lastRepeatableCommand) return;
    const descriptor = this.lastRepeatableCommand;
    this.runCommand(descriptor.commandId, descriptor.params === undefined ? undefined : structuredClone(descriptor.params));
  }

  canExecute(commandId: string, params?: unknown): boolean {
    if (!this.runtime.commands.registry.hasCommand(commandId)) return false;
    const resolvedParams = this.resolveCommandContext(commandId, params);
    return canExecuteCommand(
      this.permission,
      this.runtime.model,
      commandId,
      resolvedParams,
      this.actorId,
      this.activeSheetId,
    ).allowed;
  }

  getShareRole(): ShareRole | null {
    return this.permission.getShareRole();
  }

  protectSelection(allowedActions: string[] = ['format']): void {
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    const rule = {
      id: nextId('protect'),
      scope: 'range' as const,
      sheetId: this.activeSheetId,
      range,
      locked: true,
      allow: {},
      allowedActions,
    };
    this.runCommand('sheet.protect.set', { sheetId: this.activeSheetId, rule });
    this.notify('Selection protected');
  }

  unprotectSelection(): void {
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    const rule = findProtectionRuleCoveringRange(this.runtime.model, this.activeSheetId, range);
    if (!rule) {
      this.notify('No protection rule covers the current selection');
      return;
    }
    this.runCommand('sheet.protect.remove', { sheetId: this.activeSheetId, ruleId: rule.id });
    this.notify('Selection unprotected');
  }

  private assertPermission(commandId: string, params?: unknown): void {
    this.permission.syncFromWorkbook(this.runtime.model);
    const result = this.permission.checkCommand(commandId, params, this.actorId, this.activeSheetId);
    if (!result.allowed) {
      throw new Error(result.reason ?? 'Permission denied');
    }
  }

  private static readonly SELECTION_COMMAND_IDS = new Set([
    'navigation.goto',
    'navigation.gotoSpecial',
    'selection.gotoSpecial',
    'selection.set',
  ]);

  private applySelectionFromCommand(commandId: string, params: unknown, result: CommandResult): void {
    if (commandId === 'selection.set') {
      const selectionParams = params as {
        ranges: RangeRef[];
        primaryRangeIndex?: number;
        primaryCell?: { row: number; column: number };
        anchorCell?: { row: number; column: number };
      };
      if (selectionParams.ranges.length === 0) return;
      const primaryIndex = selectionParams.primaryRangeIndex ?? 0;
      const primaryRange = selectionParams.ranges[primaryIndex] ?? selectionParams.ranges[0]!;
      const primary = selectionParams.primaryCell ?? { row: primaryRange.startRow, column: primaryRange.startColumn };
      this.selectionService.applyState({
        ranges: selectionParams.ranges,
        primaryRangeIndex: primaryIndex,
        activeCell: primary,
        anchorCell: selectionParams.anchorCell ?? primary,
      });
      this.syncDraftFromPrimary();
      return;
    }
    if (!WorkbookSession.SELECTION_COMMAND_IDS.has(commandId)) return;
    if (result.affectedRanges.length === 0) return;
    const first = result.affectedRanges[0]!;
    this.selectionService.applyState({
      ranges: result.affectedRanges,
      primaryRangeIndex: 0,
      activeCell: { row: first.startRow, column: first.startColumn },
      anchorCell: { row: first.startRow, column: first.startColumn },
    });
    this.syncDraftFromPrimary();
  }

  getClipboard(): ClipboardPayload | null {
    return this.clipboardData;
  }

  setClipboard(data: ClipboardPayload | null): void {
    this.clipboardData = data;
  }

  clearClipboard(): void {
    this.clipboardData = null;
  }

  selectAddress(address: string): boolean {
    const trimmed = address.trim();
    if (!trimmed) return false;
    if (this.editSession.editingCell) {
      return this.selectionService.selectCell(trimmed, {
        editing: true,
        insertRef: (ref) => this.setFormulaDraft(this.formulaDraft + ref),
      });
    }
    const range = parseRangeReference(trimmed);
    if (range) {
      this.selectionService.selectRange(range, 'replace');
      this.syncDraftFromPrimary();
      this.syncTableContextFromSelection();
      this.emit();
      return true;
    }
    try {
      this.runCommand('navigation.goto', { sheetId: this.activeSheetId, reference: trimmed });
      return true;
    } catch {
      this.notify(`Invalid reference: ${trimmed}`);
      return false;
    }
  }

  goTo(reference: string): void {
    this.runCommand('navigation.goto', { sheetId: this.activeSheetId, reference: reference.trim() });
  }

  goToSpecial(kind: GoToSpecialKind): void {
    const primary = this.getPrimaryRange();
    const range = primary.startRow === primary.endRow && primary.startColumn === primary.endColumn
      ? usedRangeOfSheet(this.runtime.model.getSheet(this.activeSheetId))
      : primary;
    this.dispatch({ commandId: 'selection.gotoSpecial', params: { sheetId: this.activeSheetId, range, kind } });
  }

  getActiveSheetId(): string {
    return this.activeSheetId;
  }

  setActivePivotContext(pivotId: string | null, sheetId = this.activeSheetId): void {
    const next: ActiveContext = pivotId === null ? { kind: 'none' } : { kind: 'pivot', sheetId, pivotId };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    if (pivotId !== null) {
      this.panels = { ...this.panels, active: 'pivot', open: true };
      this.ribbonTab = 'pivotAnalyze';
    } else if (this.ribbonTab === 'pivotAnalyze' || this.ribbonTab === 'pivotDesign') {
      this.ribbonTab = 'home';
    }
    this.emit();
  }

  setActiveTableSheetContext(sheetId: string | null, viewId?: string): void {
    const sheet = sheetId ? this.runtime.model.getSheet(sheetId) : undefined;
    const next: ActiveContext = sheet && sheet.kind === 'table-sheet' && sheet.tableSheet
      ? { kind: 'table-sheet', sheetId: sheet.id, viewId: viewId ?? sheet.tableSheet.viewId }
      : { kind: 'none' };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    if (next.kind === 'table-sheet') {
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'tableSheetDesign';
    } else if (this.ribbonTab === 'tableSheetDesign') {
      this.ribbonTab = 'home';
    }
    this.emit();
  }

  setActiveGanttSheetContext(sheetId: string | null, viewId?: string): void {
    const sheet = sheetId ? this.runtime.model.getSheet(sheetId) : undefined;
    const next: ActiveContext = sheet && sheet.kind === 'gantt-sheet' && sheet.ganttSheet
      ? { kind: 'gantt-sheet', sheetId: sheet.id, viewId: viewId ?? sheet.ganttSheet.viewId }
      : { kind: 'none' };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    if (next.kind === 'gantt-sheet') {
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'ganttTask';
    } else if (['ganttTask', 'ganttProject', 'ganttView', 'ganttFormat'].includes(this.ribbonTab)) {
      this.ribbonTab = 'home';
    }
    this.emit();
  }

  setActiveReportSheetContext(sheetId: string | null): void {
    const sheet = sheetId ? this.runtime.model.getSheet(sheetId) : undefined;
    const next: ActiveContext = sheet && sheet.kind === 'report-sheet' && sheet.reportSheet
      ? { kind: 'report-sheet', sheetId: sheet.id, ...(sheet.reportSheet.tableId ? { tableId: sheet.reportSheet.tableId } : {}) }
      : { kind: 'none' };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    if (next.kind === 'report-sheet') {
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'reportSheetDesign';
    } else if (this.ribbonTab === 'reportSheetDesign') {
      this.ribbonTab = 'home';
    }
    this.emit();
  }

  setActiveDrawingContext(drawingId: string | null, sheetId = this.activeSheetId): void {
    const next: ActiveContext = drawingId === null ? { kind: 'none' } : { kind: 'drawing', sheetId, drawingId };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    this.emit();
  }

  getActiveContext(): ActiveContext {
    return structuredClone(this.activeContext);
  }

  getSelection(): SelectionState {
    return this.selectionService.getState();
  }

  getPrimaryRange(): RangeRef {
    return this.selectionService.primaryRangeOrDefault();
  }

  getCurrentRegion(): RangeRef {
    const selection = this.selectionService.getState();
    const primary = this.getPrimaryRange();
    if (primary.startRow !== primary.endRow || primary.startColumn !== primary.endColumn) return primary;
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const table = findSheetTableAt(sheet, selection.activeCell.row, selection.activeCell.column);
    if (table) return structuredClone(table.range);
    const blockRegion = sheet.dataRegions.find((region) => selection.activeCell.row >= region.range.startRow
      && selection.activeCell.row <= region.range.endRow
      && selection.activeCell.column >= region.range.startColumn
      && selection.activeCell.column <= region.range.endColumn);
    if (blockRegion) return structuredClone(blockRegion.range);
    return currentRegionOfSheet(
      sheet,
      selection.activeCell.row,
      selection.activeCell.column,
    );
  }

  async storeDataBlock(ref: DataBlockRef, bytes: ArrayBuffer): Promise<void> {
    await this.runtime.dataBlocks.put(ref, bytes);
    this.notify(`Stored data block ${ref.id}`);
  }

  addDataSource(source: DataSourceManifest): void {
    const sheetId = source.sourceSheetId ?? this.activeSheetId;
    this.runCommand('dataSource.add', { sheetId, source });
    this.refresh();
  }

  updateDataSource(source: DataSourceManifest): void {
    const sheetId = source.sourceSheetId ?? this.activeSheetId;
    this.runCommand('dataSource.update', { sheetId, source });
    this.refresh();
  }

  removeDataSource(sourceId: string): void {
    this.runCommand('dataSource.remove', { sheetId: this.activeSheetId, sourceId });
    this.refresh();
  }

  addDataRegion(region: SheetDataRegion): void {
    this.runCommand('dataRegion.add', { sheetId: region.range.sheetId, region });
    this.refresh();
  }

  removeDataRegion(regionId: string): void {
    this.runCommand('dataRegion.remove', { sheetId: this.activeSheetId, regionId });
    this.refresh();
  }

  async loadDataBlock(ref: DataBlockRef): Promise<ArrayBuffer> {
    try {
      return await this.runtime.dataBlocks.get(ref);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Data block could not be loaded';
      this.notify(message);
      throw error;
    }
  }

  getSelectedSheet(): CanvasSheetSnapshot {
    return this.getUiSnapshot().selectedSheet;
  }

  getZoom(): number {
    return this.zoom;
  }

  getCollaborationSnapshot(): CollaborationSnapshot {
    if (!this.runtime.collaboration) {
      return {
        revision: this.runtime.remoteRevision,
        pendingCount: 0,
        offlineQueueState: 'offline',
        presence: { users: [], selections: [], editSessions: [], updatedAt: Date.now() },
        peerCount: this.peers.length,
      };
    }
    return buildCollaborationSnapshot(this.runtime.collaboration, this.peers);
  }

  async createGuestShareLink(role: GuestShareRole = 'editor'): Promise<string | null> {
    if (this.runtime.localOnly || typeof window === 'undefined') {
      this.notify('Connect to the Java backend before creating a guest share link');
      return null;
    }
    if (!this.canExecute('workbook.share.set')) {
      this.notify('Server access does not allow sharing this workbook');
      return null;
    }
    try {
      const share = await this.runtime.api.createGuestShare(this.runtime.model.unitId, { role });
      if (!share.token) throw new Error('Java backend did not return a guest share token');
      const link = `${window.location.origin}/workbooks/${encodeURIComponent(share.unitId)}?share=${encodeURIComponent(share.token)}`;
      await navigator.clipboard?.writeText(link);
      this.notify('Guest editor link copied');
      return link;
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Unable to create a guest share link');
      return null;
    }
  }

  async flushPendingCollaborations(): Promise<void> {
    if (!this.runtime.collaboration) {
      this.notify('Collaboration is offline');
      return;
    }
    const result = await this.runtime.collaboration.offlineQueue.flushAll();
    if (result.failed > 0) this.notify(`${result.failed} pending change set(s) failed to sync`);
    else if (result.flushed > 0) this.notify(`${result.flushed} pending change set(s) synced`);
    else this.notify('No pending collaboration changes');
    this.refresh();
  }

  async refreshRevisionLog(): Promise<void> {
    try {
      this.remoteRevisions = await this.runtime.api.listRevisions(this.runtime.model.unitId);
      this.emit();
    } catch {
      this.notify('Failed to refresh revision log');
    }
  }

  undoToHistoryIndex(index: number): void {
    const entries = this.runtime.commands.getUndoEntries();
    if (index < 0 || index >= entries.length) return;
    const undoCount = entries.length - 1 - index;
    for (let step = 0; step < undoCount; step += 1) {
      const entry = this.runtime.commands.getUndoEntries().at(-1);
      if (entry && !this.canReplayHistory(entry.undo)) {
        this.notify('Undo is no longer allowed for the protected selection');
        break;
      }
      if (!this.runtime.commands.undo()) break;
    }
    this.ensureActiveSheetSession();
    this.projectionGeneration += 1;
    this.reconcileDrawingSessionState();
    this.syncDraftFromPrimary();
    this.notify(`Restored session history to step ${index + 1}`);
    this.refresh();
  }

  restoreFromSnapshot(snapshot: import('@react-sheets/core-model').WorkbookSnapshot, targetRevision: number, reason?: string): void {
    try {
      void snapshot;
      this.runCommand('history.restore', { targetRevision, reason });
      this.notify(`Restore request submitted for revision ${targetRevision}`);
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Permission denied');
    }
  }

  async restoreToRevision(revision: number): Promise<void> {
    const response = await this.runtime.api.restoreToRevision(
      this.runtime.model.unitId,
      revision,
      `Restore to revision ${revision}`,
    );
    hydrateRuntime(this.runtime, response.snapshot);
    this.activeSheetId = this.runtime.model.primarySheetId;
    this.selectionService.resetForSheet(this.activeSheetId);
    this.projectionGeneration += 1;
    this.reconcileDrawingSessionState();
    this.clearHistoryPreview();
    this.notify(`Restored workbook to revision ${revision}`);
    this.refresh();
    await this.refreshRevisionLog();
  }

  async previewRevision(revision: number): Promise<HistoryPreviewSession | null> {
    try {
      const response = await this.runtime.api.getRevisionSnapshot(this.runtime.model.unitId, revision);
      const record = this.remoteRevisions.find((entry) => entry.revision === revision);
      const meta: HistoryEntryMeta = record
        ? revisionToHistoryMeta(record)
        : {
          revision,
          operationId: `preview-${revision}`,
          createdAt: new Date().toISOString(),
          description: `Revision ${revision}`,
        };
      this.historyPreview?.dispose();
      this.historyPreview = await HistoryPreviewSession.fromSnapshot(meta, response.snapshot);
      this.notify(`Previewing revision #${revision}`);
      this.emit();
      return this.historyPreview;
    } catch {
      this.notify('Failed to load revision preview');
      return null;
    }
  }

  clearHistoryPreview(): void {
    if (!this.historyPreview) return;
    this.historyPreview.dispose();
    this.historyPreview = null;
    this.emit();
  }

  getHistoryPreview(): HistoryPreviewSession | null {
    return this.historyPreview;
  }

  getPersistenceSnapshot(): PersistenceSnapshotMeta {
    return buildPersistenceMeta(
      this.runtime.model.snapshot(),
      this.runtime.remoteRevision,
      this.runtime.collaboration?.offlineQueue.getPendingCount() ?? 0,
      this.runtime.workspaceRecord,
    );
  }

  async saveWorkbook(reason = 'Manual save'): Promise<void> {
    this.saveState = 'saving';
    this.emit();
    try {
      void reason;
      const remoteWorkbook = !this.runtime.localOnly && this.runtime.remoteSyncRequested;
      if (remoteWorkbook && !this.runtime.remoteConnected) {
        await this.runtime.checkpointWorkspace();
        this.saveState = 'offline';
        this.syncPersistenceMeta();
        this.notify('Local checkpoint saved; waiting to sync with the server');
        return;
      }
      if (remoteWorkbook) {
        const flushed = await this.runtime.collaboration?.offlineQueue.flushAll();
        if (flushed && (flushed.failed > 0 || this.runtime.collaboration?.offlineQueue.getPendingCount())) {
          throw new Error('Pending changes could not be committed before checkpointing');
        }
        const checkpoint = await this.runtime.api.checkpointWorkbook(this.runtime.model.unitId);
        this.runtime.remoteRevision = checkpoint.snapshot.revision;
      }
      await this.runtime.checkpointWorkspace();
      this.saveState = 'saved';
      this.syncPersistenceMeta();
      this.notify(remoteWorkbook ? 'Workbook saved to server' : 'Local workbook checkpoint saved');
    } catch (error) {
      this.saveState = error instanceof Error && error.message.includes('conflict') ? 'conflict' : 'error';
      this.notify(error instanceof Error ? error.message : 'Save failed');
      this.emit();
      throw error instanceof Error ? error : new Error('Save failed');
    }
  }

  notify(message: string): void {
    this.notice = message;
    this.emit();
  }

  /** Permission is evaluated again before replaying a historical mutation. */
  private canReplayHistory(mutations: readonly MutationInfo[]): boolean {
    return mutations.every((mutation) => {
      const base = mutation.params && typeof mutation.params === 'object' && !Array.isArray(mutation.params)
        ? mutation.params as Record<string, unknown>
        : {};
      return canExecuteCommand(
        this.permission,
        this.runtime.model,
        mutation.id,
        { ...base, ranges: mutation.affectedRanges },
        this.actorId,
        mutation.sheetId,
      ).allowed;
    });
  }

  undo(): void {
    const entry = this.runtime.commands.getUndoEntries().at(-1);
    if (entry && !this.canReplayHistory(entry.undo)) {
      this.notify('Undo is no longer allowed for the protected selection');
      return;
    }
    if (this.runtime.commands.undo()) {
      this.ensureActiveSheetSession();
      this.projectionGeneration += 1;
      this.reconcileDrawingSessionState();
      this.syncDraftFromPrimary();
      this.notify('Undo applied');
      this.refresh();
    }
  }

  redo(): void {
    const entry = this.runtime.commands.getRedoEntries().at(-1);
    if (entry && !this.canReplayHistory(entry.redo)) {
      this.notify('Redo is no longer allowed for the protected selection');
      return;
    }
    if (this.runtime.commands.redo()) {
      this.ensureActiveSheetSession();
      this.projectionGeneration += 1;
      this.reconcileDrawingSessionState();
      this.syncDraftFromPrimary();
      this.notify('Redo applied');
      this.refresh();
    }
  }

  retry(): void {
    this.phase = 'ready';
    this.notify('Workspace ready');
    this.emit();
  }

  setRibbonTab(tab: RibbonTabId): void {
    this.ribbonTab = tab;
    this.setFocusState('ribbon', 'ribbon');
    this.emit();
  }

  setActivePanel(panel: SidebarPanelId): void {
    this.panels = { ...this.panels, active: panel, open: true };
    this.emit();
  }

  setPanelOpen(open: boolean): void {
    this.panels = { ...this.panels, open };
    this.emit();
  }

  setZoom(zoom: number): void {
    this.zoom = Math.max(50, Math.min(200, zoom));
    this.emit();
  }

  setFocusState(mode: InputMode, target: FocusState['target'] = mode === 'cell-edit'
    ? 'grid'
    : mode === 'formula-edit'
      ? 'formula-bar'
      : mode === 'ribbon-keytip'
        ? 'ribbon'
        : mode === 'dropdown'
          ? 'ribbon'
          : mode): void {
    this.inputMode = mode;
    this.focus = { mode, target };
  }

  openCommandPalette = (): void => {
    this.dialogs = { ...this.dialogs, active: 'command-palette' };
    this.setFocusState('command-palette', 'command-palette');
    this.emit();
  };

  closeCommandPalette = (): void => {
    if (this.dialogs.active === 'command-palette') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };

  openDialog(dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells' | 'create-pivot' | 'create-table' | 'column-width' | 'sheet-rename' | 'sheet-tab-color' | 'sheet-delete' | 'cell-template' | 'cell-editor' | 'insert-picture', findQuery?: string, columnWidth?: { columns: number[]; defaultMode: boolean }, sheet?: SheetDialogState, operation: CellShiftOperation = 'insert'): void {
    this.setFocusState('dialog', 'dialog');
    const active = dialog === 'sheet-rename' || dialog === 'sheet-tab-color' || dialog === 'sheet-delete' ? 'sheet-dialog' : dialog;
    this.dialogs = { ...this.dialogs, active, cellShiftOperation: dialog === 'shift-cells' ? operation : this.dialogs.cellShiftOperation, findQuery: dialog === 'find-replace' ? findQuery ?? '' : this.dialogs.findQuery, columnWidth: dialog === 'column-width' ? structuredClone(columnWidth ?? { columns: [], defaultMode: false }) : null, sheet: sheet ? structuredClone(sheet) : null };
    if (dialog === 'print-preview') {
      this.rebuildPrintSnapshot();
      this.panels = { ...this.panels, active: 'print', open: true };
    }
    this.emit();
  }

  closeActiveDialog(): void {
    if (!this.dialogs.active) return;
    this.dialogs = { ...this.dialogs, active: null, columnWidth: null, sheet: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  }

  updateDialogDraft(value: string): void {
    if (!this.dialogs.sheet) return;
    this.dialogs = { ...this.dialogs, sheet: { ...this.dialogs.sheet, value } };
    this.emit();
  }

  closeFunctionWizard = (): void => {
    if (this.dialogs.active === 'function-wizard') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeSortDialog = (): void => {
    if (this.dialogs.active === 'sort-dialog') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeFindReplace = (): void => {
    if (this.dialogs.active === 'find-replace') this.dialogs = { ...this.dialogs, active: null, findQuery: '' };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeGoTo = (): void => {
    if (this.dialogs.active === 'goto') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closePasteSpecial = (): void => {
    if (this.dialogs.active === 'paste-special') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeFormatCells = (): void => {
    if (this.dialogs.active === 'format-cells') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeShiftCells = (): void => {
    if (this.dialogs.active === 'shift-cells') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeCreatePivotDialog = (): void => {
    if (this.dialogs.active === 'create-pivot') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  closeCreateTableDialog = (): void => {
    if (this.dialogs.active === 'create-table') this.dialogs = { ...this.dialogs, active: null };
    this.setFocusState('grid', 'grid');
    this.emit();
  };
  async pasteSpecial(spec: PasteSpecialSpec): Promise<DispatchOutcome> {
    const outcome = await this.paste(spec);
    this.closePasteSpecial();
    return outcome;
  };
  setShowPrintPreview = (open: boolean): void => {
    this.dialogs = { ...this.dialogs, active: open ? 'print-preview' : null };
    this.setFocusState(open ? 'dialog' : 'grid', open ? 'dialog' : 'grid');
    this.emit();
  };

  syncDraftFromPrimary(): void {
    const sel = this.selectionService.getState();
    const cell = this.readResolvedCell(this.runtime.model.getSheet(this.activeSheetId), sel.activeCell.row, sel.activeCell.column);
    this.formulaDraft = cell?.formula ?? (cell?.value == null ? '' : String(cell.value));
    this.emit();
  }

  setFormulaDraft(value: string): void {
    this.formulaDraft = value;
    if (this.editSession.active) this.editSession.setDraft(value);
    this.emit();
  }

  appendFormulaDraft(fragment: string): void {
    if (!fragment) return;
    this.setFormulaDraft(this.formulaDraft + fragment);
  }

  selectCell(address: string): void {
    const changed = this.selectionService.selectCell(address, {
      editing: Boolean(this.editSession.editingCell),
      insertRef: (ref) => this.setFormulaDraft(this.formulaDraft + ref),
    });
    if (changed) this.syncDraftFromPrimary();
    this.syncTableContextFromSelection();
    this.emit();
  }

  selectRange(range: { startRow: number; startColumn: number; endRow: number; endColumn: number }, mode: 'replace' | 'add' | 'extend' = 'replace'): void {
    this.selectionService.selectRange(range, mode);
    this.syncDraftFromPrimary();
    this.syncTableContextFromSelection();
    this.emit();
  }

  /** Commit a canvas selection exactly, including its active cell and anchor. */
  applyCanvasSelection(selection: SelectionState): void {
    this.selectionService.applyState(selection);
    if (this.formatPainter) {
      const painter = this.formatPainter;
      const targetRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
      if (targetRange) this.dispatch({ commandId: 'format.painter.apply', params: {
        sheetId: this.activeSheetId,
        sourceRange: structuredClone(painter.sourceRange),
        targetRange: structuredClone(targetRange),
        continuous: painter.mode === 'locked',
      } });
      if (painter.mode === 'once') this.formatPainter = null;
    }
    this.syncDraftFromPrimary();
    this.syncTableContextFromSelection();
    this.emit();
  }

  extendSelectionTo(row: number, column: number): void {
    this.selectRange({ startRow: row, startColumn: column, endRow: row, endColumn: column }, 'extend');
  }

  formatCells(params: { numberFormat?: string; style?: Partial<import('@react-sheets/core-model').CellStyle> }): void {
    const ranges = this.selectionService.getState().ranges;
    if (ranges.length === 0) return;
    this.dispatch({ commandId: 'sheet.format.set', params: {
      sheetId: this.activeSheetId,
      ranges,
      numberFormat: params.numberFormat,
      style: params.style,
    } });
  }

  beginFormatPainter(locked = false): void {
    const sourceRange = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    this.formatPainter = { mode: locked ? 'locked' : 'once', sourceRange };
    this.notify(locked ? 'Format Painter is locked; select target ranges and press Escape to finish' : 'Select a target range to apply copied formatting');
    this.emit();
  }

  cancelFormatPainter(): void {
    if (!this.formatPainter) return;
    this.formatPainter = null;
    this.notify('Format Painter cancelled');
    this.emit();
  }

  requestMergeCells(): void {
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
      this.notify('Select at least two cells before merging');
      return;
    }
    const regions = this.dataRegionsIntersectingRanges(this.activeSheetId, [range]);
    if (regions.length > 0) {
      void this.materializeDataRegions(regions)
        .then(() => this.requestMergeCells())
        .catch((error) => this.notify(error instanceof Error ? error.message : 'Data region could not be prepared for merging'));
      return;
    }
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    let discardCount = 0;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        if (row === range.startRow && column === range.startColumn) continue;
        const cell = this.readResolvedCell(sheet, row, column);
        if (cell && (cell.formula || (cell.value !== null && cell.value !== undefined && cell.value !== ''))) discardCount += 1;
      }
    }
    if (discardCount > 0) {
      this.pendingMergeRange = range;
      this.dialogs = { ...this.dialogs, active: 'merge-confirm', mergeDiscardCount: discardCount };
      this.setFocusState('dialog', 'dialog');
      this.emit();
      return;
    }
    this.dispatch({ commandId: 'sheet.merge.center', params: { sheetId: this.activeSheetId, range, confirmDataLoss: true } });
  }

  confirmMergeCells(): void {
    const range = this.pendingMergeRange;
    this.dialogs = { ...this.dialogs, active: null, mergeDiscardCount: 0 };
    this.setFocusState('grid', 'grid');
    this.pendingMergeRange = null;
    if (range) this.dispatch({ commandId: 'sheet.merge.center', params: { sheetId: range.sheetId, range, confirmDataLoss: true } });
    this.emit();
  }

  cancelMergeCells(): void {
    this.dialogs = { ...this.dialogs, active: null, mergeDiscardCount: 0 };
    this.setFocusState('grid', 'grid');
    this.pendingMergeRange = null;
    this.emit();
  }

  applyCellShift(operation: CellShiftOperation, axis: 'row' | 'column'): void {
    const range = this.getPrimaryRange();
    this.dispatch({ commandId: operation === 'insert' ? 'sheet.cells.insert' : 'sheet.cells.delete', params: { sheetId: this.activeSheetId, range, operation, axis } });
  }

  freezeAtPrimary(): void {
    const sel = this.selectionService.getState();
    this.runCommand('sheet.freeze.set', {
      sheetId: this.activeSheetId,
      pane: {
        kind: 'frozen',
        xSplit: sel.activeCell.column,
        ySplit: sel.activeCell.row,
        startRow: sel.activeCell.row,
        startColumn: sel.activeCell.column,
        state: 'frozen',
      },
    });
    this.refresh();
  }

  movePrimary(rowDelta: number, columnDelta: number, opts?: { extend?: boolean }): void {
    this.selectionService.movePrimary(rowDelta, columnDelta, opts);
    if (!this.editSession.editingCell) {
      this.syncDraftFromPrimary();
    }
    this.emit();
  }

  jumpEdge(direction: 'up' | 'down' | 'left' | 'right', extend = false): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const sel = this.selectionService.getState();
    let row = sel.activeCell.row;
    let column = sel.activeCell.column;
    const step = direction === 'up' ? -1 : direction === 'down' ? 1 : direction === 'left' ? -1 : 1;
    const horizontal = direction === 'left' || direction === 'right';
    let cursor = horizontal ? column : row;
    cursor += step;
    while (cursor >= 0 && (horizontal ? cursor < sheet.columnCount : cursor < sheet.rowCount)) {
      const cellValue = horizontal
        ? this.readResolvedCell(sheet, row, cursor)?.value
        : this.readResolvedCell(sheet, cursor, column)?.value;
      if (cellValue != null && cellValue !== '') break;
      cursor += step;
    }
    if (horizontal) column = Math.max(0, Math.min(sheet.columnCount - 1, cursor));
    else row = Math.max(0, Math.min(sheet.rowCount - 1, cursor));
    this.movePrimary(row - sel.activeCell.row, column - sel.activeCell.column, { extend });
    this.emit();
  }

  selectAll(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    this.selectionService.selectAll(sheet.rowCount, sheet.columnCount);
    this.syncDraftFromPrimary();
    this.emit();
  }

  selectActiveRow(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selection = this.selectionService.getState();
    this.selectionService.selectRow(selection.activeCell.row, sheet.columnCount);
    this.syncDraftFromPrimary();
    this.emit();
  }

  selectActiveColumn(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selection = this.selectionService.getState();
    this.selectionService.selectColumn(selection.activeCell.column, sheet.rowCount);
    this.syncDraftFromPrimary();
    this.emit();
  }

  selectAdjacentSheet(direction: 'previous' | 'next'): void {
    const sheets = this.runtime.model.getVisibleSheets();
    const current = sheets.findIndex((sheet) => sheet.id === this.activeSheetId);
    if (current < 0 || sheets.length < 2) return;
    const offset = direction === 'next' ? 1 : -1;
    const next = sheets[(current + offset + sheets.length) % sheets.length];
    if (next) this.selectSheet(next.id);
  }

  autoSum(): void {
    const selection = this.selectionService.getState();
    const range = this.getCurrentRegion();
    this.dispatch({
      commandId: 'formula.autosum',
      params: {
        sheetId: this.activeSheetId,
        range,
        target: { ...selection.activeCell },
      },
    });
  }

  listDefinedNames(sheetId = this.activeSheetId): readonly DefinedNameModel[] {
    this.runtime.model.getSheet(sheetId);
    return this.runtime.model.listDefinedNames(sheetId);
  }

  getDefinedName(name: string, sheetId = this.activeSheetId): DefinedNameModel | undefined {
    this.runtime.model.getSheet(sheetId);
    return this.runtime.model.getDefinedName(name, sheetId);
  }

  setDefinedName(input: DefinedNameCommandInput): DefinedNameModel {
    const scope = input.scope ?? 'workbook';
    const sheetId = scope === 'sheet' ? input.sheetId ?? this.activeSheetId : undefined;
    this.runCommand('workbook.name.set', {
      name: input.name,
      formula: input.formula,
      scope,
      ...(sheetId === undefined ? {} : { sheetId }),
      ...(input.hidden === undefined ? {} : { hidden: input.hidden }),
      ...(input.comment === undefined ? {} : { comment: input.comment }),
    });
    this.refresh();
    const model = this.runtime.model.getDefinedNameExact(input.name, scope, sheetId);
    if (!model) throw new Error(`Defined name was not persisted: ${input.name}`);
    return model;
  }

  removeDefinedName(
    name: string,
    scope: DefinedNameModel['scope'] = 'workbook',
    sheetId = scope === 'sheet' ? this.activeSheetId : undefined,
  ): DefinedNameModel | undefined {
    const previous = this.runtime.model.getDefinedNameExact(name, scope, sheetId);
    this.runCommand('workbook.name.remove', {
      name,
      scope,
      ...(sheetId === undefined ? {} : { sheetId }),
    });
    this.refresh();
    return previous;
  }

  showFormulaPrecedents(): void {
    const active = this.selectionService.getState().activeCell;
    this.runCommand('formula.audit.precedents.show', { address: { sheetId: this.activeSheetId, ...active } });
    this.refresh();
  }

  showFormulaDependents(): void {
    const active = this.selectionService.getState().activeCell;
    this.runCommand('formula.audit.dependents.show', { address: { sheetId: this.activeSheetId, ...active } });
    this.refresh();
  }

  removeFormulaAuditArrows(): void {
    this.runCommand('formula.audit.arrows.remove', {});
    this.refresh();
  }

  setShowFormulas(enabled: boolean): void {
    this.runCommand('formula.audit.formulas.show', { enabled });
    this.refresh();
  }

  scanFormulaErrors(): void {
    this.runCommand('formula.audit.errors.scan', { sheetId: this.activeSheetId });
    this.refresh();
  }

  evaluateFormulaStep(): void {
    const active = this.selectionService.getState().activeCell;
    this.runCommand('formula.audit.evaluate.step', { address: { sheetId: this.activeSheetId, ...active } });
    this.refresh();
  }

  private resolveCanonicalCellTarget(sheetId: string, row: number, column: number): { sheetId: string; row: number; column: number } {
    const sheet = this.runtime.model.getSheet(sheetId);
    const tableId = sheet.kind === 'table-sheet' ? sheet.tableSheet?.viewId : sheet.kind === 'gantt-sheet' ? sheet.ganttSheet?.viewId : undefined;
    const table = tableId ? this.runtime.model.dataModel.tables.get(tableId) : undefined;
    const field = table?.fields[column];
    if (!table?.sourceRange || !field || row <= 0) return { sheetId, row, column };
    const sourceRow = table.sourceRange.startRow + row;
    if (sourceRow > table.sourceRange.endRow) return { sheetId, row, column };
    return { sheetId: table.sourceRange.sheetId, row: sourceRow, column: table.sourceRange.startColumn + field.ordinal };
  }

  beginEdit(initialText?: string): boolean {
    const sel = this.selectionService.getState();
    const canonicalTarget = this.resolveCanonicalCellTarget(this.activeSheetId, sel.activeCell.row, sel.activeCell.column);
    const canonicalSheet = this.runtime.model.getSheet(canonicalTarget.sheetId);
    const target = { ...canonicalTarget, text: initialText ?? '' };
    const permission = canExecuteCommand(this.permission, this.runtime.model, 'sheet.cell.commitText', target, this.actorId, target.sheetId);
    if (!permission.allowed) {
      this.notify(permission.reason ?? 'This cell is not editable');
      return false;
    }
    for (const spill of canonicalSheet.spillRanges) {
      if (isSpillChild(spill, canonicalTarget.row, canonicalTarget.column)) {
        this.notify('Spill cells are read-only');
        return false;
      }
    }
    const cell = this.readResolvedCell(canonicalSheet, canonicalTarget.row, canonicalTarget.column);
    this.editSession.begin({
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      cell,
      selection: this.selectionService.getSnapshot(),
      initialText,
    });
    this.setFocusState('cell-edit', 'grid');
    this.formulaDraft = this.editSession.active?.currentDraft ?? '';
    this.emit();
    return true;
  }

  cancelEdit(): void {
    this.formulaDraft = this.editSession.cancel();
    this.setFocusState('grid', 'grid');
    this.emit();
  }

  commitFormula(overrideValue?: string): boolean {
    if (this.phase !== 'ready') return false;
    const sel = this.selectionService.getState();
    const displayRow = this.overrideTarget?.row ?? sel.activeCell.row;
    const displayColumn = this.overrideTarget?.column ?? sel.activeCell.column;
    const target = this.resolveCanonicalCellTarget(this.activeSheetId, displayRow, displayColumn);
    const row = target.row;
    const column = target.column;
    const text = overrideValue !== undefined ? overrideValue : this.formulaDraft;
    const style = this.runtime.model.getSheet(target.sheetId).cells.get(row, column)?.style;
    try {
      this.runCommand('sheet.cell.commitText', {
        sheetId: target.sheetId,
        row,
        column,
        text,
        style,
      });
      return true;
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Cell input was rejected');
      return false;
    }
  }

  commitEdit(moveAfter: 'down' | 'up' | 'left' | 'right' | 'none' = 'down'): void {
    if (!this.editSession.editingCell) return;
    const editing = this.editSession.editingCell;
    this.overrideTarget = editing;
    const committed = this.commitFormula(this.formulaDraft);
    this.overrideTarget = null;
    this.editSession.apply();
    if (!committed) {
      this.beginEdit(this.formulaDraft);
      return;
    }
    const deltas = { down: [1, 0], up: [-1, 0], left: [0, -1], right: [0, 1], none: [0, 0] } as const;
    const [dr, dc] = deltas[moveAfter];
    this.movePrimary(dr, dc);
    this.syncDraftFromPrimary();
    this.setFocusState('grid', 'grid');
    this.emit();
  }

  insertRefIntoDraft(refText: string): void {
    if (this.editSession.active) this.editSession.insertRef(refText);
    else this.setFormulaDraft(this.formulaDraft + refText);
  }

  toggleAbsoluteReference(): void {
    if (this.editSession.active) {
      this.editSession.toggleAbsoluteReference();
      this.formulaDraft = this.editSession.active.currentDraft;
      this.emit();
      return;
    }
    this.repeatLastCommand();
  }

  selectSheet(sheetId: string): void {
    const sheet = this.runtime.model.getSheet(sheetId);
    this.activeSheetId = sheetId;
    this.selectionService.resetForSheet(sheetId);
    this.editSession.cancel();
    this.formulaDraft = '';
    if (sheet.kind === 'table-sheet' && sheet.tableSheet) {
      this.activeContext = { kind: 'table-sheet', sheetId: sheet.id, viewId: sheet.tableSheet.viewId };
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'tableSheetDesign';
    } else if (sheet.kind === 'gantt-sheet' && sheet.ganttSheet) {
      this.activeContext = { kind: 'gantt-sheet', sheetId: sheet.id, viewId: sheet.ganttSheet.viewId };
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'ganttTask';
    } else if (sheet.kind === 'report-sheet' && sheet.reportSheet) {
      this.activeContext = { kind: 'report-sheet', sheetId: sheet.id, ...(sheet.reportSheet.tableId ? { tableId: sheet.reportSheet.tableId } : {}) };
      this.panels = { ...this.panels, active: 'data', open: true };
      this.ribbonTab = 'reportSheetDesign';
    } else {
      this.activeContext = { kind: 'none' };
      if (this.ribbonTab === 'tableSheetDesign' || ['ganttTask', 'ganttProject', 'ganttView', 'ganttFormat'].includes(this.ribbonTab) || this.ribbonTab === 'reportSheetDesign') this.ribbonTab = 'home';
    }
    this.runtime.drawing.deselect(sheetId);
    this.drawingSelectionMode = false;
    this.reconcileDrawingSessionState();
    this.refresh();
  }

  // ---- Pro / data features ----

  private buildSelectionWorkbookTable(prefix: string): WorkbookTableModel {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selected = normalizeRangeRef(this.getPrimaryRange());
    const sourceRange = selected.startRow === selected.endRow && selected.startColumn === selected.endColumn && !sheet.cells.has(selected.startRow, selected.startColumn)
      ? undefined
      : { ...selected, sheetId: this.activeSheetId };
    const defaults = prefix === 'gantt'
      ? ['Task ID', 'Task', 'Start', 'End', 'Progress', 'Parent ID', 'Dependencies']
      : ['Name', 'Value'];
    const fields: WorkbookTableModel['fields'] = [];
    const count = sourceRange ? sourceRange.endColumn - sourceRange.startColumn + 1 : defaults.length;
    const names = new Set<string>();
    for (let offset = 0; offset < count; offset += 1) {
      const column = (sourceRange?.startColumn ?? 0) + offset;
      const raw = sourceRange ? String(this.readResolvedCell(sheet, sourceRange.startRow, column)?.value ?? '').trim() : defaults[offset] ?? `Column ${offset + 1}`;
      let name = raw || defaults[offset] || `Column ${offset + 1}`;
      let suffix = 2;
      while (names.has(name)) name = `${raw || 'Column'} ${suffix++}`;
      names.add(name);
      const values: CellData['value'][] = [];
      if (sourceRange) for (let row = sourceRange.startRow + 1; row <= Math.min(sourceRange.endRow, sourceRange.startRow + 1000); row += 1) values.push(this.readResolvedCell(sheet, row, column)?.value ?? null);
      fields.push({ id: `${prefix}-field-${offset + 1}`, name, ordinal: offset, type: inferTableFieldType(values) });
    }
    return {
      id: nextId(`${prefix}-table`), name: `${sheet.name} ${prefix} data`, sourceSheetId: sourceRange ? this.activeSheetId : undefined,
      sourceRange, rowCount: sourceRange ? Math.max(0, sourceRange.endRow - sourceRange.startRow) : 0,
      fields, blockSize: 4096, blocks: [], revision: 0,
    };
  }

  createAdvancedSheet(kind: Exclude<SheetKind, 'worksheet'>): void {
    const table = this.buildSelectionWorkbookTable(kind === 'gantt-sheet' ? 'gantt' : kind === 'report-sheet' ? 'report' : 'table');
    const id = nextId('sheet');
    const name = kind === 'table-sheet' ? '集算表' : kind === 'gantt-sheet' ? '甘特表' : '报表';
    const cells: SheetSnapshot['cells'] = {};
    const setCell = (row: number, column: number, value: CellData['value'], style?: CellStyle) => {
      cells[String(row)] ??= {};
      cells[String(row)]![String(column)] = { value, ...(style ? { style } : {}) };
    };
    table.fields.forEach((field, column) => setCell(0, column, field.name, { bold: true, background: '#eaf2f8', verticalAlignment: 'middle' }));
    if (kind === 'report-sheet') setCell(0, 0, '报表', { bold: true, fontSizePx: 24, textColor: '#1f2937' });
    const sheet: SheetSnapshot = {
      kind, id, name, rowCount: 1000, columnCount: Math.max(26, table.fields.length), cells, merges: [], pane: { kind: 'none' }, pivots: [], sparklines: [], drawings: [], drawingPayloads: {},
      defaultRowHeightPx: 20, defaultColumnWidthPx: 80,
      ...(kind === 'table-sheet' ? { tableSheet: { viewId: table.id, columns: table.fields.map((field) => ({ fieldId: field.id, caption: field.name, type: field.type })), grouping: [] } } : {}),
      ...(kind === 'gantt-sheet' ? { ganttSheet: { viewId: table.id, fieldMap: { id: table.fields[0]!.id, title: table.fields[1]!.id, start: table.fields[2]!.id, end: table.fields[3]!.id, progress: table.fields[4]!.id, parentId: table.fields[5]?.id, dependencies: table.fields[6]?.id }, calendar: { workingDays: [1, 2, 3, 4, 5], dayStartHour: 9, dayEndHour: 18 }, timeline: { unit: 'week' }, dependencyStyle: { color: '#64748b', width: 1 } } } : {}),
      ...(kind === 'report-sheet' ? { reportSheet: { templateSheetId: this.activeSheetId, tableId: table.id, bindings: [], pagination: { enabled: true, rowsPerPage: 50, repeatHeaderRows: [0] }, renderMode: 'design' as const, layout: { orientation: 'portrait' as const, marginTopPx: 24, marginRightPx: 24, marginBottomPx: 24, marginLeftPx: 24 }, dataEntry: [] } } : {}),
    };
    this.runCommand('sheet.create.advanced', { sheet, table, index: this.runtime.model.sheetOrder.length });
    this.selectSheet(id);
    this.notify(`${name}已创建`);
  }

  updateTableSheetDefinition(definition: TableSheetDefinition): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    if (sheet.kind !== 'table-sheet') throw new Error('TableSheet designer requires a table-sheet');
    this.runCommand('tableSheet.update', { sheetId: this.activeSheetId, definition });
  }

  updateGanttSheetDefinition(definition: import('@react-sheets/core-model').GanttSheetDefinition): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    if (sheet.kind !== 'gantt-sheet') throw new Error('GanttSheet designer requires a gantt-sheet');
    this.runCommand('ganttSheet.update', { sheetId: this.activeSheetId, definition });
  }

  updateReportSheetDefinition(definition: import('@react-sheets/core-model').ReportSheetDefinition): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    if (sheet.kind !== 'report-sheet') throw new Error('ReportSheet designer requires a report-sheet');
    this.runCommand('reportSheet.update', { sheetId: this.activeSheetId, definition });
  }

  applyBarcode(symbology: BarcodeSymbology = 'qr'): void {
    this.barcodeDraftSymbology = symbology;
    const ranges = this.selectionService.getState().ranges.map((range) => ({ ...range, sheetId: this.activeSheetId }));
    this.runCommand('cell.barcode.apply', { sheetId: this.activeSheetId, ranges, presentation: { kind: 'barcode', symbology, source: { kind: 'cell-value' }, parameters: { symbology }, options: { foreground: '#111827', background: '#ffffff', showText: symbology !== 'qr' && symbology !== 'data-matrix', labelPosition: symbology === 'qr' || symbology === 'data-matrix' ? 'none' : 'below', quietZone: 2 } } });
    this.panels = { ...this.panels, active: 'barcode', open: true };
    this.notify('条形码已应用');
    this.refresh();
  }

  openBarcodePanel(symbology: BarcodeSymbology = 'qr'): void {
    this.barcodeDraftSymbology = symbology;
    this.panels = { ...this.panels, active: 'barcode', open: true };
    this.refresh();
  }

  getBarcodeDraftSymbology(): BarcodeSymbology {
    return this.barcodeDraftSymbology;
  }

  insertDataChart(type: DataChartPlotType = 'column'): void {
    const table = this.buildSelectionWorkbookTable('data-chart');
    if (!table.sourceRange) throw new Error('Data Chart requires a non-empty selected range with a header row');
    const numericFields = table.fields.filter((field) => field.type === 'number');
    if (numericFields.length === 0) throw new Error('Data Chart requires at least one numeric value field');
    const category = table.fields.find((field) => field.type !== 'number') ?? table.fields[0];
    const bindings: DataChartDrawingPayload['bindings'] = { values: numericFields.map((field) => ({ area: 'values', fieldId: field.id, aggregate: 'sum' })), category: category ? [{ area: 'category', fieldId: category.id, aggregate: 'none' }] : [], details: [], color: [], size: [], tooltip: [], filter: [] };
    const payloadId = nextId('data-chart');
    const drawing: DrawingObject = { id: nextId('drawing'), sheetId: this.activeSheetId, kind: 'data-chart', anchor: { kind: 'absolute' }, transform: { x: 96, y: 96, width: 480, height: 280, rotation: 0 }, zIndex: 0, payloadId };
    const payload: DataChartDrawingPayload = { kind: 'data-chart', source: { kind: 'table', tableId: table.id }, plotType: type, bindings, inspector: { title: '数据图表', legendPosition: 'bottom', showDataLabels: false, showHiddenData: true, chartArea: { fill: '#ffffff', border: '#cbd5e1', borderWidth: 1 }, plotArea: { fill: '#ffffff' }, axis: { showGridlines: true } } };
    this.runCommand('dataChart.create', { sheetId: this.activeSheetId, drawing, payload, table });
    this.setDrawingSelection([drawing.id]);
    this.panels = { ...this.panels, active: 'dataChart', open: true };
    this.notify('数据图表已插入');
    this.refresh();
  }

  insertCamera(): void {
    const range = { ...normalizeRangeRef(this.getPrimaryRange()), sheetId: this.activeSheetId };
    const payloadId = nextId('camera');
    const drawing: DrawingObject = { id: nextId('drawing'), sheetId: this.activeSheetId, kind: 'camera', anchor: { kind: 'absolute' }, transform: { x: 96, y: 96, width: 360, height: 220, rotation: 0 }, zIndex: 0, payloadId };
    const payload: CameraDrawingPayload = { kind: 'camera', sourceRange: range, refreshPolicy: 'live' };
    this.runCommand('drawing.add.camera', { sheetId: this.activeSheetId, drawing, payload });
    this.setDrawingSelection([drawing.id]);
    this.notify('区域快照已插入');
    this.refresh();
  }

  insertFormControl(controlType: FormControlType = 'button'): void {
    const active = this.selectionService.getState().activeCell;
    const payloadId = nextId('form-control');
    const drawing: DrawingObject = { id: nextId('drawing'), sheetId: this.activeSheetId, kind: 'form-control', anchor: { kind: 'one-cell', row: active.row, column: active.column }, transform: { x: 96, y: 96, width: 140, height: 32, rotation: 0 }, zIndex: 0, payloadId };
    const payload: FormControlDrawingPayload = { kind: 'form-control', controlType, text: controlType === 'button' ? '按钮' : controlType, cellLink: { sheetId: this.activeSheetId, row: active.row, column: active.column }, value: false, enabled: true, style: { fill: '#ffffff', border: '#7b8794', textColor: '#1f2937', fontSize: 12 } };
    this.runCommand('drawing.add.form-control', { sheetId: this.activeSheetId, drawing, payload });
    this.setDrawingSelection([drawing.id]);
    this.notify('控件已插入');
    this.refresh();
  }

  insertTextBox(): void {
    const payloadId = nextId('textbox');
    const drawing: DrawingObject = { id: nextId('drawing'), sheetId: this.activeSheetId, kind: 'textbox', anchor: { kind: 'absolute' }, transform: { x: 96, y: 96, width: 220, height: 72, rotation: 0 }, zIndex: 0, payloadId };
    this.runCommand('drawing.add.textbox', { sheetId: this.activeSheetId, drawing, payload: { kind: 'textbox', text: '文本框', textColor: '#1f2937', fontSize: 14 } });
    this.setDrawingSelection([drawing.id]);
    this.notify('文本框已插入');
    this.refresh();
  }

  addChart(drawing: DrawingObject, payload: ChartDrawingPayload): void {
    if (drawing.kind !== 'chart' || payload.kind !== 'chart' || drawing.payloadId !== payload.chartId) {
      throw new Error(`Chart drawing and payload identity mismatch: ${drawing.id}`);
    }
    this.runCommand('chart.insert', buildChartInsertParams(drawing, payload));
    this.setDrawingSelection([drawing.id]);
    this.notify(payload.elements.title ? `Added chart "${payload.elements.title}"` : `Added ${payload.chartType} chart`);
    this.refresh();
  }
  insertChart(type: ChartDrawingPayload['chartType'] = 'column'): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    const payloadId = nextId('chart');
    const drawing: DrawingObject = {
      id: nextId('draw'),
      sheetId: this.activeSheetId,
      kind: 'chart',
      anchor: { kind: 'absolute' },
      transform: { x: 96, y: 96, width: 480, height: 280, rotation: 0 },
      zIndex: 0,
      payloadId,
    };
    const payload: ChartDrawingPayload = {
      kind: 'chart',
      chartId: payloadId,
      chartType: type,
      sourceRanges: [{ ...range, sheetId: this.activeSheetId }],
      series: type === 'combo' ? [{ name: 'Series 1', range: { ...range, sheetId: this.activeSheetId }, chartType: 'column', axis: 'primary' }] : undefined,
      elements: {
        title: 'Chart',
        legend: { visible: true, position: 'bottom' },
        dataLabels: { visible: false },
        hiddenData: 'show',
        categoryAxis: { id: 'category', position: 'bottom', visible: true, majorGridlines: { visible: false } },
        valueAxis: { id: 'value', position: 'left', visible: true, majorGridlines: { visible: true, color: '#e2e8f0', width: 1, dash: 'solid' } },
        chartArea: { fill: '#ffffff', border: '#cbd5e1', borderWidth: 1 },
        plotArea: { fill: '#ffffff' },
      },
    };
    this.addChart(drawing, payload);
  }
  updateChartType(chartId: string, chartType: ChartDrawingPayload['chartType']): void {
    this.runCommand('chart.setType', { sheetId: this.activeSheetId, chartId, chartType });
    this.refresh();
  }
  updateChartSeries(chartId: string, sourceRanges: RangeRef[], series?: ChartDrawingPayload['series'], categoryRange?: RangeRef): void {
    this.runCommand('chart.setSeries', {
      sheetId: this.activeSheetId,
      chartId,
      sourceRanges,
      series: series?.map((entry) => structuredClone(entry)),
      categoryRange,
    });
    this.refresh();
  }
  setChartLegend(chartId: string, legendPosition: NonNullable<ChartDrawingPayload['elements']['legend']>['position']): void {
    this.runCommand('chart.setLegend', { sheetId: this.activeSheetId, chartId, legendPosition });
    this.refresh();
  }
  setChartDataLabels(chartId: string, showDataLabels: boolean): void {
    this.runCommand('chart.setDataLabels', { sheetId: this.activeSheetId, chartId, showDataLabels });
    this.refresh();
  }
  setChartElements(chartId: string, elements: Partial<ChartDrawingPayload['elements']>): void {
    this.runCommand('chart.setElements', { sheetId: this.activeSheetId, chartId, elements });
    this.refresh();
  }
  setChartSeriesStyle(chartId: string, seriesName: string, style: Partial<NonNullable<ChartDrawingPayload['series']>[number]>): void {
    this.runCommand('chart.setSeriesStyle', { sheetId: this.activeSheetId, chartId, seriesName, style });
    this.refresh();
  }
  updateChartBounds(id: string, bounds: DrawingTransform): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const move = resolveDrawingMoveTransform(sheet, id, bounds);
    if (move) {
      this.runCommand('drawing.move', { sheetId: this.activeSheetId, drawingId: move.drawingId, transform: move.transform });
      this.refresh();
      return;
    }
    this.notify('Chart is not registered as a drawing object');
  }
  removeChart(id: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const drawing = findDrawingByPayloadId(sheet, id);
    if (!drawing) {
      this.notify('Chart is not registered as a drawing object');
      return;
    }
    this.runCommand('chart.remove', { sheetId: this.activeSheetId, chartId: id });
    this.refresh();
  }
  addPivot(pivot: PivotModel): void {
    this.runCommand('pivot.add', pivot);
    this.notify(`Pivot ${pivot.id} added`);
    this.refresh();
  }
  insertPivotFromSelection(): string | undefined {
    const descriptor = this.buildPivotFromSelectionDescriptor();
    const pivot = descriptor?.params as PivotModel | undefined;
    if (!pivot) {
      this.notify('Select a data range with category and value fields');
      return undefined;
    }
    this.addPivot(pivot);
    return pivot.id;
  }

  buildPivotFromSelectionDescriptor(): CommandDescriptor | undefined {
    const range = normalizeRangeRef(this.getCurrentRegion());
    const pivot = buildPivotModel(this.runtime.model, this.activeSheetId, nextId('pivot'), range);
    return pivot ? { commandId: 'pivot.add', params: pivot } : undefined;
  }

  createPivotTable(params: CreatePivotTableParams): string | undefined {
    const selectedRegion = normalizeRangeRef(this.getCurrentRegion());
    const sourceRegion = params.source?.kind === 'worksheet-range' ? normalizeRangeRef(params.source.range) : selectedRegion;
    if ((!params.source || params.source.kind === 'worksheet-range')
      && (sourceRegion.endRow <= sourceRegion.startRow || sourceRegion.endColumn < sourceRegion.startColumn)) {
      this.notify('Select a tabular source range with a header row before creating a PivotTable');
      return undefined;
    }
    const pivotId = nextId('pivot');
    let targetSheetId: string;
    let targetPosition: { row: number; column: number };
    let destination: {
      kind: 'new-sheet';
      sheetId: string;
      name: string;
    } | {
      kind: 'existing-sheet';
      sheetId: string;
    };
    if (params.destination.kind === 'new-sheet') {
      targetSheetId = nextId('sheet');
      targetPosition = { row: 0, column: 0 };
      const names = new Set(this.runtime.model.getSheets().map((sheet) => sheet.name.toLocaleLowerCase()));
      let suffix = 1;
      let targetName = 'PivotTable';
      while (names.has(targetName.toLocaleLowerCase())) targetName = `PivotTable${suffix++}`;
      destination = { kind: 'new-sheet', sheetId: targetSheetId, name: targetName };
    } else {
      targetSheetId = params.destination.sheetId;
      targetPosition = { ...params.destination.anchor };
      destination = { kind: 'existing-sheet', sheetId: targetSheetId };
    }

    try {
      const blockRegion = this.runtime.model.getSheet(sourceRegion.sheetId).dataRegions.find((region) => region.range.startRow === sourceRegion.startRow
        && region.range.endRow === sourceRegion.endRow && region.range.startColumn === sourceRegion.startColumn && region.range.endColumn === sourceRegion.endColumn);
      const source = params.source
        ?? (blockRegion
          ? { kind: 'data-source' as const, dataSourceId: blockRegion.sourceId }
          : { kind: 'worksheet-range' as const, range: { ...sourceRegion } });
      const pivotDraft: PivotModel = {
        schema: 'PivotDefinition',
        id: pivotId,
        source,
        target: { sheetId: targetSheetId, anchor: targetPosition },
        fieldCatalog: { fields: [] },
        refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
        layout: {
          rows: [],
          columns: [],
          filters: [],
          values: [],
          calculatedFields: [],
          calculatedItems: [],
          showSubtotals: true,
          showGrandTotals: true,
          compact: true,
          repeatLabels: false,
          expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
        },
      };
      const pivot: PivotModel = { ...pivotDraft, fieldCatalog: buildPivotFieldCatalog(this.runtime.model, pivotDraft) };
      this.runCommand('pivot.create', { pivot, destination });
      this.activeSheetId = targetSheetId;
      this.selectionService.resetForSheet(targetSheetId);
      this.setActivePivotContext(pivotId, targetSheetId);
      this.refresh();
      return pivotId;
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Could not create PivotTable');
      return undefined;
    }
  }
  updatePivotLayout(pivotId: string, layout: PivotLayout): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, layout });
  }
  updatePivotConfiguration(
    pivotId: string,
    patch: Parameters<WorkbookSession['updatePivotLayout']>[1] extends PivotLayout
      ? { source?: PivotDefinition['source']; target?: PivotDefinition['target']; fieldCatalog?: PivotDefinition['fieldCatalog']; refreshPolicy?: PivotDefinition['refreshPolicy']; nativeMetadata?: PivotDefinition['nativeMetadata']; layout?: PivotLayout }
      : never,
  ): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, ...patch });
  }
  setPivotAggregate(pivotId: string, fieldId: string, summarizeBy: PivotAggregateFunction): void {
    this.runCommand('pivot.setAggregate', { sheetId: this.activeSheetId, pivotId, fieldId, summarizeBy });
  }

  listPivotControls(pivotId: string): readonly PivotControlRecord[] {
    const pivot = this.runtime.model.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === pivotId);
    if (!pivot) return [];
    return listPivotControlsForPivot(this.runtime.model.getSheet(pivot.target.sheetId), pivotId)
      .map((record) => ({ drawing: structuredClone(record.drawing), payload: structuredClone(record.payload) }));
  }

  createPivotSlicerControl(pivotId: string, fieldId: string): void {
    const pivot = this.runtime.model.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === pivotId);
    if (!pivot) throw new Error(`Unknown PivotTable: ${pivotId}`);
    const sheet = this.runtime.model.getSheet(pivot.target.sheetId);
    const existing = listPivotControlsForPivot(sheet, pivotId).find((entry) => entry.payload.kind === 'slicer' && entry.payload.fieldId === fieldId);
    if (existing) return;
    const sourceRange = pivot.source.kind === 'worksheet-range' ? pivot.source.range : undefined;
    const connectedPivotIds = sourceRange ? connectedPivotIdsForSource(this.runtime.model, pivot.target.sheetId, sourceRange) : [pivotId];
    const offset = listPivotControlsForPivot(sheet, pivotId).length;
    const control = buildPivotSlicerDrawing({
      drawingId: nextId('pivot-slicer'),
      payloadId: nextId('pivot-slicer-payload'),
      sheetId: sheet.id,
      pivotId,
      fieldId,
      transform: { x: 96, y: 96 + offset * 144, width: 188, height: 128 },
      zIndex: sheet.drawings.length,
      connectedPivotIds,
    });
    this.runCommand('pivot.control.slicer.create', { sheetId: sheet.id, ...control });
    this.refresh();
  }

  createPivotTimelineControl(pivotId: string, fieldId: string): void {
    const pivot = this.runtime.model.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === pivotId);
    if (!pivot) throw new Error(`Unknown PivotTable: ${pivotId}`);
    const sheet = this.runtime.model.getSheet(pivot.target.sheetId);
    const existing = listPivotControlsForPivot(sheet, pivotId).find((entry) => entry.payload.kind === 'timeline' && entry.payload.fieldId === fieldId);
    if (existing) return;
    const sourceRange = pivot.source.kind === 'worksheet-range' ? pivot.source.range : undefined;
    const connectedPivotIds = sourceRange ? connectedPivotIdsForSource(this.runtime.model, pivot.target.sheetId, sourceRange) : [pivotId];
    const offset = listPivotControlsForPivot(sheet, pivotId).length;
    const control = buildPivotTimelineDrawing({
      drawingId: nextId('pivot-timeline'),
      payloadId: nextId('pivot-timeline-payload'),
      sheetId: sheet.id,
      pivotId,
      fieldId,
      transform: { x: 312, y: 96 + offset * 96, width: 356, height: 72 },
      zIndex: sheet.drawings.length,
      connectedPivotIds,
    });
    this.runCommand('pivot.control.timeline.create', { sheetId: sheet.id, ...control });
    this.refresh();
  }

  removePivotControl(drawingId: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
    if (!drawing) return;
    this.runCommand('drawing.remove', { sheetId: sheet.id, drawingId });
    this.refresh();
  }

  setPivotSlicerFilter(drawingId: string, mode: 'all' | 'include' | 'exclude', memberKeys: readonly PivotMemberKey[]): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    this.runCommand('pivot.control.slicer.filter.set', { sheetId: sheet.id, drawingId, filter: { mode, memberKeys: [...memberKeys] } });
    this.refresh();
  }

  setPivotTimelinePeriod(drawingId: string, start?: string, end?: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    this.runCommand('pivot.control.timeline.period.set', { sheetId: sheet.id, drawingId, period: { ...(start ? { start } : {}), ...(end ? { end } : {}) } });
    this.refresh();
  }
  refreshPivot(pivotId: string): void {
    const pivot = this.runtime.model.getSheet(this.activeSheetId).pivots.find((entry) => entry.id === pivotId);
    if (!pivot) return;
    this.runCommand('pivot.refresh', { sheetId: this.activeSheetId, pivotId });
  }
  removePivot(id: string): void {
    this.runCommand('pivot.remove', id);
    this.pivotTaskGeneration.delete(id);
    delete this.runtime.pivotResults[id];
    this.refresh();
  }
  drillDownPivot(pivotId: string, label: string, paths: readonly PivotSourceRowPath[]): void {
    if (paths.length === 0) return;
    const targetSheetId = nextId('sheet');
    this.runCommand('pivot.drillDown', {
      sheetId: this.activeSheetId,
      pivotId,
      label,
      sourceRowPaths: paths.map((path) => ({ sheetId: path.sheetId, row: path.row })),
      targetSheetId,
      target: { row: 0, column: 0 },
    });
    this.selectSheet(targetSheetId);
    this.notify(`Drill-down sheet created for ${label}`);
    this.refresh();
  }
  private recomputePivotResult(pivotId: string): void {
    const owner = this.runtime.model.getSheets().find((sheet) => sheet.pivots.some((entry) => entry.id === pivotId));
    const pivot = owner?.pivots.find((entry) => entry.id === pivotId);
    if (!pivot || !owner) {
      delete this.runtime.pivotResults[pivotId];
      return;
    }
    if (pivot.source.kind === 'data-source') {
      const sourceId = pivot.source.dataSourceId;
      const query = this.runtime.dataContent.get(sourceId);
      const region = this.runtime.model.getSheets()
        .flatMap((sheet) => sheet.dataRegions.map((entry) => ({ sheet, entry })))
        .find(({ entry }) => entry.sourceId === sourceId);
      if (!query || !region) {
        delete this.runtime.pivotResults[pivotId];
        this.notify(`PivotTable source ${sourceId} is unavailable`);
        return;
      }
      const taskRevision = (this.pivotTaskGeneration.get(pivotId) ?? 0) + 1;
      this.pivotTaskGeneration.set(pivotId, taskRevision);
      void readPivotBlockSource(normalizePivotDefinition(this.runtime.model, pivot), query, {
        sourceSheetId: region.sheet.id,
        sourceRowStart: region.entry.headerRow + 1,
      }).then((result) => {
        if (this.pivotTaskGeneration.get(pivotId) !== taskRevision || result.status !== 'ready') {
          if (result.status !== 'ready') this.notify(result.error);
          return;
        }
        this.runtime.pivotResults[pivotId] = computePivotResultFromBlockSource(
          this.runtime.model,
          pivot,
          result.source,
          `${result.state.sourceId}:${result.sourceRevision}`,
        );
        this.refresh();
      }).catch((error) => {
        if (this.pivotTaskGeneration.get(pivotId) === taskRevision) this.notify(error instanceof Error ? error.message : 'PivotTable source failed to load');
      });
      return;
    }
    try {
      this.runtime.pivotResults[pivotId] = computePivotResult(this.runtime.model, pivot);
    } catch {
      delete this.runtime.pivotResults[pivotId];
    }
  }
  private recomputeAllPivotResults(): void {
    const pivots = this.runtime.model.getSheets().flatMap((sheet) => sheet.pivots);
    const activeIds = new Set(pivots.map((pivot) => pivot.id));
    for (const pivotId of Object.keys(this.runtime.pivotResults)) if (!activeIds.has(pivotId)) delete this.runtime.pivotResults[pivotId];
    for (const pivot of pivots) this.recomputePivotResult(pivot.id);
  }
  addShape(drawing: DrawingObject, payload: ShapeDrawingPayload): void {
    if (drawing.kind !== 'shape' || payload.kind !== 'shape') {
      throw new Error(`Shape drawing and payload kind mismatch: ${drawing.id}`);
    }
    this.runCommand('drawing.add.shape', buildDrawingAdd(drawing, payload));
    this.setDrawingSelection([drawing.id]);
    this.notify(`Added ${payload.type} shape`);
    this.refresh();
  }
  insertShape(type: ShapeDrawingPayload['type'] = 'rounded-rectangle'): void {
    const payloadId = nextId('shape');
    const drawing: DrawingObject = {
      id: nextId('draw'),
      sheetId: this.activeSheetId,
      kind: 'shape',
      anchor: { kind: 'absolute' },
      transform: { x: 96, y: 96, width: 160, height: 60, rotation: 0 },
      zIndex: 0,
      payloadId,
    };
    const payload: ShapeDrawingPayload = {
      kind: 'shape',
      type,
      text: type === 'callout' ? 'Note' : '',
      fill: '#dbeafe',
      stroke: '#2563eb',
      strokeWidth: 2,
      textColor: '#1e3a8a',
      fontSize: 13,
    };
    this.addShape(drawing, payload);
  }
  updateShapeBounds(id: string, bounds: DrawingTransform): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const move = resolveDrawingMoveTransform(sheet, id, bounds);
    if (move) {
      this.runCommand('drawing.move', { sheetId: this.activeSheetId, drawingId: move.drawingId, transform: move.transform });
      this.refresh();
      return;
    }
    this.notify('Shape is not registered as a drawing object');
  }
  removeShape(id: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const drawing = findDrawingByPayloadId(sheet, id);
    if (drawing) {
      this.runCommand('drawing.remove', { sheetId: this.activeSheetId, drawingId: drawing.id });
    } else {
      this.notify('Shape is not registered as a drawing object');
      return;
    }
    this.refresh();
  }
  addImage(drawing: DrawingObject, payload: ImageDrawingPayload): void {
    if (drawing.kind !== 'image' || payload.kind !== 'image') {
      throw new Error(`Image drawing and payload kind mismatch: ${drawing.id}`);
    }
    this.runCommand('drawing.add.image', buildDrawingAdd(drawing, payload));
    this.setDrawingSelection([drawing.id]);
    this.notify('Image placed on canvas');
    this.refresh();
  }
  async insertImageFile(file: File, placement: 'cell' | 'floating' = 'floating'): Promise<void> {
    if (!file.type.startsWith('image/')) throw new Error('请选择图片文件');
    const src = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(reader.error ?? new Error('图片读取失败'));
      reader.onload = () => typeof reader.result === 'string' ? resolve(reader.result) : reject(new Error('图片读取失败'));
      reader.readAsDataURL(file);
    });
    if (placement === 'cell') {
      const active = this.selectionService.getState().activeCell;
      this.runCommand('cell.image.apply', { sheetId: this.activeSheetId, row: active.row, column: active.column, presentation: { kind: 'image', src, altText: file.name, fit: 'contain' } });
      this.notify('图片已嵌入单元格');
      this.refresh();
      return;
    }
    const payloadId = nextId('image');
    this.addImage({ id: nextId('drawing'), sheetId: this.activeSheetId, kind: 'image', anchor: { kind: 'absolute' }, transform: { x: 96, y: 96, width: 320, height: 200, rotation: 0 }, zIndex: 0, payloadId }, { kind: 'image', src, name: file.name, altText: file.name });
  }
  updateImageBounds(id: string, bounds: DrawingTransform): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const move = resolveDrawingMoveTransform(sheet, id, bounds);
    if (move) {
      this.runCommand('drawing.move', { sheetId: this.activeSheetId, drawingId: move.drawingId, transform: move.transform });
      this.refresh();
      return;
    }
    this.notify('Image is not registered as a drawing object');
  }
  removeImage(id: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const drawing = findDrawingByPayloadId(sheet, id);
    if (drawing) {
      this.runCommand('drawing.remove', { sheetId: this.activeSheetId, drawingId: drawing.id });
    } else {
      this.notify('Image is not registered as a drawing object');
      return;
    }
    this.refresh();
  }
  bringSelectedDrawingForward(): void {
    const drawingId = this.resolveSelectedDrawingId();
    if (!drawingId) {
      this.notify('Select a drawing object first');
      return;
    }
    this.runCommand('drawing.zorder', { sheetId: this.activeSheetId, drawingId, direction: 'forward' });
    this.refresh();
  }
  sendSelectedDrawingBackward(): void {
    const drawingId = this.resolveSelectedDrawingId();
    if (!drawingId) {
      this.notify('Select a drawing object first');
      return;
    }
    this.runCommand('drawing.zorder', { sheetId: this.activeSheetId, drawingId, direction: 'backward' });
    this.refresh();
  }
  bringSelectedDrawingToFront(): void {
    const drawingId = this.resolveSelectedDrawingId();
    if (!drawingId) {
      this.notify('Select a drawing object first');
      return;
    }
    this.runCommand('drawing.zorder', { sheetId: this.activeSheetId, drawingId, direction: 'front' });
    this.refresh();
  }
  sendSelectedDrawingToBack(): void {
    const drawingId = this.resolveSelectedDrawingId();
    if (!drawingId) {
      this.notify('Select a drawing object first');
      return;
    }
    this.runCommand('drawing.zorder', { sheetId: this.activeSheetId, drawingId, direction: 'back' });
    this.refresh();
  }
  alignSelectedDrawings(alignment: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'): void {
    const drawingIds = this.resolveSelectedDrawingIds();
    if (drawingIds.length < 2) {
      this.notify('Select at least two drawing objects to align');
      return;
    }
    this.runCommand('drawing.align', { sheetId: this.activeSheetId, drawingIds, alignment });
    this.refresh();
  }
  distributeSelectedDrawings(axis: 'horizontal' | 'vertical'): void {
    const drawingIds = this.resolveSelectedDrawingIds();
    if (drawingIds.length < 3) {
      this.notify('Select at least three drawing objects to distribute');
      return;
    }
    this.runCommand('drawing.distribute', { sheetId: this.activeSheetId, drawingIds, axis });
    this.refresh();
  }
  removeSelectedDrawing(): void {
    if (!this.selectedFloatingId) {
      this.notify('Select a drawing object first');
      return;
    }
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const drawing = sheet.drawings.find((entry) => entry.id === this.selectedFloatingId);
    if (!drawing) {
      this.notify('Selected object is not registered as a drawing');
      return;
    }
    this.runCommand('drawing.remove', { sheetId: this.activeSheetId, drawingId: drawing.id });
    this.refresh();
  }
  private resolveSelectedDrawingId(): string | undefined {
    if (!this.selectedFloatingId) return undefined;
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    return sheet.drawings.some((entry) => entry.id === this.selectedFloatingId) ? this.selectedFloatingId : undefined;
  }
  private resolveSelectedDrawingIds(): string[] {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selected = this.runtime.drawing.getSelection(this.activeSheetId)
      .filter((id) => sheet.drawings.some((entry) => entry.id === id));
    if (selected.length > 0) return selected;
    const single = this.resolveSelectedDrawingId();
    return single ? [single] : [];
  }
  addSparkline(sparkline: SparklineModel): void {
    this.runCommand('sparkline.insert', buildSparklineInsertParams(sparkline));
    this.notify(`Sparkline inserted at row ${sparkline.anchor.row + 1}`);
    this.refresh();
  }
  insertSparklineDataLocation(
    params: {
      sparklineId: string;
      dataRange: RangeRef;
      location: { row: number; column: number };
      type?: SparklineModel['type'];
    } & Partial<Pick<SparklineModel, 'color' | 'negativeColor' | 'highlightMax' | 'highlightMin' | 'highlightFirst' | 'highlightLast' | 'highlightNegative' | 'groupId' | 'showAxis' | 'showMarkers'>>,
  ): string {
    const sparklineId = params.sparklineId;
    this.runCommand(
      'sparkline.insertDataLocation',
      buildSparklineDataLocationParams(
        this.activeSheetId,
        sparklineId,
        params.dataRange,
        params.location,
        params.type ?? 'line',
        params,
      ),
    );
    this.notify(`Sparkline inserted at row ${params.location.row + 1}`);
    this.refresh();
    return sparklineId;
  }
  insertSparkline(type: SparklineModel['type'] = 'line'): string | undefined {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endColumn <= range.startColumn && range.endRow <= range.startRow) {
      this.notify('Select a data range for the sparkline source');
      return undefined;
    }
    const placement = resolveQuickSparklinePlacement({ ...range, sheetId: this.activeSheetId });
    const sparklineId = nextId('spark');
    this.insertSparklineDataLocation({
      sparklineId,
      dataRange: placement.dataRange,
      location: placement.location,
      type,
      highlightMax: true,
      highlightMin: true,
    });
    return sparklineId;
  }
  updateSparkline(sparklineId: string, patch: Partial<SparklineModel>): void {
    this.runCommand('sparkline.update', { sheetId: this.activeSheetId, sparklineId, patch });
    this.refresh();
  }
  createSparklineGroup(sparklineIds: string[], patch?: Partial<Pick<SparklineGroup, 'showAxis' | 'showMarkers'>>, type: SparklineModel['type'] = 'line'): string {
    const groupId = nextId('sparkline-group');
    const group = buildSparklineGroup(this.activeSheetId, groupId, sparklineIds, type, patch);
    this.runCommand('sparkline.group.create', { sheetId: this.activeSheetId, group });
    this.notify(`Sparkline group created (${sparklineIds.length})`);
    this.refresh();
    return groupId;
  }
  updateSparklineGroup(groupId: string, patch: Partial<SparklineGroup>): void {
    this.runCommand('sparkline.group.update', { sheetId: this.activeSheetId, groupId, patch });
    this.refresh();
  }
  removeSparkline(id: string): void {
    this.runCommand('sparkline.remove', { sheetId: this.activeSheetId, sparklineId: id });
    this.refresh();
  }
  addConditionalFormat(rule: ConditionalFormatRule): void {
    this.dispatch({ commandId: 'sheet.cf.add', params: { sheetId: this.activeSheetId, rule } });
  }
  removeConditionalFormat(ruleId: string): void {
    this.dispatch({ commandId: 'sheet.cf.remove', params: { sheetId: this.activeSheetId, ruleId } });
  }
  addDataValidation(rule: DataValidationRule): void {
    this.dispatch({ commandId: 'sheet.dv.add', params: { sheetId: this.activeSheetId, rule } });
  }
  removeDataValidation(ruleId: string): void {
    this.dispatch({ commandId: 'sheet.dv.remove', params: { sheetId: this.activeSheetId, ruleId } });
  }
  setCellStyleTemplate(template: CellStyleTemplate): void {
    this.dispatch({ commandId: 'workbook.cellTemplate.set', params: { sheetId: this.activeSheetId, template } });
  }
  removeCellStyleTemplate(templateId: string): void {
    this.dispatch({ commandId: 'workbook.cellTemplate.remove', params: { sheetId: this.activeSheetId, templateId } });
  }
  applyCellStyleTemplate(templateId: string): void {
    const ranges = this.selectionService.getState().ranges.map((range) => ({ ...range, sheetId: this.activeSheetId }));
    this.dispatch({ commandId: 'sheet.cellTemplate.apply', params: { sheetId: this.activeSheetId, ranges, templateId } });
  }
  setCellEditor(editor?: CellEditorConfig): void {
    const ranges = this.selectionService.getState().ranges.map((range) => ({ ...range, sheetId: this.activeSheetId }));
    this.dispatch({ commandId: 'sheet.cellEditor.set', params: { sheetId: this.activeSheetId, ranges, editor } });
  }

  addComment(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const thread = buildCommentThread(
      this.activeSheetId,
      sel.activeCell.row,
      sel.activeCell.column,
      this.actorId,
      text,
      nextId('thread'),
    );
    this.runCommand('comment.add', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      thread,
    });
    this.notify('Comment added');
    this.refresh();
  }

  applyFilter(column: number, patch: { criterion?: FilterCriterion }): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const activeFilter = resolveActiveAutoFilter(sheet, column);
    const owner = resolveFilterOwner(sheet, column);
    const tableOwner = owner?.kind === 'table'
      ? sheet.sheetTables.find((table) => table.id === owner.tableId)
      : undefined;
    const baseRange =
      activeFilter?.range ?? tableOwner?.range ??
      this.getCurrentRegion();
    const columns = { ...(activeFilter?.columns ?? {}) };
    for (let current = baseRange.startColumn; current <= baseRange.endColumn; current += 1) {
      columns[current] ??= { column: current, showButton: true, hiddenButton: false };
    }
    columns[column] = { ...(columns[column] ?? { column, showButton: true, hiddenButton: false }), criterion: patch.criterion ? structuredClone(patch.criterion) : undefined };
    const autoFilter = { sheetId: this.activeSheetId, range: baseRange, columns };
    if (owner?.kind === 'table' && tableOwner) {
      this.dispatch({ commandId: 'sheetTable.autoFilter.set', params: { sheetId: this.activeSheetId, tableId: tableOwner.id, autoFilter } });
    } else {
      this.dispatch({ commandId: 'sheet.autoFilter.set', params: { sheetId: this.activeSheetId, autoFilter } });
    }
  }

  sortFilterColumn(column: number, ascending: boolean): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const filter = resolveActiveAutoFilter(sheet, column);
    if (!filter || column < filter.range.startColumn || column > filter.range.endColumn) {
      this.notify('No active filter is available for this column');
      return;
    }
    this.dispatch({ commandId: 'sheet.autoFilter.sort', params: { sheetId: this.activeSheetId, column, ascending } });
  }

  applyFilterSelection(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const activeFilter = resolveActiveAutoFilter(sheet);
    const owner = resolveFilterOwner(sheet);
    if (activeFilter && owner) {
      if (owner.kind === 'table') {
        const table = sheet.sheetTables.find((entry) => entry.id === owner.tableId);
        if (table) this.dispatch({ commandId: 'sheetTable.update', params: { ...structuredClone(table), showFilterButton: false, autoFilter: undefined } });
      }
      else this.dispatch({ commandId: 'sheet.autoFilter.toggle', params: { sheetId: this.activeSheetId, range: this.getCurrentRegion() } });
      return;
    }
    const range = this.getCurrentRegion();
    if (range.endRow <= range.startRow) {
      this.notify('Select a data region with a header row before enabling Filter');
      return;
    }
    this.dispatch({ commandId: 'sheet.autoFilter.toggle', params: {
      sheetId: this.activeSheetId,
      range,
    } });
  }

  clearFilter(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const autoFilter = resolveActiveAutoFilter(sheet);
    const owner = resolveFilterOwner(sheet);
    if (!autoFilter || !owner) {
      this.notify('No filter is active in the current region');
      return;
    }
    if (!Object.values(autoFilter.columns).some((column) => Boolean(column.criterion))) {
      this.notify('No filter criteria are active in the current region');
      return;
    }
    const columns = Object.fromEntries(Object.entries(autoFilter.columns).map(([key, value]) => [key, { ...value, criterion: undefined }]));
    if (owner.kind === 'table') this.dispatch({ commandId: 'sheetTable.autoFilter.set', params: { sheetId: this.activeSheetId, tableId: owner.tableId, autoFilter: { ...autoFilter, columns } } });
    else this.dispatch({ commandId: 'sheet.autoFilter.clearCriteria', params: { sheetId: this.activeSheetId, range: autoFilter.range } });
  }

  closeFilter(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const autoFilter = resolveActiveAutoFilter(sheet);
    const owner = resolveFilterOwner(sheet);
    if (!autoFilter || !owner) return;
    if (owner.kind === 'table') {
      const table = sheet.sheetTables.find((entry) => entry.id === owner.tableId);
      if (table) this.dispatch({ commandId: 'sheetTable.update', params: { ...structuredClone(table), showFilterButton: false, autoFilter: undefined } });
    }
    else this.dispatch({ commandId: 'sheet.autoFilter.toggle', params: { sheetId: this.activeSheetId, range: autoFilter.range } });
  }

  async findReplace(params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }): Promise<number> {
    if (!params.find) return 0;
    const command = {
      sheetId: this.activeSheetId,
      range: this.getCurrentRegion(),
      find: params.find,
      replace: params.replace,
      matchCase: params.matchCase,
      entireCell: params.entireCell,
      scope: params.scope,
      searchIn: 'both' as const,
    };
    const result = await this.executeCommandAfterMaterialization('sheet.range.replace', command);
    this.notify(`${result.mutationCount} replacement(s) applied`);
    return result.mutationCount;
  }

  insertRowsAtPrimary(count: number): void {
    this.dispatch({ commandId: 'sheet.rows.insert', params: { sheetId: this.activeSheetId, at: this.selectionService.getState().activeCell.row, count } });
  }
  deleteRowsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.dispatch({ commandId: 'sheet.rows.delete', params: { sheetId: this.activeSheetId, at: range?.startRow ?? sel.activeCell.row, count: (range?.endRow ?? sel.activeCell.row) - (range?.startRow ?? sel.activeCell.row) + 1 } });
  }
  insertColumnsAtPrimary(count: number): void {
    this.dispatch({ commandId: 'sheet.columns.insert', params: { sheetId: this.activeSheetId, at: this.selectionService.getState().activeCell.column, count } });
  }
  deleteColumnsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.dispatch({ commandId: 'sheet.columns.delete', params: { sheetId: this.activeSheetId, at: range?.startColumn ?? sel.activeCell.column, count: (range?.endColumn ?? sel.activeCell.column) - (range?.startColumn ?? sel.activeCell.column) + 1 } });
  }
  hideRowsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    const start = range?.startRow ?? sel.activeCell.row;
    const end = range?.endRow ?? sel.activeCell.row;
    for (let row = start; row <= end; row++) this.runCommand('sheet.row.hide', { sheetId: this.activeSheetId, index: row });
  }
  hideColumnsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    const start = range?.startColumn ?? sel.activeCell.column;
    const end = range?.endColumn ?? sel.activeCell.column;
    for (let column = start; column <= end; column++) this.runCommand('sheet.column.hide', { sheetId: this.activeSheetId, index: column });
  }
  unhideAll(): void {
    this.runCommand('sheet.rows.unhide.all', { sheetId: this.activeSheetId });
    this.runCommand('sheet.columns.unhide.all', { sheetId: this.activeSheetId });
  }
  toggleBandedRows(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const next = sheet.bandedRule
      ? null
      : { range: normalizeRangeRef({ sheetId: this.activeSheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }), firstColor: '#ffffff', secondColor: '#f1f5f9' };
    this.runCommand('sheet.banded.set', { sheetId: this.activeSheetId, rule: next });
  }
  transposeSelection(): void {
    this.runCommand('matrix.transpose', { sheetId: this.activeSheetId, range: this.getPrimaryRange() });
  }
  flipSelection(axis: 'h' | 'v'): void {
    this.runCommand('matrix.flip', { sheetId: this.activeSheetId, range: this.getPrimaryRange(), direction: axis === 'h' ? 'horizontal' : 'vertical' });
  }
  async splitByDelimiter(delimiter: string): Promise<void> {
    const sel = this.selectionService.getState();
    const sheet = this.getSelectedSheet();
    await this.executeCommandAfterMaterialization('data.splitColumn', { sheetId: this.activeSheetId, row: sel.activeCell.row, column: sel.activeCell.column, delimiter, maxColumns: Math.min(sheet.columnCount - sel.activeCell.column - 1, 8) });
  }

  copy(): Promise<ClipboardExecutionOutcome> {
    return this.copyOrCut(false);
  }
  cut(): Promise<ClipboardExecutionOutcome> {
    return this.copyOrCut(true);
  }

  private async copyOrCut(move: boolean): Promise<ClipboardExecutionOutcome> {
    const range = this.getPrimaryRange();
    try {
      await this.materializeDataRegions(this.dataRegionsIntersectingRanges(range.sheetId, [range]));
    } catch (error) {
      const normalized = error instanceof Error ? error : new Error('Clipboard source could not be prepared');
      this.notify(normalized.message);
      return { status: 'failed', formats: [], error: normalized, privatePayloadStored: false };
    }
    const data = copyRangeToClipboardData(this.runtime.model, range);
    this.setClipboard({ ...data, transfer: move ? 'move' : 'copy' });
    this.clipboardSystemStatus = 'unknown';
    this.clipboardSystemFormats = [];
    const system = await writeSystemClipboard(data);
    this.clipboardSystemStatus = system.status;
    this.clipboardSystemFormats = [...system.formats];
    const outcome: ClipboardExecutionOutcome = { ...system, privatePayloadStored: true };
    if (outcome.status === 'published') {
      this.notify(move ? 'Cut to clipboard' : 'Range copied');
    } else if (outcome.status === 'reduced') {
      const formats = outcome.formats.join(', ');
      this.notify(`${move ? 'Cut' : 'Range copied'} internally; external clipboard formats: ${formats}`);
    } else {
      this.notify(`${move ? 'Cut' : 'Copy'} retained in this workbook; external clipboard publication failed: ${outcome.error.message}`);
    }
    return outcome;
  }
  async paste(spec: PasteSpecialSpec = createPasteSpecialSpec()): Promise<DispatchOutcome> {
    const sel = this.selectionService.getState();
    const internal = this.clipboardData;
    if (internal) {
      const outcome = await this.dispatch({ commandId: 'sheet.range.paste', params: {
        sheetId: this.activeSheetId,
        targetOrigin: { row: sel.activeCell.row, column: sel.activeCell.column },
        clipboard: internal,
        transfer: internal.transfer,
        spec,
      } });
      if (outcome.status !== 'committed') return outcome;
      if (outcome.result.mutationCount === 0) {
        this.notify('Paste made no changes');
        return outcome;
      }
      if (internal.transfer === 'move') this.clearClipboard();
      this.syncDraftFromPrimary();
      this.notify('Pasted from clipboard');
      return outcome;
    }
    if (typeof navigator === 'undefined' || !navigator.clipboard) {
      return this.rejectDispatch(new CommandDispatchError('COMMAND_REJECTED', 'No clipboard data is available'));
    }
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return this.rejectDispatch(new CommandDispatchError('COMMAND_REJECTED', 'Clipboard permission was denied'));
    }
    if (!text) return this.rejectDispatch(new CommandDispatchError('COMMAND_REJECTED', 'Clipboard is empty'));
    const clipboard: ClipboardPayload = {
      range: this.getPrimaryRange(),
      values: parseTsv(text),
      transfer: 'copy',
      rangeMetadata: {
        columnWidths: [],
        validations: [],
        conditionalFormats: [],
        notes: [],
        comments: [],
        hyperlinks: [],
      },
    };
    const outcome = await this.dispatch({ commandId: 'sheet.range.paste', params: {
        sheetId: this.activeSheetId,
        targetOrigin: { row: sel.activeCell.row, column: sel.activeCell.column },
        clipboard,
        transfer: 'copy',
        spec,
      } });
    if (outcome.status !== 'committed') return outcome;
    if (outcome.result.mutationCount === 0) {
      this.notify('Paste made no changes');
      return outcome;
    }
    this.syncDraftFromPrimary();
    this.notify('Pasted from clipboard');
    return outcome;
  }
  clearFormats(): void {
    this.dispatch({ commandId: 'sheet.range.clear', params: { sheetId: this.activeSheetId, range: this.getPrimaryRange(), mode: 'formats' } });
  }

  clearSelection(mode: 'contents' | 'formats' = 'contents'): void {
    this.dispatch({ commandId: 'sheet.range.clear', params: { sheetId: this.activeSheetId, range: this.getPrimaryRange(), mode } });
    this.syncDraftFromPrimary();
  }

  private getRangeMatrix(range: RangeRef): CellData[][] {
    const sheet = this.runtime.model.getSheet(range.sheetId);
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const rowValues: CellData[] = [];
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = this.readResolvedCell(sheet, r, c);
        rowValues.push(cell ? structuredClone(cell) : { value: null });
      }
      rows.push(rowValues);
    }
    return rows;
  }

  addSheet(): void {
    const id = 'sheet-' + Math.random().toString(36).slice(2, 8);
    this.runCommand('sheet.add', { id, name: 'Sheet' + (this.runtime.model.getSheets().length + 1) });
    this.selectSheet(id);
    this.refresh();
  }
  renameSheet(sheetId: string, name: string): void {
    if (!name.trim()) return;
    this.runCommand('sheet.rename', { sheetId, name: name.trim() });
  }
  renameWorkbook(name: string): void {
    if (!name.trim()) return;
    this.runCommand('workbook.rename', { name });
  }
  duplicateSheet(sheetId: string): void {
    const source = this.runtime.model.getSheet(sheetId);
    const newId = 'sheet-' + Math.random().toString(36).slice(2, 8);
    const newName = `${source.name} (2)`;
    this.runCommand('sheet.duplicate', { sourceSheetId: sheetId, newId, newName });
    this.selectSheet(newId);
    this.syncDraftFromPrimary();
    this.refresh();
  }
  hideSheet(sheetId: string): void {
    try {
      this.runCommand('sheet.hide', { sheetId });
      if (this.activeSheetId === sheetId) {
        const next = this.runtime.model.getVisibleSheets()[0];
        if (next) this.selectSheet(next.id);
      }
      this.refresh();
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Cannot hide sheet');
    }
  }
  setSheetTabColor(sheetId: string, color?: string): void {
    this.runCommand('sheet.tabColor.set', { sheetId, color: color || undefined });
  }
  moveSheet(sheetId: string, toIndex: number): void {
    this.runCommand('sheet.reorder', { sheetId, toIndex });
    this.refresh();
  }
  deleteSheet(sheetId: string): void {
    try {
      this.runCommand('sheet.remove', { id: sheetId });
      if (this.activeSheetId === sheetId) {
        const remaining = this.runtime.model.getSheets()[0];
        if (remaining) this.selectSheet(remaining.id);
      }
      this.refresh();
    } catch {
      this.notify('A workbook must keep at least one sheet');
    }
  }
  resizeRow(row: number, heightPx: number): void {
    this.runCommand('sheet.dimensions.apply', { sheetId: this.activeSheetId, rows: [{ row, heightPx: Math.max(1, heightPx) }] });
  }
  resizeColumn(column: number, widthPx: number): void {
    this.runCommand('sheet.dimensions.apply', { sheetId: this.activeSheetId, columns: [{ column, widthPx: Math.max(1, widthPx) }] });
  }
  resizeColumns(columns: readonly number[], widthPx: number): void {
    const unique = [...new Set(columns)].filter((column) => Number.isSafeInteger(column) && column >= 0);
    if (!unique.length) return;
    this.runCommand('sheet.dimensions.apply', { sheetId: this.activeSheetId, columns: unique.map((column) => ({ column, widthPx: Math.max(1, widthPx) })) });
  }
  applyColumnWidths(entries: readonly { column: number; widthPx: number }[]): void {
    const unique = new Map<number, number>();
    for (const entry of entries) if (Number.isSafeInteger(entry.column) && entry.column >= 0 && Number.isFinite(entry.widthPx) && entry.widthPx > 0) unique.set(entry.column, entry.widthPx);
    if (!unique.size) return;
    this.runCommand('sheet.dimensions.apply', { sheetId: this.activeSheetId, columns: [...unique].map(([column, widthPx]) => ({ column, widthPx })) });
  }
  applyRowHeights(entries: readonly { row: number; heightPx: number }[]): void {
    const unique = new Map<number, number>();
    for (const entry of entries) if (Number.isSafeInteger(entry.row) && entry.row >= 0 && Number.isFinite(entry.heightPx) && entry.heightPx > 0) unique.set(entry.row, entry.heightPx);
    if (!unique.size) return;
    this.runCommand('sheet.dimensions.apply', { sheetId: this.activeSheetId, rows: [...unique].map(([row, heightPx]) => ({ row, heightPx })) });
  }
  setColumnsHidden(columns: readonly number[], hidden: boolean): void {
    const unique = [...new Set(columns)].filter((column) => Number.isSafeInteger(column) && column >= 0);
    if (!unique.length) return;
    this.runCommand('sheet.columns.visibility.set', { sheetId: this.activeSheetId, columns: unique, hidden });
  }
  setDefaultColumnWidth(widthPx: number): void {
    this.runCommand('sheet.column.defaultWidth.set', { sheetId: this.activeSheetId, widthPx: Math.max(1, widthPx) });
  }
  fillRange(targetRange: { startRow: number; endRow: number; startColumn: number; endColumn: number }): void {
    const sel = this.selectionService.getState();
    const primary = sel.ranges[sel.primaryRangeIndex] ?? sel.ranges[0];
    if (!primary) return;
    this.dispatch({ commandId: 'sheet.range.fill', params: {
      sheetId: this.activeSheetId,
      sourceRange: { sheetId: this.activeSheetId, startRow: primary.startRow, endRow: primary.endRow, startColumn: primary.startColumn, endColumn: primary.endColumn },
      targetRange: { sheetId: this.activeSheetId, ...targetRange },
    } });
  }

  fillSelection(direction: 'down' | 'up' | 'right' | 'left', mode: 'copy' | 'series' = 'copy'): void {
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    const sourceRange = direction === 'down'
      ? { ...range, endRow: range.startRow }
      : direction === 'up'
        ? { ...range, startRow: range.endRow }
        : direction === 'right'
          ? { ...range, endColumn: range.startColumn }
          : { ...range, startColumn: range.endColumn };
    if (sourceRange.startRow === range.startRow && sourceRange.endRow === range.endRow
      && sourceRange.startColumn === range.startColumn && sourceRange.endColumn === range.endColumn) {
      this.notify('Select a source cell and target cells before filling');
      return;
    }
    this.dispatch({ commandId: 'sheet.range.fill', params: {
      sheetId: this.activeSheetId,
      sourceRange,
      targetRange: range,
      direction,
      mode,
    } });
  }
  setSelectedFloatingId(id: string | null): void {
    this.setDrawingSelection(id ? [id] : [], 'replace');
  }

  setDrawingSelectionMode(enabled: boolean): void {
    if (this.drawingSelectionMode === enabled) return;
    this.drawingSelectionMode = enabled;
    if (!enabled) this.setDrawingSelection([], 'replace');
    else this.emit();
  }

  setDrawingSelection(ids: readonly string[], mode: 'replace' | 'add' | 'toggle' = 'replace'): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const valid = ids.filter((id) => sheet.drawings.some((drawing) => drawing.id === id));
    if (valid.length === 0 && mode === 'replace') {
      this.runCommand('drawing.deselect', { sheetId: this.activeSheetId });
      this.activeContext = { kind: 'none' };
      if (this.panels.active === 'picture') this.panels = { ...this.panels, open: false };
      if (this.ribbonTab === 'pictureFormat') this.ribbonTab = 'home';
      this.emit();
      return;
    }
    this.runCommand('drawing.select', { sheetId: this.activeSheetId, drawingIds: valid, mode });
    const selectedDrawing = this.selectedFloatingId ? sheet.drawings.find((drawing) => drawing.id === this.selectedFloatingId) : undefined;
    this.activeContext = selectedDrawing
      ? { kind: 'drawing', sheetId: this.activeSheetId, drawingId: selectedDrawing.id }
      : { kind: 'none' };
    if (selectedDrawing?.kind === 'image') {
      this.panels = { ...this.panels, active: 'picture', open: true };
      this.ribbonTab = 'pictureFormat';
    }
    this.emit();
  }

  setDrawingVisibility(drawingId: string, visible: boolean): void {
    this.runCommand('drawing.visibility.set', { sheetId: this.activeSheetId, drawingId, visible });
  }

  renameDrawing(drawingId: string, name: string): void {
    this.runCommand('drawing.rename', { sheetId: this.activeSheetId, drawingId, name });
  }

  removeFloatingObject(kind: 'chart' | 'shape' | 'image', id: string): void {
    if (kind === 'chart') this.removeChart(id);
    else if (kind === 'image') this.removeImage(id);
    else this.removeShape(id);
  }

  getConnectedPivotIds(sourceRange: RangeRef): string[] {
    return connectedPivotIdsForSource(this.runtime.model, this.activeSheetId, sourceRange);
  }
  getPivotFieldCatalog(range: RangeRef): PivotFieldDefinition[] {
    const pivot: PivotModel = {
      id: 'pivot-field-catalog',
      schema: 'PivotDefinition',
      source: { kind: 'worksheet-range', range: { ...range } },
      target: { sheetId: range.sheetId, anchor: { row: 0, column: 0 } },
      fieldCatalog: { fields: [] },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: { rows: [], columns: [], filters: [], values: [], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false, calculatedFields: [], calculatedItems: [], expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true } },
    };
    return buildPivotFieldCatalog(this.runtime.model, pivot).fields;
  }

  readDataTable(tableId: string, offset = 0, limit = 100): Promise<TableRowsResponse> {
    const table = this.runtime.model.dataModel.tables.get(tableId);
    if (!table || !table.sourceSheetId || !table.sourceRange) return Promise.reject(new Error('Data table not found'));
    const sheet = this.runtime.model.getSheet(table.sourceSheetId);
    const start = Math.max(0, offset);
    const end = Math.min(table.rowCount, start + Math.max(1, limit));
    const rows: import('@react-sheets/core-model').TableScalar[][] = [];
    for (let rowOffset = start; rowOffset < end; rowOffset += 1) {
      rows.push(table.fields.map((field) => this.readResolvedCell(sheet,
        table.sourceRange!.startRow + 1 + rowOffset,
        table.sourceRange!.startColumn + field.ordinal,
      )?.value ?? null));
    }
    return Promise.resolve({
      table: structuredClone(table),
      rows,
      ...(end < table.rowCount ? { nextOffset: end } : {}),
    });
  }
  removeDataTable(tableId: string): Promise<void> {
    const table = this.runtime.model.dataModel.tables.get(tableId);
    if (!table) return Promise.reject(new Error('Data table not found'));
    this.runCommand('table.remove', { tableId, sheetId: table.sourceSheetId ?? this.activeSheetId });
    return Promise.resolve();
  }

  showPivotDetails(pivotId: string, paths: readonly PivotSourceRowPath[], label = 'Details'): void {
    this.drillDownPivot(pivotId, label, paths);
  }

  getValidationForPrimary(): DataValidationRule | undefined {
    const sel = this.selectionService.getState();
    return findValidationRule(this.runtime.model.getSheet(this.activeSheetId), sel.activeCell.row, sel.activeCell.column);
  }
  getValidationAt(row: number, column: number): string[] | undefined {
    const rule = findValidationRule(this.runtime.model.getSheet(this.activeSheetId), row, column);
    return rule ? validationList(rule) : undefined;
  }

  printWorkbook(layout: PrintLayout): void {
    if (!this.canExecute('print.preview')) {
      this.notify('You do not have permission to print');
      return;
    }
    const range = this.selectionService.primaryRangeOrDefault();
    this.runCommand('print.preview', { layout, sheetId: this.activeSheetId, range });
    const snapshot = this.rebuildPrintSnapshot(layout, range);
    this.dialogs = { ...this.dialogs, active: 'print-preview' };
    this.setFocusState('dialog', 'dialog');
    this.notify(summarizePrintSnapshot(snapshot));
    this.emit();
  }

  exportPdf(layout: PrintLayout): void {
    if (!this.canExecute('print.export')) {
      this.notify('You do not have permission to export PDF');
      return;
    }
    const range = this.selectionService.primaryRangeOrDefault();
    this.runCommand('print.export', { layout, sheetId: this.activeSheetId, range });
    const snapshot = this.rebuildPrintSnapshot(layout, range);
    this.dialogs = { ...this.dialogs, active: 'print-preview' };
    this.setFocusState('dialog', 'dialog');
    void this.executePdfExport(snapshot);
    this.notify(summarizePrintSnapshot(snapshot));
    this.emit();
  }

  getPrintSnapshot(): PrintSnapshot | null {
    return this.printSnapshot;
  }

  getPrintPageSetup(): PageSetup {
    return structuredClone(getPrintDocument(this.runtime.model, this.activeSheetId).pageSetup);
  }

  setPrintArea(range: RangeRef): void {
    if (!this.canExecute('print.area.set')) {
      this.notify('You do not have permission to set print area');
      return;
    }
    this.runCommand('print.area.set', { sheetId: range.sheetId, range });
    this.rebuildPrintSnapshot(this.printLayout, range);
    this.notify('Print area updated');
    this.emit();
  }

  updatePrintPageSetup(layout: PrintLayout): void {
    if (!this.canExecute('print.pageSetup')) {
      this.notify('You do not have permission to change print setup');
      return;
    }
    this.runCommand('print.pageSetup', { layout, sheetId: this.activeSheetId });
    this.rebuildStoredPrintSnapshot(layout);
    this.emit();
  }

  openPrintLayout(): void {
    this.panels = { ...this.panels, active: 'print', open: true };
    this.emit();
  }

  setCurrentPrintArea(): void {
    this.setPrintArea({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
  }

  clearPrintArea(): void {
    this.runCommand('print.area.clear', { sheetId: this.activeSheetId });
    this.rebuildStoredPrintSnapshot();
    this.notify('Print area cleared');
    this.emit();
  }

  setPrintTitles(axis: 'rows' | 'columns'): void {
    const range = this.getPrimaryRange();
    const params = axis === 'rows'
      ? { sheetId: this.activeSheetId, repeatRows: { start: range.startRow, end: range.endRow } }
      : { sheetId: this.activeSheetId, repeatColumns: { start: range.startColumn, end: range.endColumn } };
    this.runCommand('print.titles.set', params);
    this.rebuildStoredPrintSnapshot();
    this.notify(axis === 'rows' ? 'Rows to repeat at top updated' : 'Columns to repeat at left updated');
    this.emit();
  }

  clearPrintTitles(): void {
    this.runCommand('print.titles.clear', { sheetId: this.activeSheetId });
    this.rebuildStoredPrintSnapshot();
    this.notify('Print titles cleared');
    this.emit();
  }

  setPrintPageBreak(pageBreak: { row?: number; column?: number }): void {
    const next: PrintPageBreak = { sheetId: this.activeSheetId, ...pageBreak };
    this.runCommand('print.pageBreak.set', { sheetId: this.activeSheetId, pageBreak: next });
    this.rebuildStoredPrintSnapshot();
    this.emit();
  }

  removePrintPageBreak(pageBreak: { row?: number; column?: number }): void {
    const target: PrintPageBreak = { sheetId: this.activeSheetId, ...pageBreak };
    this.runCommand('print.pageBreak.remove', { sheetId: this.activeSheetId, pageBreak: target });
    this.rebuildStoredPrintSnapshot();
    this.emit();
  }

  clearPrintPageBreaks(): void {
    this.runCommand('print.pageBreaks.clear', { sheetId: this.activeSheetId });
    this.rebuildStoredPrintSnapshot();
    this.notify('Manual page breaks cleared');
    this.emit();
  }

  setPrintScale(scale: number, fitToWidth?: number | null, fitToHeight?: number | null): void {
    this.runCommand('print.scale.set', {
      sheetId: this.activeSheetId,
      scale,
      ...(fitToWidth === undefined ? {} : { fitToWidth }),
      ...(fitToHeight === undefined ? {} : { fitToHeight }),
    });
    const document = getPrintDocument(this.runtime.model, this.activeSheetId);
    this.printLayout = {
      ...this.printLayout,
      ...pageSetupToPrintLayout(document.pageSetup),
      scale: document.pageSetup.scale,
      fitToWidth: Boolean(document.pageSetup.fitToWidth),
      fitToHeight: Boolean(document.pageSetup.fitToHeight),
    };
    this.rebuildStoredPrintSnapshot(this.printLayout);
    this.emit();
  }

  setPrintScaleToFit(fitToWidth: number | null, fitToHeight: number | null): void {
    const scale = getPrintDocument(this.runtime.model, this.activeSheetId).pageSetup.scale;
    this.setPrintScale(scale, fitToWidth, fitToHeight);
  }

  setPrintGridlines(enabled: boolean): void {
    this.runCommand('print.gridlines.set', { sheetId: this.activeSheetId, enabled });
    this.printLayout = { ...this.printLayout, printGridlines: enabled };
    this.rebuildStoredPrintSnapshot(this.printLayout);
    this.emit();
  }

  setPrintHeadings(enabled: boolean): void {
    this.runCommand('print.headings.set', { sheetId: this.activeSheetId, enabled });
    this.printLayout = { ...this.printLayout, printHeadings: enabled };
    this.rebuildStoredPrintSnapshot(this.printLayout);
    this.emit();
  }

  setViewGridlines(enabled: boolean): void {
    this.runCommand('pageLayout.gridlines.view.set', { sheetId: this.activeSheetId, enabled });
    this.refresh();
  }

  setViewHeadings(enabled: boolean): void {
    this.runCommand('pageLayout.headings.view.set', { sheetId: this.activeSheetId, enabled });
    this.refresh();
  }

  toggleViewGridlines(): void {
    this.setViewGridlines(!this.runtime.model.getSheet(this.activeSheetId).showGridlines);
  }

  toggleViewHeadings(): void {
    this.setViewHeadings(!this.runtime.model.getSheet(this.activeSheetId).showHeaders);
  }

  togglePrintGridlines(): void {
    this.setPrintGridlines(!getPrintDocument(this.runtime.model, this.activeSheetId).pageSetup.printGridlines);
  }

  togglePrintHeadings(): void {
    this.setPrintHeadings(!getPrintDocument(this.runtime.model, this.activeSheetId).pageSetup.printHeadings);
  }

  private rebuildPrintSnapshot(layout?: PrintLayout, range?: RangeRef): PrintSnapshot {
    const uiLayout = layout ?? this.printLayout;
    const selectionRange = range ?? this.selectionService.primaryRangeOrDefault();
    const snapshot = buildPrintSnapshot(
      this.runtime.model,
      this.activeSheetId,
      uiLayout,
      selectionRange,
    );
    this.printLayout = uiLayout;
    this.printSnapshot = snapshot;
    return snapshot;
  }

  private rebuildStoredPrintSnapshot(layout?: PrintLayout): PrintSnapshot {
    const uiLayout = layout ?? this.printLayout;
    const snapshot = buildPrintSnapshot(this.runtime.model, this.activeSheetId, uiLayout);
    this.printLayout = uiLayout;
    this.printSnapshot = snapshot;
    return snapshot;
  }

  private async executePdfExport(snapshot: PrintSnapshot): Promise<void> {
    const service = new PdfExportService(browserPrintHook);
    try {
      const output = await service.export(snapshot.model, snapshot.pages, {
        filename: `${this.runtime.model.name || 'workbook'}.pdf`,
        title: this.runtime.model.name,
      });
      if (typeof document !== 'undefined' && typeof URL !== 'undefined') {
        const blob = output instanceof Blob
          ? output
          : new Blob([
              output.buffer.slice(output.byteOffset, output.byteOffset + output.byteLength) as ArrayBuffer,
            ], { type: 'application/pdf' });
        const href = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = href;
        link.download = `${this.runtime.model.name || 'workbook'}.pdf`;
        link.click();
        URL.revokeObjectURL(href);
      }
      this.notify('PDF exported');
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'PDF export failed');
    }
  }

  async loadQuery(query: QueryDefinition, target?: LoadTarget): Promise<void> {
    if (!this.canExecute('query.load')) {
      const error = new Error('You do not have permission to load queries');
      this.notify(error.message);
      throw error;
    }
    try {
      const resolvedTarget = target ?? resolveLoadTarget(
        this.activeSheetId,
        this.selectionService.primaryRangeOrDefault(),
      );
      const result = await this.executeQuery(query);
      const persistedQuery = { ...structuredClone(query), lastTarget: structuredClone(resolvedTarget) };
      this.runCommand('query.load', { query: persistedQuery, target: resolvedTarget, result });
      const snapshot = buildQueryResultSnapshot(persistedQuery, result, resolvedTarget);
      this.querySessions.set(query.id, { definition: persistedQuery, lastResult: snapshot });
      this.lastQueryResult = snapshot;
    this.panels = { ...this.panels, active: 'query', open: true };
      this.notify(summarizeQueryResult(snapshot));
      this.refresh();
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Query load failed');
      this.emit();
      throw error;
    }
  }

  async refreshQuery(queryId: string): Promise<void> {
    const session = this.querySessions.get(queryId);
    if (!session) {
      const error = new Error('Query not found');
      this.notify(error.message);
      throw error;
    }
    if (!this.canExecute('query.refresh')) {
      const error = new Error('You do not have permission to refresh queries');
      this.notify(error.message);
      throw error;
    }
    try {
      const target = session.lastResult?.target ?? session.definition.lastTarget ?? resolveLoadTarget(
        this.activeSheetId,
        this.selectionService.primaryRangeOrDefault(),
      );
      const result = await this.executeQuery(session.definition);
      session.definition.lastTarget = structuredClone(target);
      this.runCommand('query.refresh', {
        queryId,
        query: session.definition,
        target,
        result,
      });
      const snapshot = buildQueryResultSnapshot(session.definition, result, target);
      session.lastResult = snapshot;
      this.lastQueryResult = snapshot;
      this.notify(summarizeQueryResult(snapshot));
      this.refresh();
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Query refresh failed');
      this.emit();
      throw error;
    }
  }

  async testQueryConnection(
    connectorId: string,
    config: Record<string, unknown>,
  ): Promise<{ ok: boolean; message?: string }> {
    const connector = this.runtime.connectors.get(connectorId);
    return connector.testConnection(config);
  }

  private async executeQuery(query: QueryDefinition): Promise<import('./features/query').QueryResult> {
    if (query.connectorId !== 'sqlite' && query.connectorId !== 'jdbc' && query.connectorId !== 'rest') {
      return executeQueryDefinition(this.runtime.connectors, query);
    }
    if (this.runtime.localOnly) {
      throw new Error(`Connector ${query.connectorId} requires the Java backend`);
    }
    const config = query.connectorConfig;
    const sourceRef = typeof config.sourceRef === 'string' ? config.sourceRef.trim() : '';
    const statement = typeof config.statement === 'string'
      ? config.statement
      : typeof config.query === 'string'
        ? config.query
        : '';
    if (!sourceRef || !statement) {
      throw new Error(`Connector ${query.connectorId} requires server sourceRef and statement`);
    }
    const method = typeof config.method === 'string' && (config.method === 'GET' || config.method === 'POST')
      ? config.method
      : undefined;
    const request: ServerQueryRequest = {
      queryId: query.id,
      name: query.name,
      connectorId: query.connectorId,
      sourceRef,
      statement,
      ...(method === undefined ? {} : { method }),
      ...(Array.isArray(config.parameters) ? { parameters: structuredClone(config.parameters) } : {}),
      ...(config.body === undefined ? {} : { body: structuredClone(config.body) }),
      steps: query.steps.map((step) => ({
        id: step.id,
        kind: step.kind,
        name: step.name,
        config: structuredClone(step.config),
        enabled: step.enabled,
      })),
    };
    const response = await this.runtime.api.executeServerQuery(this.runtime.model.unitId, request);
    if (response.rowCount !== response.rows.length) throw new Error('Java backend returned an invalid query row count');
    return { columns: response.columns, rows: response.rows, rowCount: response.rowCount };
  }

  getQuerySnapshot(): {
    lastResult: QueryResultSnapshot | null;
    connectors: string[];
    loadedQueries: QueryResultSnapshot[];
  } {
    return {
      lastResult: this.lastQueryResult,
      connectors: this.runtime.connectors.list().map((connector) => connector.id),
      loadedQueries: [...this.querySessions.values()]
        .map((session) => session.lastResult)
        .filter((result): result is QueryResultSnapshot => Boolean(result)),
    };
  }

  async runAutomationScript(source: string): Promise<void> {
    if (!this.canExecute('automation.run')) {
      this.notify('You do not have permission to run scripts');
      return;
    }
    this.lastScriptResult = await runAutomationScriptAsync(this.runtime.model, this.runtime.commands, source, undefined, {
      workerFactory: this.automationWorkerFactory,
    });
    this.panels = { ...this.panels, active: 'automate', open: true };
    this.ribbonTab = 'automate';
    this.notify(summarizeScriptResult(this.lastScriptResult));
    this.refresh();
  }

  async runSampleAutomationScript(): Promise<void> {
    await this.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
  }

  startAutomationRecording(): void {
    if (!this.canExecute('automation.record.start')) {
      this.notify('You do not have permission to record scripts');
      return;
    }
    this.runCommand('automation.record.start', {});
    this.recorderDetach?.();
    this.commandRecorder.start();
    this.recorderDetach = this.runtime.commands.onCommand(this.commandRecorder.createListener());
    this.automationRecording = true;
    this.recordedScript = '';
    this.panels = { ...this.panels, active: 'automate', open: true };
    this.ribbonTab = 'automate';
    this.notify('Recording automation script');
    this.emit();
  }

  stopAutomationRecording(): string {
    if (!this.canExecute('automation.record.stop')) {
      this.notify('You do not have permission to stop recording');
      return this.recordedScript;
    }
    this.recorderDetach?.();
    this.recorderDetach = null;
    this.runCommand('automation.record.stop', {});
    const statements = this.commandRecorder.stop();
    this.automationRecording = false;
    this.recordedScript = this.commandRecorder.toScript();
    this.notify(`Recorded ${statements.length} statement(s)`);
    this.emit();
    return this.recordedScript;
  }

  getAutomationSnapshot(): AutomationSnapshot {
    return {
      recording: this.automationRecording,
      recordedScript: this.recordedScript,
      lastResult: this.lastScriptResult,
      lastRunAt: this.lastScriptResult ? new Date().toISOString() : null,
    };
  }

  runGoalSeek(params: GoalSeekParams): GoalSeekResult {
    if (!this.canExecute('extended.whatIf.goalSeek')) {
      this.notify('You do not have permission to run Goal Seek');
      return {
        kind: 'goal-seek',
        status: 'not-converged',
        iterations: 0,
        message: 'Permission denied',
      };
    }
    const command = this.runCommand('extended.whatIf.goalSeek', { ...params, sheetId: this.activeSheetId }) as import('@react-sheets/command-runtime').CommandResult & { plan?: WhatIfPlan };
    const result = command.plan?.result as GoalSeekResult | undefined;
    if (!result) throw new Error('Goal Seek command did not return a plan');
    this.lastWhatIfResult = result;
    this.panels = { ...this.panels, active: 'extended', open: true };
    this.notify(result.message ?? `Goal Seek ${result.status}`);
    this.refresh();
    return result;
  }

  runScenarioAnalysis(scenario: ScenarioDefinition): ScenarioResult {
    if (!this.canExecute('extended.whatIf.scenario')) {
      this.notify('You do not have permission to run scenarios');
      return {
        kind: 'scenario',
        status: 'failed',
        scenarioId: scenario.id,
        message: 'Permission denied',
        outputs: [],
      };
    }
    const command = this.runCommand('extended.whatIf.scenario', { sheetId: this.activeSheetId, scenario }) as import('@react-sheets/command-runtime').CommandResult & { plan?: WhatIfPlan };
    const result = command.plan?.result as ScenarioResult | undefined;
    if (!result) throw new Error('Scenario command did not return a plan');
    this.lastWhatIfResult = result;
    this.panels = { ...this.panels, active: 'extended', open: true };
    this.notify(result.message);
    this.refresh();
    return result;
  }

  runDataTableAnalysis(params: DataTableParams): DataTableResult {
    if (!this.canExecute('extended.whatIf.dataTable')) {
      this.notify('You do not have permission to run data tables');
      return { kind: 'data-table', status: 'failed', message: 'Permission denied', filledCells: 0, writes: [] };
    }
    const command = this.runCommand('extended.whatIf.dataTable', { ...params, sheetId: this.activeSheetId }) as import('@react-sheets/command-runtime').CommandResult & { plan?: WhatIfPlan };
    const result = command.plan?.result as DataTableResult | undefined;
    if (!result) throw new Error('Data Table command did not return a plan');
    this.lastWhatIfResult = result;
    this.panels = { ...this.panels, active: 'extended', open: true };
    this.notify(result.message);
    this.refresh();
    return result;
  }

  getExtendedSnapshot(): ExtendedSnapshot {
    return {
      lastWhatIfResult: this.lastWhatIfResult,
    };
  }

  async exportXlsxWorkbook(fileName?: string): Promise<{ buffer: ArrayBuffer; fileName: string } | null> {
    if (!this.canExecute('xlsx.export')) {
      this.notify('You do not have permission to export workbooks');
      return null;
    }
    try {
      const exported = await exchangeExportXlsx(this.runtime.model.snapshot(), {
      fileName: fileName ?? `${this.runtime.model.name || 'workbook'}.xlsx`,
        nativePackage: this.nativePackage,
        execution: this.xlsxExecution,
        revision: this.version,
      });
      this.compatibilityReport = exported.report;
      this.notify(summarizeCompatibilityReport(exported.report));
      this.refresh();
      if (!exported.buffer || !exported.fileName) return null;
      return { buffer: exported.buffer, fileName: exported.fileName };
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'XLSX export failed');
      return null;
    }
  }

  clearCompatibilityReport(): void {
    this.compatibilityReport = null;
    this.refresh();
  }

  sortRange(criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader?: boolean): void {
    if (criteria.length === 0) return;
    const range = normalizeRangeRef(this.getCurrentRegion());
    if (range.endRow <= range.startRow) {
      this.notify('Select a data region with at least one data row before sorting');
      return;
    }
    const width = range.endColumn - range.startColumn + 1;
    const normalizedCriteria = criteria.map((criterion) => {
      const column = range.startColumn + criterion.colIdx;
      if (criterion.colIdx < 0 || criterion.colIdx >= width) {
        throw new Error('Selected sort column is outside the current data region');
      }
      return { column, ascending: criterion.ascending };
    });
    const detectedHeader = this.inferSortHeader(range);
    this.dispatch({ commandId: 'sheet.sort.multi', params: {
      sheetId: this.activeSheetId,
      range,
      criteria: normalizedCriteria,
      hasHeader: hasHeader ?? detectedHeader,
    } });
  }

  private inferSortHeader(range: RangeRef): boolean {
    const sheet = this.runtime.model.getSheet(range.sheetId);
    let textHeaders = 0;
    let populated = 0;
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const value = this.readResolvedCell(sheet, range.startRow, column)?.value;
      if (value == null || value === '') continue;
      populated += 1;
      if (typeof value === 'string') textHeaders += 1;
    }
    return populated > 0 && textHeaders === populated;
  }

  getRecalculationMode(): RecalculationMode {
    return this.runtime.formula.getRecalculationMode();
  }

  hasPendingFormulaRecalculation(): boolean {
    return this.runtime.formula.hasPendingRecalculation();
  }

  setRecalculationMode(mode: RecalculationMode): void {
    this.runCommand('formula.calculation.mode.set', { mode });
    if (mode === 'automatic') void scheduleFormulaRecalculation(this.runtime);
    this.refresh();
  }

  async recalculateFormulas(): Promise<void> {
    await scheduleFormulaRecalculation(this.runtime, true);
    this.refresh();
    this.notify('Formulas recalculated');
  }

  /** Resolves after the most recent asynchronous formula projection is visible. */
  async waitForFormulaCalculation(): Promise<void> {
    await this.runtime.formulaCalculation;
  }

  openCreateTableDialog(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const range = normalizeRangeRef(this.getPrimaryRange());
    const regions = this.dataRegionsIntersectingRanges(this.activeSheetId, [range]);
    if (regions.length > 0) {
      void this.materializeDataRegions(regions)
        .then(() => this.openCreateTableDialog())
        .catch((error) => this.notify(error instanceof Error ? error.message : 'Data region could not be prepared as a table'));
      return;
    }
    if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) {
      this.notify('Select a multi-cell range before creating a table');
      return;
    }
    this.openDialog('create-table');
  }

  createSheetTableFromDialog(request: { name: string; hasHeaderRow: boolean; styleName?: string }): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const range = normalizeRangeRef(this.getPrimaryRange());
    const usedNames = sheet.sheetTables.map((table) => table.name);
    const plan = planSheetTableCreation({
      sheetId: this.activeSheetId,
      range,
      name: request.name,
      hasHeaderRow: request.hasHeaderRow,
      ...(request.styleName ? { styleName: request.styleName } : {}),
      existingNames: usedNames,
      nextId,
      readCell: (row, column) => this.readResolvedCell(sheet, row, column)?.value,
    }, sheet);
    this.runCommand('sheetTable.add', plan.table);
    this.closeCreateTableDialog();
    this.notify(`Sheet table ${plan.table.name} created`);
    this.refresh();
  }

  /** Build the canonical table descriptor for APIs that already own confirmation. */
  planSheetTableFromSelection(request: { name: string; hasHeaderRow: boolean; styleName?: string }): SheetTableModel {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const range = normalizeRangeRef(this.getPrimaryRange());
    const usedNames = sheet.sheetTables.map((table) => table.name);
    return planSheetTableCreation({
      sheetId: this.activeSheetId,
      range,
      name: request.name,
      hasHeaderRow: request.hasHeaderRow,
      ...(request.styleName ? { styleName: request.styleName } : {}),
      existingNames: usedNames,
      nextId,
      readCell: (row, column) => this.readResolvedCell(sheet, row, column)?.value,
    }, sheet).table;
  }

  async toggleSheetTableTotalRow(tableId?: string, enabled?: boolean): Promise<void> {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selection = this.selectionService.getState();
    const table = tableId
      ? sheet.sheetTables.find((entry) => entry.id === tableId)
      : findSheetTableAt(sheet, selection.activeCell.row, selection.activeCell.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    const nextEnabled = enabled ?? !table.hasTotalRow;
    await this.executeCommandAfterMaterialization('sheetTable.toggleTotalRow', { sheetId: this.activeSheetId, tableId: table.id, enabled: nextEnabled });
    this.notify(nextEnabled ? `Total row added to ${table.name}` : `Total row removed from ${table.name}`);
    this.refresh();
  }

  openTableSettings(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const table = findSheetTableAt(sheet, this.selectionService.getState().activeCell.row, this.selectionService.getState().activeCell.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    this.panels = { ...this.panels, active: 'data', open: true };
    this.emit();
  }

  toggleActiveSheetTableOption(option: 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton'): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const active = this.selectionService.getState().activeCell;
    const table = findSheetTableAt(sheet, active.row, active.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    const next = { ...structuredClone(table), [option]: !table[option] };
    if (option === 'hasHeaderRow' && !next.hasHeaderRow) {
      next.autoFilter = undefined;
      next.showFilterButton = false;
    }
    if (option === 'showFilterButton' && !next.showFilterButton) next.autoFilter = undefined;
    this.runCommand('sheetTable.update', next);
  }

  convertActiveSheetTableToRange(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const active = this.selectionService.getState().activeCell;
    const table = findSheetTableAt(sheet, active.row, active.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    this.runCommand('sheetTable.convertToRange', { sheetId: this.activeSheetId, tableId: table.id });
  }

  resizeActiveSheetTable(range: RangeRef): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const active = this.selectionService.getState().activeCell;
    const table = findSheetTableAt(sheet, active.row, active.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    this.runCommand('sheetTable.update', { ...structuredClone(table), range: { ...structuredClone(range), sheetId: this.activeSheetId } });
  }

  setActiveSheetTableStyle(styleName: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const active = this.selectionService.getState().activeCell;
    const table = findSheetTableAt(sheet, active.row, active.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    const normalized = styleName.trim();
    if (!/^TableStyle[A-Za-z0-9]+$/.test(normalized)) {
      this.notify('Invalid table style');
      return;
    }
    this.runCommand('sheetTable.update', { ...structuredClone(table), styleName: normalized });
  }

  setActiveSheetTableName(name: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const active = this.selectionService.getState().activeCell;
    const table = findSheetTableAt(sheet, active.row, active.column);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    const normalized = name.trim();
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(normalized)) {
      this.notify('Invalid table name');
      return;
    }
    if (sheet.sheetTables.some((entry) => entry.id !== table.id && entry.name.toLocaleLowerCase() === normalized.toLocaleLowerCase())) {
      this.notify(`Sheet Table already exists: ${normalized}`);
      return;
    }
    this.runCommand('sheetTable.update', { ...structuredClone(table), name: normalized });
  }

  groupRowsFromSelection(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow) {
      this.notify('Select multiple rows to group');
      return;
    }
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const group = buildRowOutlineGroup(this.activeSheetId, range, sheet, nextId('outline'));
    this.runCommand('outline.group.add', { sheetId: this.activeSheetId, group });
    this.notify(`Grouped rows ${range.startRow + 1}-${range.endRow + 1}`);
    this.refresh();
  }

  ungroupRowsFromSelection(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const groups = groupsWithinRange(sheet.outline, 'row', range);
    if (groups.length === 0) {
      this.notify('No row groups in the current selection');
      return;
    }
    for (const group of groups) {
      this.runCommand('outline.group.remove', { sheetId: this.activeSheetId, groupId: group.id });
    }
    this.notify(`Ungrouped ${groups.length} row group(s)`);
    this.refresh();
  }

  toggleOutlineGroup(groupId: string): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const group = sheet.outline?.groups.find((entry) => entry.id === groupId);
    if (!group) return;
    this.runCommand('outline.group.toggle', { sheetId: this.activeSheetId, groupId, collapsed: !group.collapsed });
    this.refresh();
  }

  showOutlineLevel(level: 1 | 2 | 3): void {
    this.runCommand('outline.showLevel', { sheetId: this.activeSheetId, level });
    this.refresh();
  }

  groupColumnsFromSelection(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endColumn <= range.startColumn) {
      this.notify('Select multiple columns to group');
      return;
    }
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const group = buildColumnOutlineGroup(this.activeSheetId, range, sheet, nextId('outline'));
    this.runCommand('outline.group.add', { sheetId: this.activeSheetId, group });
    this.notify(`Grouped columns ${columnLabel(range.startColumn)}-${columnLabel(range.endColumn)}`);
    this.refresh();
  }

  ungroupColumnsFromSelection(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const groups = groupsWithinRange(sheet.outline, 'column', range);
    if (groups.length === 0) {
      this.notify('No column groups in the current selection');
      return;
    }
    for (const group of groups) {
      this.runCommand('outline.group.remove', { sheetId: this.activeSheetId, groupId: group.id });
    }
    this.notify(`Ungrouped ${groups.length} column group(s)`);
    this.refresh();
  }

  async textToColumnsFromSelection(delimiter = ','): Promise<void> {
    const range = normalizeRangeRef(this.getPrimaryRange());
    const selection = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const column = selection.activeCell.column;
    const targetRange: RangeRef = {
      sheetId: this.activeSheetId,
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: column,
      endColumn: column,
    };
    await this.executeCommandAfterMaterialization('data.textToColumns', {
      sheetId: this.activeSheetId,
      range: targetRange,
      delimiter,
      maxColumns: Math.min(8, Math.max(2, sheet.columnCount - column)),
    });
    this.notify('Text split into columns');
    this.refresh();
  }

  async applyDataSubtotal(): Promise<void> {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) {
      this.notify('Select a data range with at least two columns');
      return;
    }
    await this.executeCommandAfterMaterialization('data.subtotal', {
      sheetId: this.activeSheetId,
      range,
      groupColumn: range.startColumn,
      valueColumn: range.startColumn + 1,
      functionName: 'SUM',
    });
    this.notify('Subtotal summary created below selection');
    this.refresh();
  }

  async removeDuplicatesFromSelection(): Promise<void> {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow) {
      this.notify('Select a multi-row range before removing duplicates');
      return;
    }
    const columns: number[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) columns.push(column);
    await this.executeCommandAfterMaterialization('data.removeDuplicates', {
      sheetId: this.activeSheetId,
      range,
      columns,
      hasHeader: true,
    });
    this.notify('Duplicate rows removed');
    this.refresh();
  }

  async createDataTableFromSelection(): Promise<void> {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const primaryRange = this.getPrimaryRange();
    const sourceRange = primaryRange.startRow !== primaryRange.endRow || primaryRange.startColumn !== primaryRange.endColumn ? primaryRange : usedRangeOfSheet(sheet);
    await this.materializeDataRegions(this.dataRegionsIntersectingRanges(sourceRange.sheetId, [sourceRange]));
    const fieldNames = new Set<string>();
    const fields: WorkbookTableModel['fields'] = [];
    for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column++) {
      const rawName = String(this.readResolvedCell(sheet, sourceRange.startRow, column)?.value ?? '').trim() || `Column ${column - sourceRange.startColumn + 1}`;
      let name = rawName;
      let suffix = 2;
      while (fieldNames.has(name)) name = `${rawName} ${suffix++}`;
      fieldNames.add(name);
      const sample: CellData['value'][] = [];
      for (let row = sourceRange.startRow + 1; row <= Math.min(sourceRange.endRow, sourceRange.startRow + 1000); row++) sample.push(this.readResolvedCell(sheet, row, column)?.value ?? null);
      fields.push({ id: name, name, ordinal: fields.length, type: inferTableFieldType(sample) });
    }
    const table: WorkbookTableModel = { id: nextId('table'), name: `${sheet.name} table`, sourceSheetId: this.activeSheetId, sourceRange: { ...sourceRange }, rowCount: Math.max(0, sourceRange.endRow - sourceRange.startRow), fields, blockSize: 4096, blocks: [], revision: 0 };
    this.runCommand('table.add', table);
    this.notify(`Data table ${table.name} created`);
    this.refresh();
  }

  replyComment(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.activeCell.row, sel.activeCell.column);
    if (!thread) return;
    const reply = buildCommentReply(this.actorId, text, nextId('reply'));
    this.runCommand('comment.reply', { sheetId: this.activeSheetId, threadId: thread.id, reply });
    this.notify('Reply added');
    this.refresh();
  }
  resolveComment(): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.activeCell.row, sel.activeCell.column);
    if (!thread || thread.resolved) return;
    this.runCommand('comment.resolve', {
      sheetId: this.activeSheetId,
      threadId: thread.id,
      resolved: true,
      resolvedAt: new Date().toISOString(),
    });
    this.notify('Comment resolved');
    this.refresh();
  }
  removeComment(): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.activeCell.row, sel.activeCell.column);
    if (!thread) return;
    this.runCommand('comment.remove', { sheetId: this.activeSheetId, threadId: thread.id });
    this.notify('Comment removed');
    this.refresh();
  }
  addNote(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const note = buildCellNote(this.actorId, text, nextId('note'));
    this.runCommand('note.set', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      note,
    });
    this.notify('Note added');
    this.refresh();
  }
  removeNote(): void {
    const sel = this.selectionService.getState();
    this.runCommand('note.remove', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
    });
    this.notify('Note removed');
    this.refresh();
  }
  setNoteVisibility(visible: boolean): void {
    const sel = this.selectionService.getState();
    this.runCommand('note.visibility', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      visible,
    });
    this.refresh();
  }
  setHyperlink(url: string): void {
    if (!url.trim()) return;
    const sel = this.selectionService.getState();
    const hyperlink = parseUrlHyperlink(url, nextId('link'));
    this.runCommand('hyperlink.set', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      hyperlink,
    });
    this.notify('Hyperlink inserted');
    this.refresh();
  }
  removeHyperlink(): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    if (!getCellHyperlink(sheet, sel.activeCell.row, sel.activeCell.column)) return;
    this.runCommand('hyperlink.remove', {
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
    });
    this.notify('Hyperlink removed');
    this.refresh();
  }
}

export { resolveUnitId, resolveActorId, resolveShareToken };
