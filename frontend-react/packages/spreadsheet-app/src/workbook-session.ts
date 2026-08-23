import type {
  CellData,
  CellStyle,
  ChartDrawingPayload,
  ConditionalFormatRule,
  DataBlockRef,
  DataSourceManifest,
  DefinedNameModel,
  DataValidationRule,
  DrawingObject,
  DrawingTransform,
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
  SheetSnapshot,
  SheetTableModel,
  SheetDataRegion,
  SparklineModel,
  SparklineGroup,
  WorkbookTableModel,
  WorkbookSnapshot,
} from '@react-sheets/core-model';
import type { HistoryEntry, CommandDescriptor, CommandResult } from '@react-sheets/command-runtime';
import type { AuthTokenProvider, GuestShareRole, RevisionRecord, ServerQueryRequest, ShareTokenProvider, TableRowsResponse } from '@react-sheets/protocol';
import type { XlsxPackage, XlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import { computePivotResult, computePivotResultFromBlockSource, getPivotFieldCatalog as buildPivotFieldCatalog, normalizePivotDefinition } from './features/pivot/engine';
import {
  collectFindReplacements,
  copyRangeToClipboardData,
  createFilterModelForTable,
  defaultTotalsFunction,
  findSheetTableAt,
  findValidationRule,
  formatTsv,
  groupsWithinRange,
  buildRowOutlineGroup,
  buildColumnOutlineGroup,
  normalizeRangeRef,
  parseTsv,
  validationList,
  type ClipboardData,
  type GoToSpecialKind,
  type PasteMode,
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
import { buildAllSheetSnapshots, type CanvasSheetSnapshot } from './ui-snapshot';
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
import { encodeSheetDataRegion, type EncodedSheetDataBlock } from './features/data-source';
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
  type PersistenceSnapshotMeta,
} from './features/persistence';
import type { FormulaAuditProjection } from './features/formula-audit';
import {
  exchangeImportXlsx,
  exchangeExportXlsx,
  summarizeCompatibilityReport,
} from './features/xlsx';
import {
  buildPrintSnapshot,
  getPrintDocument,
  summarizePrintSnapshot,
  type PrintPageSnapshot,
  type PrintSnapshot,
} from './features/print';
import { browserPrintHook, PdfExportService, type PrintLayout } from './features/print';
import type { LoadTarget, QueryDefinition } from './features/query/query-steps';
import {
  buildQueryResultSnapshot,
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
import type { ActiveContext, AppPhase, PeerCursor, RibbonTabId, SaveState, SidebarPanelId, UiSessionIntent } from './types';

export interface WorkbookSessionOptions {
  initialPhase?: AppPhase;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  automationWorkerFactory?: AutomationWorkerFactory;
  /** Only Node/unit harnesses may opt into the inline exchange implementation. */
  xlsxExecution?: 'worker' | 'inline-test';
}

export interface UiSnapshot {
  unitId: string;
  workbookName: string;
  phase: AppPhase;
  saveState: SaveState;
  notice: string;
  selection: SelectionState;
  activeCell: string;
  activeSheetId: string;
  activePanel: SidebarPanelId;
  ribbonTab: RibbonTabId;
  formulaDraft: string;
  editingCell: { row: number; column: number } | null;
  zoom: number;
  sheets: CanvasSheetSnapshot[];
  selectedSheet: CanvasSheetSnapshot;
  selectedFloatingId: string | null;
  activeContext: ActiveContext;
  peers: PeerCursor[];
  collabStatus: 'connecting' | 'open' | 'closed';
  collabRevision: number;
  pendingChangeSetCount: number;
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
  dataSources: readonly DataSourceManifest[];
  definedNameModels: readonly DefinedNameModel[];
  showFunctionWizard: boolean;
  showSortDialog: boolean;
  showFindReplace: boolean;
  showGoTo: boolean;
  showPasteSpecial: boolean;
  showFormatCells: boolean;
  showShiftCells: boolean;
  showCreatePivotDialog: boolean;
  findQuery: string;
  showPrintPreview: boolean;
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
}

export interface ExtendedSnapshot {
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
}

export interface CreatePivotTableParams {
  source?: PivotDefinition['source'];
  destination: { kind: 'new-sheet' } | { kind: 'existing-sheet'; sheetId: string; anchor: { row: number; column: number } };
}

function snapshotUsedRange(sheet: SheetSnapshot): RangeRef | undefined {
  let startRow = Number.POSITIVE_INFINITY;
  let startColumn = Number.POSITIVE_INFINITY;
  let endRow = -1;
  let endColumn = -1;
  for (const [rawRow, columns] of Object.entries(sheet.cells)) {
    const row = Number(rawRow);
    if (!Number.isSafeInteger(row)) continue;
    for (const rawColumn of Object.keys(columns)) {
      const column = Number(rawColumn);
      if (!Number.isSafeInteger(column)) continue;
      startRow = Math.min(startRow, row);
      startColumn = Math.min(startColumn, column);
      endRow = Math.max(endRow, row);
      endColumn = Math.max(endColumn, column);
    }
  }
  if (!Number.isFinite(startRow) || !Number.isFinite(startColumn)) return undefined;
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

function preserveSparseOverlayCell(cell: CellData): boolean {
  return cell.formula !== undefined || cell.style !== undefined || cell.numberFormat !== undefined
    || cell.comment !== undefined || cell.note !== undefined || cell.hyperlinkDetail !== undefined;
}

function removeBlockMaterializedCells(sheet: SheetSnapshot, range: RangeRef): void {
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const columns = sheet.cells[String(row)];
    if (!columns) continue;
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = columns[String(column)];
      if (cell && !preserveSparseOverlayCell(cell)) delete columns[String(column)];
    }
    if (Object.keys(columns).length === 0) delete sheet.cells[String(row)];
  }
}

export function getInitialSessionPhase(): AppPhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'error' || queryPhase === 'empty' ? queryPhase : 'loading';
}

export class WorkbookSession {
  private readonly runtime: SpreadsheetRuntime;
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
  private activePanel: SidebarPanelId = 'inspector';
  private ribbonTab: RibbonTabId = 'home';
  private formulaDraft = '';
  private zoom = 100;
  private selectedFloatingId: string | null = null;
  private activeContext: ActiveContext = { kind: 'none' };
  private peers: PeerCursor[] = [];
  private collabStatus: 'connecting' | 'open' | 'closed' = 'closed';
  private remoteRevisions: RevisionRecord[] = [];
  private historyPreview: HistoryPreviewSession | null = null;
  private hasPendingOperations = false;
  private persistenceChecksum = '';
  private compatibilityReport: CompatibilityReport | null = null;
  /** Runtime package plus local-only original ZIP context for honest export. */
  private xlsxPackage: XlsxPackage | undefined;
  private xlsxSourceArtifact: XlsxSourceArtifact | undefined;
  private showFunctionWizard = false;
  private showSortDialog = false;
  private showFindReplace = false;
  private showGoTo = false;
  private showPasteSpecial = false;
  private showFormatCells = false;
  private showShiftCells = false;
  private showCreatePivotDialog = false;
  private findQuery = '';
  private showPrintPreview = false;
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
  private overrideTarget: { row: number; column: number } | null = null;
  private clipboardData: ClipboardData | null = null;
  private snapshotGeneration = 0;
  private cachedUiSnapshot: UiSnapshot | null = null;
  private cachedUiSnapshotGeneration = -1;

  constructor({ initialPhase = 'ready', authTokenProvider, shareTokenProvider, automationWorkerFactory, xlsxExecution = 'worker' }: WorkbookSessionOptions = {}) {
    const routeShareToken = shareTokenProvider ? null : resolveShareToken();
    this.runtime = createSpreadsheetRuntime({
      authTokenProvider,
      shareTokenProvider: shareTokenProvider ?? (routeShareToken ? () => routeShareToken : undefined),
    });
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
      this.ensureActiveSheetSession();
      this.runtime.formulaAudit.refresh();
      this.refresh();
    };
    this.runtime.handlers.onPhaseChange = (phase) => {
      this.phase = phase;
      this.emit();
    };
    this.runtime.handlers.onActiveSheetChange = (sheetId) => {
      this.activeSheetId = sheetId;
      this.emit();
    };
    this.runtime.handlers.onRemoteRevisions = (revisions) => {
      this.remoteRevisions = revisions;
      this.emit();
    };
    this.runtime.handlers.onWorkspacePersisted = () => {
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
    this.persistenceDispose = startPersistenceSession(this.runtime);
    this.collabDispose = startCollaborationSession(this.runtime, () =>
      `${this.activeSheetId}:${this.selectionService.getState().activeCell.row}:${this.selectionService.getState().activeCell.column}`,
      this.runtime.authTokenProvider,
      this.runtime.shareTokenProvider,
    );
    void this.runtime.persistenceReady.then(async () => {
      const artifact = await this.runtime.workspacePersistence.xlsxArtifacts.load(this.runtime.model.unitId);
      if (artifact) this.xlsxSourceArtifact = artifact;
    }).catch(() => undefined);
  }

  dispose(): void {
    this.recorderDetach?.();
    this.recorderDetach = null;
    this.collabDispose?.();
    this.persistenceDispose?.();
    this.collabDispose = null;
    this.persistenceDispose = null;
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  private emit(): void {
    this.snapshotGeneration += 1;
    for (const listener of this.listeners) listener();
  }

  private refresh(): void {
    this.syncPersistenceMeta();
    this.version += 1;
    this.emit();
  }

  private syncPersistenceMeta(): void {
    const meta = buildPersistenceMeta(
      this.runtime.model.snapshot(),
      this.runtime.remoteRevision,
      this.runtime.collaboration?.offlineQueue.getPendingCount() ?? 0,
      this.runtime.workspaceRecord,
    );
    this.hasPendingOperations = meta.hasPendingOperations;
    this.persistenceChecksum = meta.checksum;
  }

  getUiSnapshot = (): UiSnapshot => {
    if (this.cachedUiSnapshot && this.cachedUiSnapshotGeneration === this.snapshotGeneration) {
      return this.cachedUiSnapshot;
    }
    const sheets = buildAllSheetSnapshots(this.runtime.model, this.runtime.formula, this.runtime.pivotResults, this.runtime.dataContent);
    const selectedSheet = sheets.find((sheet) => sheet.id === this.activeSheetId) ?? sheets[0]!;
    const selection = this.selectionService.getState();
    const collaboration = this.getCollaborationSnapshot();
    const snapshot: UiSnapshot = {
      unitId: this.runtime.model.unitId,
      workbookName: this.runtime.model.name,
      phase: this.phase,
      saveState: this.saveState,
      notice: this.notice,
      selection,
      activeCell: this.selectionService.activeCell,
      activeSheetId: this.activeSheetId,
      activePanel: this.activePanel,
      ribbonTab: this.ribbonTab,
      formulaDraft: this.formulaDraft,
      editingCell: this.editSession.editingCell,
      zoom: this.zoom,
      sheets,
      selectedSheet,
      selectedFloatingId: this.selectedFloatingId,
      activeContext: this.activeContext,
      peers: this.peers,
      collabStatus: this.collabStatus,
      collabRevision: collaboration.revision,
      pendingChangeSetCount: collaboration.pendingCount,
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
      tables: [...this.runtime.model.tables.values()].map((table) => structuredClone(table)),
      dataSources: [...this.runtime.model.dataSources.values()].map((source) => structuredClone(source)),
      definedNameModels: structuredClone(this.runtime.model.definedNameModels),
      showFunctionWizard: this.showFunctionWizard,
      showSortDialog: this.showSortDialog,
      showFindReplace: this.showFindReplace,
      showGoTo: this.showGoTo,
      showPasteSpecial: this.showPasteSpecial,
      showFormatCells: this.showFormatCells,
      showShiftCells: this.showShiftCells,
      showCreatePivotDialog: this.showCreatePivotDialog,
      findQuery: this.findQuery,
      showPrintPreview: this.showPrintPreview,
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
  dispatch(descriptor: CommandDescriptor): void {
    try {
      if (this.phase !== 'ready') return;
      this.runCommand(descriptor.commandId, descriptor.params);
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Permission denied');
    }
  }

  /**
   * Catalog commands deliberately express intent, not duplicated UI state.
   * This is the single boundary that supplies the active worksheet/range for
   * commands whose domain contract operates on the current selection.
   */
  private resolveCommandContext(commandId: string, params?: unknown): unknown {
    if (!params || typeof params !== 'object' || Array.isArray(params)) return params;
    const input = params as Record<string, unknown>;
    const range = normalizeRangeRef({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
    const sheetId = typeof input.sheetId === 'string' && input.sheetId.trim() ? input.sheetId : this.activeSheetId;
    if (commandId === 'sheet.style.set' || commandId === 'sheet.merge.set' || commandId === 'sheet.range.clear') {
      return { ...input, sheetId, range: input.range ?? range };
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
        this.setActivePanel(intent.panel);
        if (intent.notice) this.notify(intent.notice);
        return;
      case 'dialog.open':
        this.openDialog(intent.dialog, intent.findQuery);
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

  runCommand(commandId: string, params?: unknown): CommandResult {
    const resolvedParams = this.resolveCommandContext(commandId, params);
    if (!this.runtime.commands.registry.hasCommand(commandId)) {
      throw new Error(`Unknown command: ${commandId}`);
    }
    this.assertPermission(commandId, resolvedParams);
    const result = this.runtime.commands.execute(commandId, resolvedParams);
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

  getClipboard(): ClipboardData | null {
    return this.clipboardData;
  }

  setClipboard(data: ClipboardData | null): void {
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
    const range = this.getPrimaryRange();
    this.runCommand('navigation.gotoSpecial', { sheetId: this.activeSheetId, range, kind });
  }

  getActiveSheetId(): string {
    return this.activeSheetId;
  }

  setActivePivotContext(pivotId: string | null, sheetId = this.activeSheetId): void {
    const next: ActiveContext = pivotId === null ? { kind: 'none' } : { kind: 'pivot', sheetId, pivotId };
    if (JSON.stringify(next) === JSON.stringify(this.activeContext)) return;
    this.activeContext = next;
    if (pivotId !== null) {
      this.activePanel = 'pivot';
      this.ribbonTab = 'pivotAnalyze';
    } else if (this.ribbonTab === 'pivotAnalyze' || this.ribbonTab === 'pivotDesign') {
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
      if (!this.runtime.commands.undo()) break;
    }
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
      await this.runtime.checkpointWorkspace();
      this.saveState = 'saved';
      this.syncPersistenceMeta();
      this.notify('Local workbook checkpoint saved');
    } catch (error) {
      this.saveState = error instanceof Error && error.message.includes('conflict') ? 'conflict' : 'error';
      this.notify(error instanceof Error ? error.message : 'Save failed');
      this.emit();
    }
  }

  notify(message: string): void {
    this.notice = message;
    this.emit();
  }

  undo(): void {
    if (this.runtime.commands.undo()) {
      this.ensureActiveSheetSession();
      this.syncDraftFromPrimary();
      this.notify('Undo applied');
      this.refresh();
    }
  }

  redo(): void {
    if (this.runtime.commands.redo()) {
      this.ensureActiveSheetSession();
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
    this.emit();
  }

  setActivePanel(panel: SidebarPanelId): void {
    this.activePanel = panel;
    this.emit();
  }

  setZoom(zoom: number): void {
    this.zoom = Math.max(50, Math.min(200, zoom));
    this.emit();
  }

  openDialog(dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells' | 'create-pivot', findQuery?: string): void {
    if (dialog === 'function-wizard') this.showFunctionWizard = true;
    if (dialog === 'sort-dialog') this.showSortDialog = true;
    if (dialog === 'find-replace') {
      this.findQuery = findQuery ?? '';
      this.showFindReplace = true;
    }
    if (dialog === 'goto') this.showGoTo = true;
    if (dialog === 'paste-special') this.showPasteSpecial = true;
    if (dialog === 'format-cells') this.showFormatCells = true;
    if (dialog === 'shift-cells') this.showShiftCells = true;
    if (dialog === 'create-pivot') this.showCreatePivotDialog = true;
    if (dialog === 'print-preview') {
      this.rebuildPrintSnapshot();
      this.showPrintPreview = true;
      this.activePanel = 'print';
    }
    this.emit();
  }

  closeFunctionWizard = (): void => {
    this.showFunctionWizard = false;
    this.emit();
  };
  closeSortDialog = (): void => {
    this.showSortDialog = false;
    this.emit();
  };
  closeFindReplace = (): void => {
    this.showFindReplace = false;
    this.findQuery = '';
    this.emit();
  };
  closeGoTo = (): void => {
    this.showGoTo = false;
    this.emit();
  };
  closePasteSpecial = (): void => {
    this.showPasteSpecial = false;
    this.emit();
  };
  closeFormatCells = (): void => {
    this.showFormatCells = false;
    this.emit();
  };
  closeShiftCells = (): void => {
    this.showShiftCells = false;
    this.emit();
  };
  closeCreatePivotDialog = (): void => {
    this.showCreatePivotDialog = false;
    this.emit();
  };
  pasteSpecial(mode: PasteMode): void {
    this.paste(mode);
    this.closePasteSpecial();
  };
  setShowPrintPreview = (open: boolean): void => {
    this.showPrintPreview = open;
    this.emit();
  };

  syncDraftFromPrimary(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.activeCell.row, sel.activeCell.column);
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
    this.emit();
  }

  selectRange(range: { startRow: number; startColumn: number; endRow: number; endColumn: number }, mode: 'replace' | 'add' | 'extend' = 'replace'): void {
    this.selectionService.selectRange(range, mode);
    this.syncDraftFromPrimary();
    this.emit();
  }

  /** Commit a canvas selection exactly, including its active cell and anchor. */
  applyCanvasSelection(selection: SelectionState): void {
    this.selectionService.applyState(selection);
    this.syncDraftFromPrimary();
    this.emit();
  }

  extendSelectionTo(row: number, column: number): void {
    this.selectRange({ startRow: row, startColumn: column, endRow: row, endColumn: column }, 'extend');
  }

  formatCells(params: { numberFormat?: string; style?: Partial<import('@react-sheets/core-model').CellStyle> }): void {
    const ranges = this.selectionService.getState().ranges;
    if (ranges.length === 0) return;
    this.runCommand('sheet.format.set', {
      sheetId: this.activeSheetId,
      ranges,
      numberFormat: params.numberFormat,
      style: params.style,
    });
    this.refresh();
  }

  shiftCells(direction: 'down' | 'up' | 'right' | 'left'): void {
    const range = this.getPrimaryRange();
    this.runCommand('sheet.cells.shift', { sheetId: this.activeSheetId, range, direction });
    this.refresh();
  }

  freezeAtPrimary(): void {
    const sel = this.selectionService.getState();
    this.runCommand('sheet.freeze.set', {
      sheetId: this.activeSheetId,
      freeze: {
        xSplit: sel.activeCell.column,
        ySplit: sel.activeCell.row,
        startRow: sel.activeCell.row,
        startColumn: sel.activeCell.column,
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
      const cellValue = horizontal ? sheet.cells.get(row, cursor)?.value : sheet.cells.get(cursor, column)?.value;
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
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selection = this.selectionService.getState();
    const range = this.getPrimaryRange();
    let target = { ...selection.activeCell };
    let source: RangeRef | undefined;
    if (range.startRow !== range.endRow || range.startColumn !== range.endColumn) {
      target = { row: Math.min(sheet.rowCount - 1, range.endRow + 1), column: range.startColumn };
      source = { ...range, sheetId: this.activeSheetId };
    } else {
      let startRow = target.row - 1;
      while (startRow >= 0) {
        const value = sheet.cells.get(startRow, target.column)?.value;
        if (value == null || value === '') break;
        startRow -= 1;
      }
      if (startRow + 1 < target.row) {
        source = {
          sheetId: this.activeSheetId,
          startRow: startRow + 1,
          endRow: target.row - 1,
          startColumn: target.column,
          endColumn: target.column,
        };
      }
    }
    if (!source || (source.startRow === target.row && source.startColumn === target.column)) {
      this.notify('Select values above or a range before using AutoSum');
      return;
    }
    const formula = `=SUM(${cellAddress(source.startRow, source.startColumn)}:${cellAddress(source.endRow, source.endColumn)})`;
    this.overrideTarget = target;
    const committed = this.commitFormula(formula);
    this.overrideTarget = null;
    if (committed) {
      this.selectionService.selectRange({ startRow: target.row, endRow: target.row, startColumn: target.column, endColumn: target.column }, 'replace');
      this.syncDraftFromPrimary();
      this.refresh();
    }
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

  beginEdit(initialText?: string): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    for (const spill of sheet.spillRanges) {
      if (isSpillChild(spill, sel.activeCell.row, sel.activeCell.column)) {
        this.notify('Spill cells are read-only');
        return;
      }
    }
    const cell = sheet.cells.get(sel.activeCell.row, sel.activeCell.column);
    this.editSession.begin({
      sheetId: this.activeSheetId,
      row: sel.activeCell.row,
      column: sel.activeCell.column,
      cell,
      selection: this.selectionService.getSnapshot(),
      initialText,
    });
    this.formulaDraft = this.editSession.active?.currentDraft ?? '';
    this.emit();
  }

  cancelEdit(): void {
    this.formulaDraft = this.editSession.cancel();
    this.emit();
  }

  commitFormula(overrideValue?: string): boolean {
    if (this.phase !== 'ready') return false;
    const sel = this.selectionService.getState();
    const row = this.overrideTarget?.row ?? sel.activeCell.row;
    const column = this.overrideTarget?.column ?? sel.activeCell.column;
    const text = overrideValue !== undefined ? overrideValue : this.formulaDraft;
    const style = this.runtime.model.getSheet(this.activeSheetId).cells.get(row, column)?.style;
    try {
      this.runCommand('sheet.cell.commitText', {
        sheetId: this.activeSheetId,
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
    this.runtime.model.getSheet(sheetId);
    this.activeSheetId = sheetId;
    this.selectionService.resetForSheet(sheetId);
    this.editSession.cancel();
    this.formulaDraft = '';
    this.activeContext = { kind: 'none' };
    this.refresh();
  }

  // ---- Pro / data features ----

  addChart(drawing: DrawingObject, payload: ChartDrawingPayload): void {
    if (drawing.kind !== 'chart' || payload.kind !== 'chart' || drawing.payloadId !== payload.chartId) {
      throw new Error(`Chart drawing and payload identity mismatch: ${drawing.id}`);
    }
    this.runCommand('chart.insert', buildChartInsertParams(drawing, payload));
    this.runCommand('drawing.select', { sheetId: drawing.sheetId, drawingIds: [drawing.id] });
    this.selectedFloatingId = drawing.id;
    this.notify(payload.title ? `Added chart "${payload.title}"` : `Added ${payload.chartType} chart`);
    this.refresh();
  }
  insertQuickChart(type: ChartDrawingPayload['chartType'] = 'column'): void {
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
      title: 'Chart',
      sourceRanges: [{ ...range, sheetId: this.activeSheetId }],
      legendPosition: 'bottom',
      showDataLabels: false,
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
      series: series?.map((entry) => ({ name: entry.name, range: entry.range, color: entry.color })),
      categoryRange,
    });
    this.refresh();
  }
  setChartLegend(chartId: string, legendPosition: NonNullable<ChartDrawingPayload['legendPosition']>): void {
    this.runCommand('chart.setLegend', { sheetId: this.activeSheetId, chartId, legendPosition });
    this.refresh();
  }
  setChartDataLabels(chartId: string, showDataLabels: boolean): void {
    this.runCommand('chart.setDataLabels', { sheetId: this.activeSheetId, chartId, showDataLabels });
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
    if (drawing) this.runtime.drawing.deselect(this.activeSheetId, [drawing.id]);
    if (this.selectedFloatingId === drawing.id) this.selectedFloatingId = null;
    this.refresh();
  }
  addPivot(pivot: PivotModel): void {
    this.runCommand('pivot.add', pivot);
    this.recomputePivotResult(pivot.id);
    this.notify(`Pivot ${pivot.id} added`);
    this.refresh();
  }
  insertQuickPivot(): string | undefined {
    const descriptor = this.buildQuickPivotDescriptor();
    const pivot = descriptor?.params as PivotModel | undefined;
    if (!pivot) {
      this.notify('Select a data range with category and value fields');
      return undefined;
    }
    this.addPivot(pivot);
    return pivot.id;
  }

  buildQuickPivotDescriptor(): CommandDescriptor | undefined {
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
    let createdSheetId: string | undefined;
    if (params.destination.kind === 'new-sheet') {
      targetSheetId = nextId('sheet');
      targetPosition = { row: 0, column: 0 };
      const names = new Set(this.runtime.model.getSheets().map((sheet) => sheet.name.toLocaleLowerCase()));
      let suffix = 1;
      let targetName = 'PivotTable';
      while (names.has(targetName.toLocaleLowerCase())) targetName = `PivotTable${suffix++}`;
      try {
        this.runCommand('sheet.add', { id: targetSheetId, name: targetName });
        createdSheetId = targetSheetId;
      } catch (error) {
        this.notify(error instanceof Error ? error.message : 'Could not create PivotTable worksheet');
        return undefined;
      }
    } else {
      targetSheetId = params.destination.sheetId;
      targetPosition = { ...params.destination.anchor };
      try {
        this.runtime.model.getSheet(targetSheetId);
      } catch {
        this.notify('PivotTable destination worksheet does not exist');
        return undefined;
      }
    }

    const blockRegion = this.runtime.model.getSheet(sourceRegion.sheetId).dataRegions.find((region) => region.range.startRow === sourceRegion.startRow
      && region.range.endRow === sourceRegion.endRow && region.range.startColumn === sourceRegion.startColumn && region.range.endColumn === sourceRegion.endColumn);
    const source = params.source
      ?? (blockRegion
        ? { kind: 'data-source' as const, dataSourceId: blockRegion.sourceId }
        : { kind: 'worksheet-range' as const, range: { ...sourceRegion } });
    const draft = {
      schema: 'PivotDefinition' as const,
      id: pivotId,
      source,
      target: { sheetId: targetSheetId, anchor: targetPosition },
      fieldCatalog: { fields: [] },
      refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
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
    } satisfies PivotDefinition;
    const pivotDraft: PivotModel = draft;
    const pivot: PivotModel = { ...pivotDraft, fieldCatalog: buildPivotFieldCatalog(this.runtime.model, pivotDraft) };
    if (pivot.fieldCatalog.fields.length === 0) {
      if (createdSheetId) this.runCommand('sheet.remove', { sheetId: createdSheetId });
      this.notify('PivotTable source does not contain usable fields');
      return undefined;
    }
    try {
      this.addPivot(pivot);
      this.activeSheetId = targetSheetId;
      this.selectionService.resetForSheet(targetSheetId);
      this.activePanel = 'pivot';
      this.notify('Blank PivotTable created. Drag fields into the Field List to build the report.');
      this.refresh();
      return pivotId;
    } catch (error) {
      if (createdSheetId && this.runtime.model.sheets.has(createdSheetId)) {
        try { this.runCommand('sheet.remove', { sheetId: createdSheetId }); } catch { /* keep original failure authoritative */ }
      }
      this.notify(error instanceof Error ? error.message : 'Could not create PivotTable');
      return undefined;
    }
  }
  updatePivotLayout(pivotId: string, layout: PivotLayout): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, layout });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  updatePivotConfiguration(
    pivotId: string,
    patch: Parameters<WorkbookSession['updatePivotLayout']>[1] extends PivotLayout
      ? { source?: PivotDefinition['source']; target?: PivotDefinition['target']; fieldCatalog?: PivotDefinition['fieldCatalog']; refreshPolicy?: PivotDefinition['refreshPolicy']; nativeMetadata?: PivotDefinition['nativeMetadata']; layout?: PivotLayout }
      : never,
  ): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, ...patch });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  setPivotAggregate(pivotId: string, fieldId: string, summarizeBy: PivotAggregateFunction): void {
    this.runCommand('pivot.setAggregate', { sheetId: this.activeSheetId, pivotId, fieldId, summarizeBy });
    this.recomputePivotResult(pivotId);
    this.refresh();
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
    this.recomputePivotResult(pivotId);
    this.refresh();
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
  addShape(drawing: DrawingObject, payload: ShapeDrawingPayload): void {
    if (drawing.kind !== 'shape' || payload.kind !== 'shape') {
      throw new Error(`Shape drawing and payload kind mismatch: ${drawing.id}`);
    }
    this.runCommand('drawing.add.shape', buildDrawingAdd(drawing, payload));
    this.runCommand('drawing.select', { sheetId: drawing.sheetId, drawingIds: [drawing.id] });
    this.selectedFloatingId = drawing.id;
    this.notify(`Added ${payload.type} shape`);
    this.refresh();
  }
  insertQuickShape(type: ShapeDrawingPayload['type'] = 'rounded-rectangle'): void {
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
      this.runtime.drawing.deselect(this.activeSheetId, [drawing.id]);
    } else {
      this.notify('Shape is not registered as a drawing object');
      return;
    }
    if (this.selectedFloatingId === drawing.id) this.selectedFloatingId = null;
    this.refresh();
  }
  addImage(drawing: DrawingObject, payload: ImageDrawingPayload): void {
    if (drawing.kind !== 'image' || payload.kind !== 'image') {
      throw new Error(`Image drawing and payload kind mismatch: ${drawing.id}`);
    }
    this.runCommand('drawing.add.image', buildDrawingAdd(drawing, payload));
    this.runCommand('drawing.select', { sheetId: drawing.sheetId, drawingIds: [drawing.id] });
    this.selectedFloatingId = drawing.id;
    this.notify('Image placed on canvas');
    this.refresh();
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
      this.runtime.drawing.deselect(this.activeSheetId, [drawing.id]);
    } else {
      this.notify('Image is not registered as a drawing object');
      return;
    }
    if (this.selectedFloatingId === drawing.id) this.selectedFloatingId = null;
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
    this.runtime.drawing.deselect(this.activeSheetId, [drawing.id]);
    this.selectedFloatingId = null;
    this.refresh();
  }
  private resolveSelectedDrawingId(): string | undefined {
    if (!this.selectedFloatingId) return undefined;
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    return sheet.drawings.some((entry) => entry.id === this.selectedFloatingId) ? this.selectedFloatingId : undefined;
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
    } & Partial<Pick<SparklineModel, 'color' | 'negativeColor' | 'highlightMax' | 'highlightMin' | 'groupId'>>,
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
    const stylePatch: Partial<SparklineModel> = {};
    if (params.color) stylePatch.color = params.color;
    if (params.negativeColor) stylePatch.negativeColor = params.negativeColor;
    if (params.highlightMax != null) stylePatch.highlightMax = params.highlightMax;
    if (params.highlightMin != null) stylePatch.highlightMin = params.highlightMin;
    if (Object.keys(stylePatch).length > 0) {
      this.runCommand('sparkline.update', { sheetId: this.activeSheetId, sparklineId, patch: stylePatch });
    }
    this.notify(`Sparkline inserted at row ${params.location.row + 1}`);
    this.refresh();
    return sparklineId;
  }
  insertQuickSparkline(type: SparklineModel['type'] = 'line'): string | undefined {
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
    this.runCommand('sheet.cf.add', { sheetId: this.activeSheetId, rule });
  }
  removeConditionalFormat(ruleId: string): void {
    this.runCommand('sheet.cf.remove', { sheetId: this.activeSheetId, ruleId });
  }
  addDataValidation(rule: DataValidationRule): void {
    this.runCommand('sheet.dv.add', { sheetId: this.activeSheetId, rule });
  }
  removeDataValidation(ruleId: string): void {
    this.runCommand('sheet.dv.remove', { sheetId: this.activeSheetId, ruleId });
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

  applyFilter(column: number, patch: { selectedValues?: string[] | null; conditionOperator?: string; conditionValue?: string }): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const baseRange =
      sheet.filter?.range ??
      normalizeRangeRef({ sheetId: this.activeSheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) });
    const criteria = { ...(sheet.filter?.criteria ?? {}) };
    criteria[column] = { column, selectedValues: patch.selectedValues ?? undefined, conditionOperator: patch.conditionOperator, conditionValue: patch.conditionValue };
    this.runCommand('sheet.filter.set', { sheetId: this.activeSheetId, filter: { sheetId: this.activeSheetId, range: baseRange, criteria } });
  }

  applyFilterSelection(): void {
    const sel = this.selectionService.getState();
    const activeRange = sel.ranges[sel.primaryRangeIndex];
    if (!activeRange) return;
    const sheetModel = this.runtime.model.getSheet(this.activeSheetId);
    this.runCommand('sheet.filter.set', {
      sheetId: this.activeSheetId,
      filter: {
        sheetId: this.activeSheetId,
        range: normalizeRangeRef({ sheetId: this.activeSheetId, startRow: 0, endRow: Math.max(0, sheetModel.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheetModel.columnCount - 1) }),
        criteria: {},
      },
    });
  }

  clearFilter(): void {
    this.runCommand('sheet.filter.remove', { sheetId: this.activeSheetId });
  }

  findReplace(params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }): number {
    if (!params.find) return 0;
    const patches = collectFindReplacements(this.runtime.model, params);
    let count = 0;
    for (const patch of patches) {
      count += patch.values[0]!.length;
      this.runCommand('sheet.range.set', { sheetId: patch.sheetId, startRow: patch.startRow, startColumn: patch.startColumn, values: patch.values });
    }
    this.notify(`${count} replacement(s) applied`);
    return count;
  }

  insertRowsAtPrimary(count: number): void {
    this.runCommand('sheet.rows.insert', { sheetId: this.activeSheetId, at: this.selectionService.getState().activeCell.row, count });
  }
  deleteRowsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.runCommand('sheet.rows.delete', { sheetId: this.activeSheetId, at: range?.startRow ?? sel.activeCell.row, count: (range?.endRow ?? sel.activeCell.row) - (range?.startRow ?? sel.activeCell.row) + 1 });
  }
  insertColumnsAtPrimary(count: number): void {
    this.runCommand('sheet.columns.insert', { sheetId: this.activeSheetId, at: this.selectionService.getState().activeCell.column, count });
  }
  deleteColumnsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.runCommand('sheet.columns.delete', { sheetId: this.activeSheetId, at: range?.startColumn ?? sel.activeCell.column, count: (range?.endColumn ?? sel.activeCell.column) - (range?.startColumn ?? sel.activeCell.column) + 1 });
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
  splitByDelimiter(delimiter: string): void {
    const sel = this.selectionService.getState();
    const sheet = this.getSelectedSheet();
    this.runCommand('data.splitColumn', { sheetId: this.activeSheetId, row: sel.activeCell.row, column: sel.activeCell.column, delimiter, maxColumns: Math.min(sheet.columnCount - sel.activeCell.column - 1, 8) });
  }

  copy(): void {
    const range = this.getPrimaryRange();
    const data = copyRangeToClipboardData(this.runtime.model, range);
    this.setClipboard({ ...data, isCut: false });
    void navigator.clipboard.writeText(formatTsv(data.values));
    this.notify('Range copied');
  }
  cut(): void {
    const range = this.getPrimaryRange();
    const data = copyRangeToClipboardData(this.runtime.model, range);
    this.setClipboard({ ...data, isCut: true });
    void navigator.clipboard.writeText(formatTsv(data.values));
    this.notify('Cut to clipboard');
  }
  paste(mode: PasteMode = 'all'): void {
    const sel = this.selectionService.getState();
    const internal = this.clipboardData;
    if (internal) {
      this.runCommand('sheet.range.paste', {
        sheetId: this.activeSheetId,
        targetOrigin: { row: sel.activeCell.row, column: sel.activeCell.column },
        clipboard: internal,
        mode,
      });
      if (internal.isCut) this.clearClipboard();
      this.syncDraftFromPrimary();
      this.notify('Pasted from clipboard');
      return;
    }
    void navigator.clipboard.readText().then((text) => {
      if (!text) return;
      const clipboard: ClipboardData = {
        range: this.getPrimaryRange(),
        values: parseTsv(text),
      };
      this.runCommand('sheet.range.paste', {
        sheetId: this.activeSheetId,
        targetOrigin: { row: sel.activeCell.row, column: sel.activeCell.column },
        clipboard,
        mode,
      });
      this.syncDraftFromPrimary();
      this.notify('Pasted from clipboard');
    });
  }
  clearFormats(): void {
    this.runCommand('sheet.range.clear', { sheetId: this.activeSheetId, range: this.getPrimaryRange(), mode: 'formats' });
  }

  clearSelection(mode: 'contents' | 'formats' = 'contents'): void {
    this.runCommand('sheet.range.clear', { sheetId: this.activeSheetId, range: this.getPrimaryRange(), mode });
    this.syncDraftFromPrimary();
  }

  private getRangeMatrix(range: RangeRef): CellData[][] {
    const sheet = this.runtime.model.getSheet(range.sheetId);
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const rowValues: CellData[] = [];
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = sheet.cells.get(r, c);
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
    this.runCommand('sheet.row.resize', { sheetId: this.activeSheetId, row, height: Math.max(18, heightPx) });
  }
  resizeColumn(column: number, widthPx: number): void {
    this.runCommand('sheet.column.resize', { sheetId: this.activeSheetId, column, width: Math.max(24, widthPx) });
  }
  fillRange(targetRange: { startRow: number; endRow: number; startColumn: number; endColumn: number }): void {
    const sel = this.selectionService.getState();
    const primary = sel.ranges[sel.primaryRangeIndex] ?? sel.ranges[0];
    if (!primary) return;
    this.runCommand('sheet.autofill', {
      sheetId: this.activeSheetId,
      sourceRange: { sheetId: this.activeSheetId, startRow: primary.startRow, endRow: primary.endRow, startColumn: primary.startColumn, endColumn: primary.endColumn },
      targetRange: { sheetId: this.activeSheetId, ...targetRange },
    });
  }
  setSelectedFloatingId(id: string | null): void {
    this.selectedFloatingId = id;
    if (id) {
      const sheet = this.runtime.model.getSheet(this.activeSheetId);
      const drawing = sheet.drawings.find((entry) => entry.id === id);
      if (drawing) {
        this.runCommand('drawing.select', { sheetId: this.activeSheetId, drawingIds: [drawing.id] });
        this.activeContext = { kind: 'drawing', sheetId: this.activeSheetId, drawingId: drawing.id };
      } else {
        this.runCommand('drawing.deselect', { sheetId: this.activeSheetId });
        this.activeContext = { kind: 'none' };
      }
    } else {
      this.runCommand('drawing.deselect', { sheetId: this.activeSheetId });
      this.activeContext = { kind: 'none' };
    }
    this.emit();
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
    const table = this.runtime.model.tables.get(tableId);
    if (!table || !table.sourceSheetId || !table.sourceRange) return Promise.reject(new Error('Data table not found'));
    const sheet = this.runtime.model.getSheet(table.sourceSheetId);
    const start = Math.max(0, offset);
    const end = Math.min(table.rowCount, start + Math.max(1, limit));
    const rows: import('@react-sheets/core-model').TableScalar[][] = [];
    for (let rowOffset = start; rowOffset < end; rowOffset += 1) {
      rows.push(table.fields.map((field) => sheet.cells.get(
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
    const table = this.runtime.model.tables.get(tableId);
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
    this.showPrintPreview = true;
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
    this.showPrintPreview = true;
    void this.executePdfExport(snapshot);
    this.notify(summarizePrintSnapshot(snapshot));
    this.emit();
  }

  getPrintSnapshot(): PrintSnapshot | null {
    return this.printSnapshot;
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
    this.rebuildPrintSnapshot(layout);
    this.emit();
  }

  openPrintLayout(): void {
    this.activePanel = 'print';
    this.emit();
  }

  setCurrentPrintArea(): void {
    this.setPrintArea({ ...this.getPrimaryRange(), sheetId: this.activeSheetId });
  }

  clearPrintArea(): void {
    this.runCommand('print.area.clear', { sheetId: this.activeSheetId });
    this.rebuildPrintSnapshot();
    this.notify('Print area cleared');
    this.emit();
  }

  setPrintTitles(axis: 'rows' | 'columns'): void {
    const range = this.getPrimaryRange();
    const params = axis === 'rows'
      ? { sheetId: this.activeSheetId, repeatRows: { start: range.startRow, end: range.endRow } }
      : { sheetId: this.activeSheetId, repeatColumns: { start: range.startColumn, end: range.endColumn } };
    this.runCommand('print.titles.set', params);
    this.rebuildPrintSnapshot();
    this.notify(axis === 'rows' ? 'Rows to repeat at top updated' : 'Columns to repeat at left updated');
    this.emit();
  }

  setPrintScale(scale: number): void {
    this.runCommand('print.scale.set', { sheetId: this.activeSheetId, scale });
    const document = getPrintDocument(this.runtime.model, this.activeSheetId);
    this.printLayout = { ...this.printLayout, scale: document.pageSetup.scale };
    this.rebuildPrintSnapshot(this.printLayout);
    this.emit();
  }

  setPrintGridlines(enabled: boolean): void {
    this.runCommand('print.gridlines.set', { sheetId: this.activeSheetId, enabled });
    this.printLayout = { ...this.printLayout, printGridlines: enabled };
    this.rebuildPrintSnapshot(this.printLayout);
    this.emit();
  }

  setPrintHeadings(enabled: boolean): void {
    this.runCommand('print.headings.set', { sheetId: this.activeSheetId, enabled });
    this.printLayout = { ...this.printLayout, printHeadings: enabled };
    this.rebuildPrintSnapshot(this.printLayout);
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
      this.notify('You do not have permission to load queries');
      return;
    }
    try {
      const resolvedTarget = target ?? resolveLoadTarget(
        this.activeSheetId,
        this.selectionService.primaryRangeOrDefault(),
      );
      const result = await this.executeQuery(query);
      this.runCommand('query.load', { query, target: resolvedTarget, result });
      const snapshot = buildQueryResultSnapshot(query, result, resolvedTarget);
      this.querySessions.set(query.id, { definition: structuredClone(query), lastResult: snapshot });
      this.lastQueryResult = snapshot;
      this.activePanel = 'query';
      this.notify(summarizeQueryResult(snapshot));
      this.refresh();
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Query load failed');
      this.emit();
    }
  }

  async refreshQuery(queryId: string): Promise<void> {
    const session = this.querySessions.get(queryId);
    if (!session) {
      this.notify('Query not found');
      return;
    }
    if (!this.canExecute('query.refresh')) {
      this.notify('You do not have permission to refresh queries');
      return;
    }
    try {
      const target = session.lastResult?.target ?? resolveLoadTarget(
        this.activeSheetId,
        this.selectionService.primaryRangeOrDefault(),
      );
      const result = await this.executeQuery(session.definition);
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
    this.activePanel = 'automate';
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
    this.activePanel = 'automate';
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
    this.activePanel = 'extended';
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
    this.activePanel = 'extended';
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
    this.activePanel = 'extended';
    this.notify(result.message);
    this.refresh();
    return result;
  }

  getExtendedSnapshot(): ExtendedSnapshot {
    return {
      lastWhatIfResult: this.lastWhatIfResult,
    };
  }

  private async prepareLargeDataRegions(snapshot: WorkbookSnapshot): Promise<EncodedSheetDataBlock[]> {
    const preparedBlocks: EncodedSheetDataBlock[] = [];
    snapshot.dataSources ??= [];
    const existingSourceIds = new Set(snapshot.dataSources.map((source) => source.id));
    for (const sheet of snapshot.sheets) {
      if (sheet.dataRegions?.length) continue;
      const range = snapshotUsedRange(sheet);
      if (!range || range.endRow <= range.startRow) continue;
      const sourceId = `source:${sheet.id}:${range.startRow}:${range.startColumn}:${range.endRow}:${range.endColumn}`;
      if (existingSourceIds.has(sourceId)) continue;
      const encoded = await encodeSheetDataRegion({
        sheet,
        range,
        sourceId,
        sourceName: `${sheet.name} data`,
        regionId: `${sourceId}:region`,
        revision: this.version,
      });
      if (!encoded) continue;
      // Bytes are committed first. The snapshot is changed only after every
      // referenced block has passed its local checksum write.
      for (const block of encoded.blocks) {
        await this.runtime.workspacePersistence.dataBlocks.put(block.ref, block.payload);
      }
      snapshot.dataSources.push(encoded.manifest);
      sheet.dataRegions ??= [];
      sheet.dataRegions.push(encoded.region);
      removeBlockMaterializedCells(sheet, range);
      preparedBlocks.push(...encoded.blocks);
      existingSourceIds.add(sourceId);
    }
    return preparedBlocks;
  }

  async importXlsxBuffer(buffer: ArrayBuffer, fileName = 'import.xlsx'): Promise<void> {
    if (!this.canExecute('xlsx.import')) {
      this.notify('You do not have permission to import workbooks');
      return;
    }
    this.phase = 'loading';
    this.emit();
    try {
      const imported = await exchangeImportXlsx({ fileName, buffer, execution: this.xlsxExecution, revision: this.version });
      if (!imported.snapshot) throw new Error('XLSX import did not produce a workbook snapshot');
      const preparedBlocks = await this.prepareLargeDataRegions(imported.snapshot);
      hydrateRuntime(this.runtime, {
        snapshot: imported.snapshot,
        revision: this.runtime.remoteRevision,
      });
      this.compatibilityReport = imported.report as CompatibilityReport;
      // Preserve only the original archive across the session. Reconstructing
      // the package inside the export Worker avoids retaining a duplicate
      // expanded OOXML graph on the UI thread.
      this.xlsxPackage = undefined;
      this.xlsxSourceArtifact = imported.sourceArtifact;
      this.runtime.remoteConnected = false;
      if (typeof window !== 'undefined') {
        window.history.replaceState({}, '', `/workbooks/${encodeURIComponent(imported.snapshot.unitId)}`);
      }
      this.activeSheetId = this.runtime.model.primarySheetId;
      this.selectionService.resetForSheet(this.activeSheetId);
      for (const block of preparedBlocks) {
        // The synchronizer uploads only metadata-addressed bytes; it never
        // places bytes in an operation envelope.
        await this.runtime.dataBlocks.put(block.ref, block.payload);
      }
      this.runtime.localRevision += 1;
      const checkpointSnapshot = this.runtime.model.snapshot();
      const pendingJournal = this.runtime.operationJournal.read(checkpointSnapshot.unitId);
      this.runtime.workspaceRecord = imported.sourceArtifact
        ? await this.runtime.workspacePersistence.checkpointWithArtifact(
          checkpointSnapshot,
          this.runtime.localRevision,
          this.runtime.remoteRevision,
          this.runtime.localOnly ? 'local-only' : 'remote',
          imported.sourceArtifact,
          pendingJournal,
        )
        : await this.runtime.workspacePersistence.checkpoint(
          checkpointSnapshot,
          this.runtime.localRevision,
          this.runtime.remoteRevision,
          this.runtime.localOnly ? 'local-only' : 'remote',
          pendingJournal,
        );
      this.phase = 'ready';
      this.notify(summarizeCompatibilityReport(this.compatibilityReport));
      this.refresh();
    } catch (error) {
      this.phase = 'error';
      this.notify(error instanceof Error ? error.message : 'XLSX import failed');
      this.emit();
    }
  }

  async exportXlsxWorkbook(fileName?: string): Promise<{ buffer: ArrayBuffer; fileName: string } | null> {
    if (!this.canExecute('xlsx.export')) {
      this.notify('You do not have permission to export workbooks');
      return null;
    }
    try {
      const exported = await exchangeExportXlsx(this.runtime.model.snapshot(), {
        fileName: fileName ?? `${this.runtime.model.name || 'workbook'}.xlsx`,
        package: this.xlsxPackage,
        sourceArtifact: this.xlsxSourceArtifact,
        execution: this.xlsxExecution,
        revision: this.version,
      });
      this.compatibilityReport = exported.report;
      this.notify(summarizeCompatibilityReport(exported.report));
      this.refresh();
      this.xlsxPackage = undefined;
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

  sortRange(criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader: boolean): void {
    if (criteria.length === 0) return;
    const sel = this.selectionService.getState();
    const sheet = this.getSelectedSheet();
    const range =
      sel.ranges[sel.primaryRangeIndex] ??
      normalizeRangeRef({ sheetId: this.activeSheetId, startRow: 0, endRow: Math.min(40, sheet.rowCount - 1), startColumn: 0, endColumn: Math.min(sheet.columnCount - 1, 6) });
    this.runCommand('sheet.sort.multi', { sheetId: this.activeSheetId, range, criteria: criteria.map((c) => ({ column: c.colIdx, ascending: c.ascending })), hasHeader });
  }

  setRecalculationMode(mode: RecalculationMode): void {
    this.runtime.formula.setRecalculationMode(mode);
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

  createSheetTableFromSelection(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) {
      this.notify('Select a multi-cell range before creating a table');
      return;
    }
    const usedNames = new Set(sheet.sheetTables.map((table) => table.name.trim().toUpperCase()));
    let tableIndex = sheet.sheetTables.length + 1;
    while (usedNames.has(`TABLE${tableIndex}`)) tableIndex += 1;
    const fieldNames = new Set<string>();
    const columns: SheetTableModel['columns'] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      const rawName = String(sheet.cells.get(range.startRow, column)?.value ?? '').trim() || `Column${column - range.startColumn + 1}`;
      let name = rawName;
      let suffix = 2;
      while (fieldNames.has(name.toUpperCase())) name = `${rawName}${suffix++}`;
      fieldNames.add(name.toUpperCase());
      const columnIndex = column - range.startColumn;
      columns.push({ id: nextId('col'), name, totalsFunction: defaultTotalsFunction(columnIndex) });
    }
    const table: SheetTableModel = {
      id: nextId('sheet-table'),
      sheetId: this.activeSheetId,
      name: `Table${tableIndex}`,
      range,
      hasHeaderRow: true,
      hasTotalRow: false,
      showBandedRows: true,
      showBandedColumns: false,
      showFilterButton: true,
      columns,
    };
    this.runCommand('sheetTable.add', table);
    if (table.showFilterButton) {
      this.runCommand('sheet.filter.set', { sheetId: this.activeSheetId, filter: createFilterModelForTable(table) });
    }
    this.notify(`Sheet table ${table.name} created`);
    this.refresh();
  }

  toggleSheetTableTotalRow(tableId?: string, enabled?: boolean): void {
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
    this.runCommand('sheetTable.toggleTotalRow', { sheetId: this.activeSheetId, tableId: table.id, enabled: nextEnabled });
    this.notify(nextEnabled ? `Total row added to ${table.name}` : `Total row removed from ${table.name}`);
    this.refresh();
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

  textToColumnsFromSelection(delimiter = ','): void {
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
    this.runCommand('data.textToColumns', {
      sheetId: this.activeSheetId,
      range: targetRange,
      delimiter,
      maxColumns: Math.min(8, Math.max(2, sheet.columnCount - column)),
    });
    this.notify('Text split into columns');
    this.refresh();
  }

  applyDataSubtotal(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow || range.endColumn <= range.startColumn) {
      this.notify('Select a data range with at least two columns');
      return;
    }
    this.runCommand('data.subtotal', {
      sheetId: this.activeSheetId,
      range,
      groupColumn: range.startColumn,
      valueColumn: range.startColumn + 1,
      functionName: 'SUM',
    });
    this.notify('Subtotal summary created below selection');
    this.refresh();
  }

  removeDuplicatesFromSelection(): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    if (range.endRow <= range.startRow) {
      this.notify('Select a multi-row range before removing duplicates');
      return;
    }
    const columns: number[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) columns.push(column);
    this.runCommand('data.removeDuplicates', {
      sheetId: this.activeSheetId,
      range,
      columns,
      hasHeader: true,
    });
    this.notify('Duplicate rows removed');
    this.refresh();
  }

  createDataTableFromSelection(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const primaryRange = this.getPrimaryRange();
    const sourceRange = primaryRange.startRow !== primaryRange.endRow || primaryRange.startColumn !== primaryRange.endColumn ? primaryRange : usedRangeOfSheet(sheet);
    const fieldNames = new Set<string>();
    const fields: WorkbookTableModel['fields'] = [];
    for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column++) {
      const rawName = String(sheet.cells.get(sourceRange.startRow, column)?.value ?? '').trim() || `Column ${column - sourceRange.startColumn + 1}`;
      let name = rawName;
      let suffix = 2;
      while (fieldNames.has(name)) name = `${rawName} ${suffix++}`;
      fieldNames.add(name);
      const sample: CellData['value'][] = [];
      for (let row = sourceRange.startRow + 1; row <= Math.min(sourceRange.endRow, sourceRange.startRow + 1000); row++) sample.push(sheet.cells.get(row, column)?.value ?? null);
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
