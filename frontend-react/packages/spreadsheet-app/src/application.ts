import type {
  CellData,
  CellStyle,
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  PivotFieldDefinition,
  PivotLayout,
  PivotModel,
  PivotSourceRowPath,
  PivotSlicer,
  PivotTimeline,
  PivotAggregateFunction,
  RangeRef,
  ShapeModel,
  FloatingImage,
  SheetTableModel,
  SparklineModel,
  SparklineGroup,
  WorkbookTableModel,
} from '@react-sheets/core-model';
import type { HistoryEntry, CommandResult } from '@react-sheets/command-runtime';
import type { RevisionRecord, TableRowsResponse } from '@react-sheets/protocol';
import type { PrintLayout } from '@react-sheets/pro-features';
import { getPivotFieldCatalog as buildPivotFieldCatalog, computePivotResult } from '@react-sheets/pro-features';
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
  validateDataInput,
  validationList,
  type ClipboardData,
  type GoToSpecialKind,
  type PasteMode,
} from '@react-sheets/sheet-features';
import { isSpillChild, type RecalculationMode } from '@react-sheets/formula-engine';
import { EditSession } from './edit-session';
import { executeUiCommand, isUiCommand } from './execute-command';
import {
  buildCollaborationSnapshot,
  type CollaborationSnapshot,
} from './collaboration-bridge';
import { PermissionService, type PermissionCapabilities, type ShareRole } from './permission-service';
import {
  canExecuteCommand,
  findProtectionRuleCoveringRange,
} from './permission-bridge';
import {
  createSpreadsheetRuntime,
  hydrateRuntime,
  rehydrateFormulaAfterRestore,
  resolveActorId,
  resolveUnitId,
  startCollaborationSession,
  startPersistenceSession,
  type SpreadsheetRuntime,
} from './runtime';
import { createInitialSelection, SelectionService, parseRangeReference, type SelectionState } from './selection-service';
import { columnLabel } from './address';
import { buildAllSheetSnapshots, type CanvasSheetSnapshot } from './ui-snapshot';
import { syncWorkbookSheetTables, syncWorkbookSpills } from './formula-spill-sync';
import {
  buildImageDrawingAdd,
  buildShapeDrawingAdd,
  findDrawingByPayloadId,
  resolveDrawingMoveTransform,
} from './drawing-bridge';
import {
  buildChartInsertParams,
  buildChartMetadataPatch,
  resolveChartInsertCommandId,
} from './chart-bridge';
import {
  buildPivotModel,
  connectedPivotIdsForSource,
} from './pivot-bridge';
import {
  buildCellNote,
  buildCommentReply,
  buildCommentThread,
  findCommentThreadAt,
  parseUrlHyperlink,
} from './review-bridge';
import {
  buildSparklineDataLocationParams,
  buildSparklineGroup,
  buildSparklineInsertParams,
  resolveQuickSparklinePlacement,
} from './sparkline-bridge';
import {
  buildRestoreParams,
  revisionToHistoryMeta,
} from './history-bridge';
import {
  buildPersistenceMeta,
  buildLocalDraftRecord,
  type PersistenceSnapshotMeta,
} from './persistence-bridge';
import {
  exchangeExportXlsx,
  summarizeCompatibilityReport,
} from './xlsx-bridge';
import {
  buildPrintSnapshot,
  summarizePrintSnapshot,
  type PrintPageSnapshot,
  type PrintSnapshot,
} from './print-bridge';
import { browserPrintHook, PdfExportService } from './features/print';
import type { LoadTarget, QueryDefinition } from './features/query/query-steps';
import {
  buildQueryResultSnapshot,
  executeQueryDefinition,
  resolveLoadTarget,
  summarizeQueryResult,
  type QueryResultSnapshot,
  type QuerySessionEntry,
} from './query-bridge';
import {
  createCommandRecorder,
  SAMPLE_AUTOMATION_SCRIPT,
  summarizeScriptResult,
  type AutomationSnapshot,
} from './automation-bridge';
import { CommandRecorder } from './features/automation/command-recorder';
import type { ScriptRunResult } from './features/automation';
import type { CapabilityDescriptor, PlatformCapability } from './features/extended';
import type {
  DataTableParams,
  DataTableResult,
  GoalSeekParams,
  GoalSeekResult,
  ScenarioDefinition,
  ScenarioResult,
} from './features/extended/what-if';
import type { WhatIfPlan } from './features/extended';
import type { CompatibilityReport } from './xlsx-bridge';
import {
  HistoryPreviewSession,
  type HistoryEntryMeta,
} from './features/history';
import { inferTableFieldType, nextId, usedRangeOfSheet } from './application-helpers';
import type { AppPhase, PeerCursor, RibbonTabId, SaveState, SidebarPanelId } from './types';

const UNIT_ID_STORAGE_KEY = 'react-sheets:unitId';

export interface SpreadsheetApplicationOptions {
  initialPhase?: AppPhase;
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
  peers: PeerCursor[];
  collabStatus: 'connecting' | 'open' | 'closed';
  collabRevision: number;
  pendingChangeSetCount: number;
  offlineQueueState: string;
  actorId: string;
  shareRole: ShareRole;
  permissions: PermissionCapabilities;
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  historyPreviewRevision: number | null;
  hasLocalDraft: boolean;
  persistenceChecksum: string;
  compatibilityReport: CompatibilityReport | null;
  tables: readonly WorkbookTableModel[];
  showFunctionWizard: boolean;
  showSortDialog: boolean;
  showFindReplace: boolean;
  showGoTo: boolean;
  showPasteSpecial: boolean;
  showFormatCells: boolean;
  showShiftCells: boolean;
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
  platformCapabilities: readonly CapabilityDescriptor[];
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
  version: number;
}

export interface ExtendedSnapshot {
  capabilities: readonly CapabilityDescriptor[];
  lastWhatIfResult: GoalSeekResult | ScenarioResult | DataTableResult | null;
}

export function getInitialAppPhase(): AppPhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'loading' || queryPhase === 'error' || queryPhase === 'empty' ? queryPhase : 'ready';
}

export class SpreadsheetApplication {
  private readonly runtime: SpreadsheetRuntime;
  private readonly permission = new PermissionService();
  private readonly editSession = new EditSession();
  private readonly listeners = new Set<() => void>();
  private readonly actorId: string;

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
  private peers: PeerCursor[] = [];
  private collabStatus: 'connecting' | 'open' | 'closed' = 'closed';
  private remoteRevisions: RevisionRecord[] = [];
  private historyPreview: HistoryPreviewSession | null = null;
  private hasLocalDraft = false;
  private persistenceChecksum = '';
  private compatibilityReport: CompatibilityReport | null = null;
  private showFunctionWizard = false;
  private showSortDialog = false;
  private showFindReplace = false;
  private showGoTo = false;
  private showPasteSpecial = false;
  private showFormatCells = false;
  private showShiftCells = false;
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

  private selectionService: SelectionService;
  private collabDispose: (() => void) | null = null;
  private persistenceDispose: (() => void) | null = null;
  private overrideTarget: { row: number; column: number } | null = null;
  private clipboardData: ClipboardData | null = null;
  private snapshotGeneration = 0;
  private cachedUiSnapshot: UiSnapshot | null = null;
  private cachedUiSnapshotGeneration = -1;

  constructor({ initialPhase = 'ready' }: SpreadsheetApplicationOptions = {}) {
    this.runtime = createSpreadsheetRuntime();
    this.actorId = resolveActorId();
    this.phase = initialPhase;
    this.activeSheetId = this.runtime.model.activeSheetId;
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

  private wireRuntimeHandlers(): void {
    this.runtime.handlers.onSaveState = (state) => {
      this.saveState = state;
      this.emit();
    };
    this.runtime.handlers.onNotice = (message) => {
      this.notice = message;
      this.emit();
    };
    this.runtime.handlers.onMutationsApplied = () => this.refresh();
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
    this.runtime.handlers.onDraftUpdated = () => {
      this.syncPersistenceMeta();
    };
    this.runtime.handlers.onCollabStatus = (status) => {
      this.collabStatus = status;
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
      `${this.activeSheetId}:${this.selectionService.getState().primaryRowIndex}:${this.selectionService.getState().primaryColumnIndex}`,
      this.runtime.authTokenProvider,
    );
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
    const draft = this.runtime.draftStore.read(this.runtime.model.unitId);
    const meta = buildPersistenceMeta(this.runtime.model.snapshot(), this.runtime.remoteRevision, draft);
    this.hasLocalDraft = meta.hasLocalDraft;
    this.persistenceChecksum = meta.checksum;
  }

  getUiSnapshot = (): UiSnapshot => {
    if (this.cachedUiSnapshot && this.cachedUiSnapshotGeneration === this.snapshotGeneration) {
      return this.cachedUiSnapshot;
    }
    const sheets = buildAllSheetSnapshots(this.runtime.model, this.runtime.formula, this.runtime.pivotResults);
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
      peers: this.peers,
      collabStatus: this.collabStatus,
      collabRevision: collaboration.revision,
      pendingChangeSetCount: collaboration.pendingCount,
      offlineQueueState: collaboration.offlineQueueState,
      actorId: this.actorId,
      shareRole: this.getShareRole(),
      permissions: this.permission.getCapabilities(this.actorId),
      historyEntries: this.runtime.commands.getUndoEntries(),
      remoteRevisions: this.remoteRevisions,
      historyPreviewRevision: this.historyPreview?.revision ?? null,
      hasLocalDraft: this.hasLocalDraft,
      persistenceChecksum: this.persistenceChecksum,
      compatibilityReport: this.compatibilityReport,
      tables: [...this.runtime.model.tables.values()].map((table) => structuredClone(table)),
      showFunctionWizard: this.showFunctionWizard,
      showSortDialog: this.showSortDialog,
      showFindReplace: this.showFindReplace,
      showGoTo: this.showGoTo,
      showPasteSpecial: this.showPasteSpecial,
      showFormatCells: this.showFormatCells,
      showShiftCells: this.showShiftCells,
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
      platformCapabilities: this.runtime.capabilities.list(),
      lastWhatIfResult: this.lastWhatIfResult,
      version: this.version,
    };
    this.cachedUiSnapshotGeneration = this.snapshotGeneration;
    this.cachedUiSnapshot = snapshot;
    return snapshot;
  };

  execute(commandId: string, params?: unknown): void {
    if (this.phase !== 'ready' && !commandId.startsWith('ui.')) return;
    try {
      this.assertPermission(commandId, params);
      if (isUiCommand(commandId)) {
        if (executeUiCommand(this, commandId, params)) {
          this.refresh();
        }
        return;
      }
      this.runCommand(commandId, params);
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Permission denied');
    }
  }

  runCommand(commandId: string, params?: unknown): CommandResult {
    this.assertPermission(commandId, params);
    const result = this.runtime.commands.execute(commandId, params);
    if (commandId === 'history.restore') {
      const restoreParams = params as { targetRevision?: number };
      rehydrateFormulaAfterRestore(this.runtime, restoreParams.targetRevision);
      this.activeSheetId = this.runtime.model.activeSheetId;
      this.selectionService.resetForSheet(this.activeSheetId);
      this.clearHistoryPreview();
    }
    this.applySelectionFromCommand(commandId, params, result);
    this.refresh();
    return result;
  }

  canExecute(commandId: string, params?: unknown): boolean {
    return canExecuteCommand(
      this.permission,
      this.runtime.model,
      commandId,
      params,
      this.actorId,
      this.activeSheetId,
    ).allowed;
  }

  getShareRole(): ShareRole {
    return this.permission.getShareRole(this.actorId);
  }

  setShareRole(role: ShareRole): void {
    this.permission.setShareRole(this.actorId, role);
    this.notify(`Share role set to ${role}`);
    this.refresh();
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
      this.selectionService.applyFromRanges(selectionParams.ranges);
      const primaryIndex = selectionParams.primaryRangeIndex ?? 0;
      const primaryRange = selectionParams.ranges[primaryIndex] ?? selectionParams.ranges[0]!;
      const primary = selectionParams.primaryCell ?? { row: primaryRange.startRow, column: primaryRange.startColumn };
      this.selectionService.setPrimaryCell(primary.row, primary.column);
      if (selectionParams.anchorCell) {
        this.selectionService.setAnchor(selectionParams.anchorCell.row, selectionParams.anchorCell.column);
      }
      this.syncDraftFromPrimary();
      return;
    }
    if (!SpreadsheetApplication.SELECTION_COMMAND_IDS.has(commandId)) return;
    if (result.affectedRanges.length === 0) return;
    this.selectionService.applyFromRanges(result.affectedRanges);
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

  getWorkbook() {
    return this.runtime.model;
  }

  getActiveSheetId(): string {
    return this.activeSheetId;
  }

  getSelection(): SelectionState {
    return this.selectionService.getState();
  }

  getPrimaryRange(): RangeRef {
    return this.selectionService.primaryRangeOrDefault();
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

  restoreFromSnapshot(snapshot: import('@react-sheets/core-model').WorkbookSnapshotV1, targetRevision: number, reason?: string): void {
    try {
      this.runCommand('history.restore', buildRestoreParams(snapshot, targetRevision, reason));
      this.notify(`Restored workbook to revision ${targetRevision}`);
    } catch (error) {
      this.notify(error instanceof Error ? error.message : 'Permission denied');
    }
  }

  async restoreToRevision(revision: number): Promise<void> {
    const response = await this.runtime.api.getRevisionSnapshot(this.runtime.model.unitId, revision);
    this.restoreFromSnapshot(response.snapshot, revision, `Restore to revision ${revision}`);
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
      this.historyPreview = HistoryPreviewSession.fromSnapshot(meta, response.snapshot);
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
    const draft = this.runtime.draftStore.read(this.runtime.model.unitId);
    return buildPersistenceMeta(this.runtime.model.snapshot(), this.runtime.remoteRevision, draft);
  }

  async saveWorkbook(reason = 'Manual save'): Promise<void> {
    if (!this.canExecute('persistence.save')) {
      this.notify('You do not have permission to save this workbook');
      return;
    }
    this.saveState = 'saving';
    this.emit();
    try {
      this.runCommand('persistence.save', { reason, baseRevision: this.runtime.remoteRevision });
      const response = await this.runtime.api.saveSnapshot(
        this.runtime.model.unitId,
        this.runtime.model.snapshot(),
        this.runtime.remoteRevision,
      );
      this.runtime.remoteRevision = response.revision;
      this.runtime.collaboration?.setRevision(response.revision);
      this.runtime.draftStore.clear(this.runtime.model.unitId);
      this.runCommand('persistence.draft.clear');
      this.saveState = 'saved';
      this.syncPersistenceMeta();
      this.notify('Workbook saved');
      await this.refreshRevisionLog();
    } catch (error) {
      this.saveState = error instanceof Error && error.message.includes('conflict') ? 'conflict' : 'offline';
      this.notify(error instanceof Error ? error.message : 'Save failed');
      this.emit();
    }
  }

  recoverLocalDraft(): boolean {
    const draft = this.runtime.draftStore.read(this.runtime.model.unitId);
    if (!draft) {
      this.notify('No local draft available');
      return false;
    }
    hydrateRuntime(this.runtime, { snapshot: draft.snapshot, revision: draft.revision });
    this.activeSheetId = this.runtime.model.activeSheetId;
    this.selectionService.resetForSheet(this.activeSheetId);
    this.saveState = 'offline';
    this.syncPersistenceMeta();
    this.notify('Local draft recovered');
    this.refresh();
    return true;
  }

  clearLocalDraft(): void {
    this.runtime.draftStore.clear(this.runtime.model.unitId);
    this.runCommand('persistence.draft.clear');
    this.syncPersistenceMeta();
    this.notify('Local draft cleared');
    this.refresh();
  }

  notify(message: string): void {
    this.notice = message;
    this.emit();
  }

  undo(): void {
    if (this.runtime.commands.undo()) {
      this.syncDraftFromPrimary();
      this.notify('Undo applied');
      this.refresh();
    }
  }

  redo(): void {
    if (this.runtime.commands.redo()) {
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

  openDialog(dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells', findQuery?: string): void {
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
  pasteSpecial(mode: PasteMode): void {
    this.execute('ui.clipboard.paste', { mode });
    this.closePasteSpecial();
  };
  setShowPrintPreview = (open: boolean): void => {
    this.showPrintPreview = open;
    this.emit();
  };

  syncDraftFromPrimary(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
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
        xSplit: sel.primaryColumnIndex,
        ySplit: sel.primaryRowIndex,
        startRow: sel.primaryRowIndex,
        startColumn: sel.primaryColumnIndex,
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
    let row = sel.primaryRowIndex;
    let column = sel.primaryColumnIndex;
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
    if (extend) this.movePrimary(row - sel.primaryRowIndex, column - sel.primaryColumnIndex, { extend: true });
    else {
      this.selectionService.setPrimary(row, column);
      this.syncDraftFromPrimary();
    }
    this.emit();
  }

  selectAll(): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    this.selectionService.selectAll(sheet.rowCount, sheet.columnCount);
    this.syncDraftFromPrimary();
    this.emit();
  }

  beginEdit(initialText?: string): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    for (const spill of sheet.spillRanges) {
      if (isSpillChild(spill, sel.primaryRowIndex, sel.primaryColumnIndex)) {
        this.notify('Spill cells are read-only');
        return;
      }
    }
    const cell = sheet.cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    this.editSession.begin({
      sheetId: this.activeSheetId,
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
      cell,
      selection: this.selectionService.getSnapshot(),
      initialText,
    });
    this.formulaDraft = this.editSession.active?.currentDraft ?? '';
    this.emit();
    if (typeof document !== 'undefined' && initialText === undefined) {
      queueMicrotask(() => {
        const input = document.querySelector<HTMLInputElement>('[data-testid="formula-input"]');
        input?.focus();
        const length = input?.value.length ?? 0;
        input?.setSelectionRange(length, length);
      });
    }
  }

  cancelEdit(): void {
    this.formulaDraft = this.editSession.cancel();
    this.emit();
  }

  commitFormula(overrideValue?: string): boolean {
    if (this.phase !== 'ready') return false;
    const sel = this.selectionService.getState();
    const row = this.overrideTarget?.row ?? sel.primaryRowIndex;
    const column = this.overrideTarget?.column ?? sel.primaryColumnIndex;
    const raw = (overrideValue !== undefined ? overrideValue : this.formulaDraft).trim();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const existingStyle = sheet.cells.get(row, column)?.style;
    const isFormula = raw.startsWith('=');
    const value = isFormula ? null : raw === '' ? null : Number.isFinite(Number(raw)) ? Number(raw) : raw;
    const validation = validateDataInput(sheet, row, column, value);
    if (!validation.valid) {
      if (validation.blocking) {
        this.notify(validation.message ?? '输入不符合数据验证规则');
        return false;
      }
      this.notify(`警告: ${validation.message ?? '数据验证未通过'}`);
    }
    const cellData: CellData = isFormula ? { value: null, formula: raw, style: existingStyle } : { value, formula: undefined, style: existingStyle };
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row, column, value: cellData });
    return true;
  }

  commitEdit(moveAfter: 'down' | 'up' | 'left' | 'right' | 'none' = 'down'): void {
    if (!this.editSession.editingCell) return;
    const editing = this.editSession.editingCell;
    this.overrideTarget = editing;
    const committed = this.commitFormula(this.formulaDraft.trim());
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
    } else {
      this.setFormulaDraft(
        this.formulaDraft.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (_m, dCol, col, dRow, row) => `${dCol ? '' : '$'}${col}${dRow ? '' : '$'}${row}`),
      );
    }
    this.emit();
  }

  selectSheet(sheetId: string): void {
    this.runtime.model.activeSheetId = sheetId;
    this.activeSheetId = sheetId;
    this.selectionService.resetForSheet(sheetId);
    this.editSession.cancel();
    this.formulaDraft = '';
    this.refresh();
  }

  // ---- Pro / data features ----

  addChart(chart: ChartModel): void {
    const drawingId = nextId('draw');
    const insertCommand = resolveChartInsertCommandId(chart.type);
    if (insertCommand === 'chart.insert') {
      this.runCommand('chart.insert', buildChartInsertParams(chart, drawingId));
    } else {
      this.runCommand(insertCommand, {
        sheetId: chart.sheetId,
        chartId: chart.id,
        drawingId,
        bounds: chart.bounds,
        sourceRanges: chart.sourceRanges,
        title: chart.title,
      });
    }
    const metadataPatch = buildChartMetadataPatch(chart);
    if (metadataPatch) {
      this.runCommand('chart.update', { sheetId: chart.sheetId, chartId: chart.id, payload: metadataPatch });
    }
    this.runCommand('drawing.select', { sheetId: chart.sheetId, drawingIds: [drawingId] });
    this.selectedFloatingId = chart.id;
    this.notify(chart.title ? `Added chart "${chart.title}"` : `Added ${chart.type} chart`);
    this.refresh();
  }
  insertQuickChart(type: ChartModel['type'] = 'column'): void {
    const range = normalizeRangeRef(this.getPrimaryRange());
    this.addChart({
      id: nextId('chart'),
      sheetId: this.activeSheetId,
      type,
      title: 'Chart',
      sourceRanges: [{ ...range, sheetId: this.activeSheetId }],
      bounds: { x: 96, y: 96, width: 480, height: 280 },
    });
  }
  updateChartType(chartId: string, chartType: ChartModel['type']): void {
    this.runCommand('chart.setType', { sheetId: this.activeSheetId, chartId, chartType });
    this.refresh();
  }
  updateChartSeries(chartId: string, sourceRanges: RangeRef[], series?: ChartModel['series'], categoryRange?: RangeRef): void {
    this.runCommand('chart.setSeries', {
      sheetId: this.activeSheetId,
      chartId,
      sourceRanges,
      series: series?.map((entry) => ({ name: entry.name, range: entry.range, color: entry.color })),
      categoryRange,
    });
    this.refresh();
  }
  setChartLegend(chartId: string, legendPosition: NonNullable<ChartModel['legendPosition']>): void {
    this.runCommand('chart.setLegend', { sheetId: this.activeSheetId, chartId, legendPosition });
    this.refresh();
  }
  setChartDataLabels(chartId: string, showDataLabels: boolean): void {
    this.runCommand('chart.setDataLabels', { sheetId: this.activeSheetId, chartId, showDataLabels });
    this.refresh();
  }
  updateChartBounds(id: string, bounds: ChartModel['bounds']): void {
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
    if (this.selectedFloatingId === id) this.selectedFloatingId = null;
    this.refresh();
  }
  addPivot(pivot: PivotModel): void {
    this.runCommand('pivot.add', pivot);
    this.recomputePivotResult(pivot.id);
    this.notify(`Pivot ${pivot.id} added`);
    this.refresh();
  }
  insertQuickPivot(): string | undefined {
    const range = normalizeRangeRef(this.getPrimaryRange());
    const pivotId = nextId('pivot');
    const pivot = buildPivotModel(this.runtime.model, this.activeSheetId, pivotId, range);
    if (!pivot) {
      this.notify('Select a data range with category and value fields');
      return undefined;
    }
    this.addPivot(pivot);
    return pivotId;
  }
  updatePivotLayout(pivotId: string, layout: PivotLayout): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, layout });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  updatePivotConfiguration(
    pivotId: string,
    patch: Parameters<SpreadsheetApplication['updatePivotLayout']>[1] extends PivotLayout
      ? { sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotModel['slicers']; timelines?: PivotModel['timelines']; chartReferences?: PivotModel['chartReferences'] }
      : never,
  ): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, ...patch });
    if (patch.sourceRange) {
      const pivot = this.runtime.model.getSheet(this.activeSheetId).pivots.find((entry) => entry.id === pivotId);
      if (pivot) pivot.fieldCatalog = undefined;
    }
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  setPivotAggregate(pivotId: string, field: string, summarizeBy: PivotAggregateFunction): void {
    this.runCommand('pivot.setAggregate', { sheetId: this.activeSheetId, pivotId, field, summarizeBy });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  setPivotSlicer(pivotId: string, slicer: PivotSlicer): void {
    this.runCommand('pivot.slicer.set', { sheetId: this.activeSheetId, pivotId, slicer });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  setPivotTimeline(pivotId: string, timeline: PivotTimeline): void {
    this.runCommand('pivot.timeline.set', { sheetId: this.activeSheetId, pivotId, timeline });
    this.recomputePivotResult(pivotId);
    this.refresh();
  }
  refreshPivot(pivotId: string): void {
    const pivot = this.runtime.model.getSheet(this.activeSheetId).pivots.find((entry) => entry.id === pivotId);
    if (!pivot) return;
    this.runCommand('pivot.refresh', { sheetId: this.activeSheetId, pivotId });
    this.recomputePivotResult(pivotId);
    this.notify('Pivot refreshed');
    this.refresh();
  }
  removePivot(id: string): void {
    this.runCommand('pivot.remove', id);
    delete this.runtime.pivotResults[id];
    this.refresh();
  }
  drillDownPivot(pivotId: string, label: string, paths: PivotSourceRowPath[]): void {
    if (paths.length === 0) return;
    const targetSheetId = nextId('sheet');
    this.runCommand('pivot.drillDown', {
      sheetId: this.activeSheetId,
      pivotId,
      label,
      sourceRowPaths: paths.map((path) => ({ sheetId: path.sheetId, row: path.row })),
      targetSheetId,
      targetAnchor: { row: 0, column: 0 },
    });
    this.activeSheetId = targetSheetId;
    this.notify(`Drill-down sheet created for ${label}`);
    this.refresh();
  }
  private recomputePivotResult(pivotId: string): void {
    const pivot = this.runtime.model.getSheet(this.activeSheetId).pivots.find((entry) => entry.id === pivotId);
    if (!pivot) {
      delete this.runtime.pivotResults[pivotId];
      return;
    }
    try {
      this.runtime.pivotResults[pivotId] = computePivotResult(this.runtime.model, pivot);
    } catch {
      delete this.runtime.pivotResults[pivotId];
    }
  }
  addShape(shape: ShapeModel): void {
    const drawingId = nextId('draw');
    this.runCommand('drawing.add.shape', buildShapeDrawingAdd(this.activeSheetId, shape, drawingId));
    this.runCommand('drawing.select', { sheetId: this.activeSheetId, drawingIds: [drawingId] });
    this.selectedFloatingId = shape.id;
    this.notify(`Added ${shape.type} shape`);
    this.refresh();
  }
  insertQuickShape(type: ShapeModel['type'] = 'rounded-rectangle'): void {
    const shape: ShapeModel = {
      id: nextId('shape'),
      sheetId: this.activeSheetId,
      type,
      text: type === 'callout' ? 'Note' : '',
      fill: '#dbeafe',
      stroke: '#2563eb',
      strokeWidth: 2,
      textColor: '#1e3a8a',
      fontSize: 13,
      bounds: { x: 96, y: 96, width: 160, height: 60 },
    };
    this.addShape(shape);
  }
  updateShapeBounds(id: string, bounds: ShapeModel['bounds']): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const move = resolveDrawingMoveTransform(sheet, id, bounds);
    if (move) {
      this.runCommand('drawing.move', { sheetId: this.activeSheetId, drawingId: move.drawingId, transform: move.transform });
      this.refresh();
      return;
    }
    this.runCommand('shape.move', { id, sheetId: this.activeSheetId, bounds });
    this.refresh();
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
    if (this.selectedFloatingId === id) this.selectedFloatingId = null;
    this.refresh();
  }
  addImage(image: FloatingImage): void {
    const drawingId = nextId('draw');
    this.runCommand('drawing.add.image', buildImageDrawingAdd(this.activeSheetId, image, drawingId));
    this.runCommand('drawing.select', { sheetId: this.activeSheetId, drawingIds: [drawingId] });
    this.selectedFloatingId = image.id;
    this.notify('Image placed on canvas');
    this.refresh();
  }
  updateImageBounds(id: string, bounds: FloatingImage['bounds']): void {
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
    if (this.selectedFloatingId === id) this.selectedFloatingId = null;
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
    const chart = sheet.charts.find((entry) => entry.id === this.selectedFloatingId);
    const image = sheet.images.find((entry) => entry.id === this.selectedFloatingId);
    if (chart) this.removeChart(chart.id);
    else if (image) this.removeImage(image.id);
    else this.removeShape(this.selectedFloatingId);
  }
  private resolveSelectedDrawingId(): string | undefined {
    if (!this.selectedFloatingId) return undefined;
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    return findDrawingByPayloadId(sheet, this.selectedFloatingId)?.id;
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
    this.runCommand('sheet.cf.add', { rule });
  }
  removeConditionalFormat(ruleId: string): void {
    this.runCommand('sheet.cf.remove', { sheetId: this.activeSheetId, ruleId });
  }
  addDataValidation(rule: DataValidationRule): void {
    this.runCommand('sheet.dv.add', { rule });
  }
  removeDataValidation(ruleId: string): void {
    this.runCommand('sheet.dv.remove', { sheetId: this.activeSheetId, ruleId });
  }

  addComment(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const thread = buildCommentThread(
      this.activeSheetId,
      sel.primaryRowIndex,
      sel.primaryColumnIndex,
      this.actorId,
      text,
      nextId('thread'),
    );
    this.runCommand('comment.add', {
      sheetId: this.activeSheetId,
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
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
    this.runCommand('sheet.rows.insert', { sheetId: this.activeSheetId, at: this.selectionService.getState().primaryRowIndex, count });
  }
  deleteRowsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.runCommand('sheet.rows.delete', { sheetId: this.activeSheetId, at: range?.startRow ?? sel.primaryRowIndex, count: (range?.endRow ?? sel.primaryRowIndex) - (range?.startRow ?? sel.primaryRowIndex) + 1 });
  }
  insertColumnsAtPrimary(count: number): void {
    this.runCommand('sheet.columns.insert', { sheetId: this.activeSheetId, at: this.selectionService.getState().primaryColumnIndex, count });
  }
  deleteColumnsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    this.runCommand('sheet.columns.delete', { sheetId: this.activeSheetId, at: range?.startColumn ?? sel.primaryColumnIndex, count: (range?.endColumn ?? sel.primaryColumnIndex) - (range?.startColumn ?? sel.primaryColumnIndex) + 1 });
  }
  hideRowsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    const start = range?.startRow ?? sel.primaryRowIndex;
    const end = range?.endRow ?? sel.primaryRowIndex;
    for (let row = start; row <= end; row++) this.runCommand('sheet.row.hide', { sheetId: this.activeSheetId, index: row });
  }
  hideColumnsAtPrimary(): void {
    const sel = this.selectionService.getState();
    const range = sel.ranges[sel.primaryRangeIndex];
    const start = range?.startColumn ?? sel.primaryColumnIndex;
    const end = range?.endColumn ?? sel.primaryColumnIndex;
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
    this.runCommand('sheet.splitColumn', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, delimiter, maxColumns: Math.min(sheet.columnCount - sel.primaryColumnIndex - 1, 8) });
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
  paste(): void {
    const sel = this.selectionService.getState();
    const internal = this.clipboardData;
    if (internal) {
      this.runCommand('sheet.range.paste', {
        sheetId: this.activeSheetId,
        targetOrigin: { row: sel.primaryRowIndex, column: sel.primaryColumnIndex },
        clipboard: internal,
        mode: 'all',
      });
      if (internal.isCut) {
        this.runCommand('sheet.range.clear', { sheetId: this.activeSheetId, range: internal.range, mode: 'contents' });
        this.clearClipboard();
      }
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
        targetOrigin: { row: sel.primaryRowIndex, column: sel.primaryColumnIndex },
        clipboard,
        mode: 'all',
      });
      this.syncDraftFromPrimary();
      this.notify('Pasted from clipboard');
    });
  }
  clearFormats(): void {
    this.runCommand('sheet.range.clear', { sheetId: this.activeSheetId, range: this.getPrimaryRange(), mode: 'formats' });
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
    this.activeSheetId = id;
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
    this.activeSheetId = newId;
    this.selectionService.resetForSheet(newId);
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
        if (remaining) this.activeSheetId = remaining.id;
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
      const drawing = findDrawingByPayloadId(sheet, id);
      if (drawing) {
        this.runCommand('drawing.select', { sheetId: this.activeSheetId, drawingIds: [drawing.id] });
      } else {
        this.runCommand('drawing.deselect', { sheetId: this.activeSheetId });
      }
    } else {
      this.runCommand('drawing.deselect', { sheetId: this.activeSheetId });
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
      sheetId: range.sheetId,
      sourceRange: range,
      layout: { rows: [], columns: [], filters: [], values: [], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false, calculatedFields: [], calculatedItems: [] },
    };
    return buildPivotFieldCatalog(this.runtime.model, pivot).fields;
  }

  readDataTable(tableId: string, offset = 0, limit = 100): Promise<TableRowsResponse> {
    return this.runtime.api.readDataRows(this.runtime.model.unitId, tableId, offset, limit);
  }
  removeDataTable(tableId: string): Promise<void> {
    const table = this.runtime.model.tables.get(tableId);
    if (!table) return Promise.reject(new Error('Data table not found'));
    return this.runtime.api.deleteDataTable(this.runtime.model.unitId, tableId).then(() => {
      this.runCommand('table.remove', { tableId, sheetId: table.sourceSheetId ?? this.activeSheetId });
    });
  }

  showPivotDetails(pivotId: string, paths: PivotSourceRowPath[], label = 'Details'): void {
    this.drillDownPivot(pivotId, label, paths);
  }

  getValidationForPrimary(): DataValidationRule | undefined {
    const sel = this.selectionService.getState();
    return findValidationRule(this.runtime.model.getSheet(this.activeSheetId), sel.primaryRowIndex, sel.primaryColumnIndex);
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
      await service.export(snapshot.model, snapshot.pages, {
        filename: `${this.runtime.model.name || 'workbook'}.pdf`,
        title: this.runtime.model.name,
      });
      this.notify('Choose Save as PDF in the print dialog');
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
      const result = await executeQueryDefinition(this.runtime.connectors, query);
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
      const result = await executeQueryDefinition(this.runtime.connectors, session.definition);
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

  runAutomationScript(source: string): void {
    if (!this.canExecute('automation.run')) {
      this.notify('You do not have permission to run scripts');
      return;
    }
    const started = Date.now();
    try {
      const result = this.runCommand('automation.run', { source });
      this.lastScriptResult = { ok: true, durationMs: Date.now() - started, mutationCount: result.mutationCount };
    } catch (error) {
      this.lastScriptResult = {
        ok: false,
        durationMs: Date.now() - started,
        error: error instanceof Error ? error.message : 'Automation failed',
      };
    }
    this.activePanel = 'automate';
    this.ribbonTab = 'automate';
    this.notify(summarizeScriptResult(this.lastScriptResult));
    this.refresh();
  }

  runSampleAutomationScript(): void {
    this.runAutomationScript(SAMPLE_AUTOMATION_SCRIPT);
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
    this.runCommand('automation.record.stop', {});
    this.recorderDetach?.();
    this.recorderDetach = null;
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
    if (!this.runtime.capabilities.isEnabled('what-if')) {
      this.notify('What-if analysis is disabled for this workbook');
      return {
        kind: 'goal-seek',
        status: 'not-converged',
        iterations: 0,
        message: 'What-if capability disabled',
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
    if (!this.runtime.capabilities.isEnabled('what-if')) {
      this.notify('What-if analysis is disabled for this workbook');
      return {
        kind: 'scenario',
        status: 'failed',
        scenarioId: scenario.id,
        message: 'What-if capability disabled',
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
    if (!this.runtime.capabilities.isEnabled('what-if')) {
      this.notify('What-if analysis is disabled for this workbook');
      return { kind: 'data-table', status: 'failed', message: 'What-if capability disabled', filledCells: 0, writes: [] };
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

  evaluatePlatformCapability(capability: PlatformCapability): { canEnable: boolean; reason?: string } {
    this.runCommand('extended.capability.evaluate', { capability });
    return this.runtime.capabilities.evaluate(capability);
  }

  getExtendedSnapshot(): ExtendedSnapshot {
    return {
      capabilities: this.runtime.capabilities.list(),
      lastWhatIfResult: this.lastWhatIfResult,
    };
  }

  async importXlsxBase64(base64: string, fileName = 'import.xlsx'): Promise<void> {
    if (!this.canExecute('xlsx.import')) {
      this.notify('You do not have permission to import workbooks');
      return;
    }
    this.phase = 'loading';
    this.emit();
    try {
      this.runCommand('xlsx.import', { fileName, base64 });
      const response = await this.runtime.api.importXlsxBase64(base64, fileName);
      hydrateRuntime(this.runtime, response);
      this.compatibilityReport = response.report as CompatibilityReport;
      this.runtime.remoteConnected = true;
      if (typeof window !== 'undefined') {
        window.localStorage.setItem(UNIT_ID_STORAGE_KEY, response.snapshot.unitId);
        window.history.replaceState({}, '', `/workbooks/${encodeURIComponent(response.snapshot.unitId)}`);
      }
      this.activeSheetId = this.runtime.model.activeSheetId;
      this.selectionService.resetForSheet(this.activeSheetId);
      this.runtime.draftStore.clear(this.runtime.model.unitId);
      this.phase = 'ready';
      this.notify(summarizeCompatibilityReport(this.compatibilityReport));
      this.refresh();
    } catch (error) {
      this.phase = 'error';
      this.notify(error instanceof Error ? error.message : 'XLSX import failed');
      this.emit();
    }
  }

  async exportXlsxWorkbook(fileName?: string): Promise<{ base64: string; fileName: string } | null> {
    if (!this.canExecute('xlsx.export')) {
      this.notify('You do not have permission to export workbooks');
      return null;
    }
    try {
      this.runCommand('xlsx.export', { fileName });
      const exported = await exchangeExportXlsx(this.runtime.model.snapshot(), {
        fileName: fileName ?? `${this.runtime.model.name || 'workbook'}.xlsx`,
      });
      this.compatibilityReport = exported.report;
      this.notify(summarizeCompatibilityReport(exported.report));
      this.refresh();
      if (!exported.base64 || !exported.fileName) return null;
      return { base64: exported.base64, fileName: exported.fileName };
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
    this.refresh();
  }

  recalculateFormulas(): void {
    this.runtime.formula.recalculate();
    syncWorkbookSpills(this.runtime.formula, this.runtime.model);
    this.refresh();
    this.notify('Formulas recalculated');
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
    syncWorkbookSheetTables(this.runtime.formula, this.runtime.model);
    this.notify(`Sheet table ${table.name} created`);
    this.refresh();
  }

  toggleSheetTableTotalRow(tableId?: string, enabled?: boolean): void {
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const selection = this.selectionService.getState();
    const table = tableId
      ? sheet.sheetTables.find((entry) => entry.id === tableId)
      : findSheetTableAt(sheet, selection.primaryRowIndex, selection.primaryColumnIndex);
    if (!table) {
      this.notify('Select a cell inside a sheet table first');
      return;
    }
    const nextEnabled = enabled ?? !table.hasTotalRow;
    this.runCommand('sheetTable.toggleTotalRow', { sheetId: this.activeSheetId, tableId: table.id, enabled: nextEnabled });
    syncWorkbookSheetTables(this.runtime.formula, this.runtime.model);
    syncWorkbookSpills(this.runtime.formula, this.runtime.model);
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
    const column = selection.primaryColumnIndex;
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
    const table: WorkbookTableModel = { id: nextId('table'), name: `${sheet.name} table`, sourceSheetId: this.activeSheetId, rowCount: Math.max(0, sourceRange.endRow - sourceRange.startRow), fields, blockSize: 4096, blocks: [], revision: 0 };
    void (async () => {
      this.saveState = 'saving';
      this.emit();
      await this.runtime.api.createDataTable(this.runtime.model.unitId, table);
      for (let startRow = sourceRange.startRow + 1; startRow <= sourceRange.endRow; startRow += table.blockSize) {
        const rows: import('@react-sheets/core-model').TableScalar[][] = [];
        const endRow = Math.min(sourceRange.endRow, startRow + table.blockSize - 1);
        for (let row = startRow; row <= endRow; row++) rows.push(fields.map((_f, offset) => sheet.cells.get(row, sourceRange.startColumn + offset)?.value ?? null));
        if (rows.length > 0) table.blocks.push(await this.runtime.api.appendDataBlock(this.runtime.model.unitId, table.id, startRow - sourceRange.startRow - 1, rows));
      }
      table.revision = table.blocks.length;
      this.runCommand('table.add', table);
      this.saveState = 'saved';
      this.notify(`Data table ${table.name} created`);
    })().catch((error: unknown) => {
      this.saveState = 'offline';
      this.notify(error instanceof Error ? error.message : 'Data table creation failed');
      this.emit();
    });
  }

  replyComment(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!thread) return;
    const reply = buildCommentReply(this.actorId, text, nextId('reply'));
    this.runCommand('comment.reply', { sheetId: this.activeSheetId, threadId: thread.id, reply });
    this.notify('Reply added');
    this.refresh();
  }
  resolveComment(): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!thread || thread.resolved) return;
    this.runCommand('comment.resolve', { sheetId: this.activeSheetId, threadId: thread.id, resolved: true });
    this.notify('Comment resolved');
    this.refresh();
  }
  removeComment(): void {
    const sel = this.selectionService.getState();
    const sheet = this.runtime.model.getSheet(this.activeSheetId);
    const thread = findCommentThreadAt(sheet, sel.primaryRowIndex, sel.primaryColumnIndex);
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
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
      note,
    });
    this.notify('Note added');
    this.refresh();
  }
  removeNote(): void {
    const sel = this.selectionService.getState();
    this.runCommand('note.remove', {
      sheetId: this.activeSheetId,
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
    });
    this.notify('Note removed');
    this.refresh();
  }
  setNoteVisibility(visible: boolean): void {
    const sel = this.selectionService.getState();
    this.runCommand('note.visibility', {
      sheetId: this.activeSheetId,
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
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
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
      hyperlink,
    });
    this.notify('Hyperlink inserted');
    this.refresh();
  }
  removeHyperlink(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!cell?.hyperlink && !cell?.hyperlinkDetail) return;
    this.runCommand('hyperlink.remove', {
      sheetId: this.activeSheetId,
      row: sel.primaryRowIndex,
      column: sel.primaryColumnIndex,
    });
    this.notify('Hyperlink removed');
    this.refresh();
  }
}

export { resolveUnitId, resolveActorId };
