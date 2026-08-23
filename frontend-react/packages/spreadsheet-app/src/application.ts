import type {
  CellComment,
  CellData,
  CellStyle,
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  PivotFieldDefinition,
  PivotLayout,
  PivotModel,
  PivotSourceRowPath,
  RangeRef,
  ShapeModel,
  SparklineModel,
  WorkbookTableModel,
} from '@react-sheets/core-model';
import type { HistoryEntry, CommandResult } from '@react-sheets/command-runtime';
import type { RevisionRecord, TableRowsResponse } from '@react-sheets/protocol';
import type { PrintLayout } from '@react-sheets/pro-features';
import { getPivotFieldCatalog as buildPivotFieldCatalog } from '@react-sheets/pro-features';
import {
  collectFindReplacements,
  copyRangeToClipboardData,
  findValidationRule,
  formatTsv,
  normalizeRangeRef,
  parseTsv,
  validateDataInput,
  validationList,
  type ClipboardData,
  type GoToSpecialKind,
  type PasteMode,
} from '@react-sheets/sheet-features';
import { EditSession } from './edit-session';
import { executeUiCommand, isUiCommand } from './execute-command';
import { PermissionService } from './permission-service';
import {
  createSpreadsheetRuntime,
  hydrateRuntime,
  resolveActorId,
  resolveUnitId,
  startCollaborationSession,
  startPersistenceSession,
  type SpreadsheetRuntime,
} from './runtime';
import { createInitialSelection, SelectionService, parseRangeReference, type SelectionState } from './selection-service';
import { buildAllSheetSnapshots, type CanvasSheetSnapshot } from './ui-snapshot';
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
  actorId: string;
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
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
  version: number;
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
    this.collabDispose = startCollaborationSession(this.runtime, this.actorId, () =>
      `${this.activeSheetId}:${this.selectionService.getState().primaryRowIndex}:${this.selectionService.getState().primaryColumnIndex}`,
    );
  }

  dispose(): void {
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
    this.version += 1;
    this.emit();
  }

  getUiSnapshot = (): UiSnapshot => {
    if (this.cachedUiSnapshot && this.cachedUiSnapshotGeneration === this.snapshotGeneration) {
      return this.cachedUiSnapshot;
    }
    const sheets = buildAllSheetSnapshots(this.runtime.model, this.runtime.formula, this.runtime.pivotResults);
    const selectedSheet = sheets.find((sheet) => sheet.id === this.activeSheetId) ?? sheets[0]!;
    const selection = this.selectionService.getState();
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
      actorId: this.actorId,
      historyEntries: this.runtime.commands.getUndoEntries(),
      remoteRevisions: this.remoteRevisions,
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
      version: this.version,
    };
    this.cachedUiSnapshotGeneration = this.snapshotGeneration;
    this.cachedUiSnapshot = snapshot;
    return snapshot;
  };

  execute(commandId: string, params?: unknown): void {
    if (this.phase !== 'ready' && !commandId.startsWith('ui.')) return;
    this.permission.assert(commandId, params, this.actorId);
    if (isUiCommand(commandId)) {
      if (executeUiCommand(this, commandId, params)) {
        this.refresh();
      }
      return;
    }
    this.runCommand(commandId, params);
  }

  runCommand(commandId: string, params?: unknown): CommandResult {
    this.permission.assert(commandId, params, this.actorId);
    const result = this.runtime.commands.execute(commandId, params);
    this.applySelectionFromCommand(commandId, params, result);
    this.refresh();
    return result;
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
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
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
    this.runCommand('chart.add', chart);
  }
  updateChartBounds(id: string, bounds: ChartModel['bounds']): void {
    this.runCommand('chart.move', { id, sheetId: this.activeSheetId, bounds });
  }
  removeChart(id: string): void {
    this.runCommand('chart.remove', id);
  }
  addPivot(pivot: PivotModel): void {
    this.runCommand('pivot.add', pivot);
  }
  updatePivotLayout(pivotId: string, layout: PivotLayout): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, layout });
  }
  updatePivotConfiguration(
    pivotId: string,
    patch: Parameters<SpreadsheetApplication['updatePivotLayout']>[1] extends PivotLayout
      ? { sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotModel['slicers']; timelines?: PivotModel['timelines']; chartReferences?: PivotModel['chartReferences'] }
      : never,
  ): void {
    this.runCommand('pivot.update', { sheetId: this.activeSheetId, pivotId, ...patch });
  }
  refreshPivot(pivotId: string): void {
    const pivot = this.runtime.model.getSheet(this.activeSheetId).pivots.find((entry) => entry.id === pivotId);
    if (!pivot) return;
    this.runCommand('pivot.refresh', { sheetId: this.activeSheetId, pivotId });
    this.saveState = 'syncing';
    this.notify('Calculating pivot…');
    this.emit();
    void this.runtime.api
      .calculatePivot(this.runtime.model.unitId, pivotId)
      .then((response) => {
        this.runtime.pivotResults[pivotId] = response.result;
        this.notify('Pivot calculation complete');
        this.saveState = 'saved';
        this.refresh();
      })
      .catch((error: unknown) => {
        this.saveState = 'offline';
        this.notify(error instanceof Error ? error.message : 'Pivot calculation failed');
        this.emit();
      });
  }
  removePivot(id: string): void {
    this.runCommand('pivot.remove', id);
  }
  addShape(shape: ShapeModel): void {
    this.runCommand('shape.add', shape);
  }
  updateShapeBounds(id: string, bounds: ShapeModel['bounds']): void {
    this.runCommand('shape.move', { id, sheetId: this.activeSheetId, bounds });
  }
  removeShape(id: string): void {
    this.runCommand('shape.remove', id);
  }
  addSparkline(sparkline: SparklineModel): void {
    this.runCommand('sparkline.add', sparkline);
  }
  removeSparkline(id: string): void {
    this.runCommand('sparkline.remove', id);
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
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    const mentions = [...text.matchAll(/@([\w-]+)/g)].map((m) => m[1]).filter(Boolean) as string[];
    const comment: CellComment = { id: nextId('cmt'), author: this.actorId, text: text.trim(), createdAt: new Date().toISOString(), mentions, replies: [], resolved: false };
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: { ...(cell ?? { value: null }), comment } });
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
    this.emit();
  }
  removeFloatingObject(kind: 'chart' | 'shape', id: string): void {
    if (kind === 'chart') this.removeChart(id);
    else this.removeShape(id);
    this.selectedFloatingId = null;
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

  showPivotDetails(paths: PivotSourceRowPath[]): void {
    if (paths.length === 0) return;
    const first = paths[0]!;
    const source = this.runtime.model.getSheet(first.sheetId);
    const id = 'sheet-' + Math.random().toString(36).slice(2, 8);
    this.runCommand('sheet.add', { id, name: 'Pivot Details' });
    const values: CellData[][] = [];
    values.push(Array.from({ length: source.columnCount }, (_, column) => structuredClone(source.cells.get(0, column) ?? { value: null })));
    for (const path of paths) {
      const rowSheet = this.runtime.model.getSheet(path.sheetId);
      values.push(Array.from({ length: source.columnCount }, (_, column) => structuredClone(rowSheet.cells.get(path.row, column) ?? { value: null })));
    }
    if (values.length > 0) this.runCommand('sheet.range.set', { sheetId: id, startRow: 0, startColumn: 0, values });
    this.activeSheetId = id;
    this.refresh();
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
    this.printLayout = layout;
    this.showPrintPreview = true;
    this.notify('Preparing print preview');
    this.emit();
  }
  exportPdf(layout: PrintLayout): void {
    this.printLayout = layout;
    this.showPrintPreview = true;
    this.notify('Choose Save as PDF in the print dialog');
    this.emit();
  }

  async importXlsxBase64(base64: string): Promise<void> {
    const response = await fetch('/api/v1/files/import-xlsx', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ base64 }) });
    if (!response.ok) throw new Error('Import failed: ' + response.status);
    const result = (await response.json()) as import('@react-sheets/protocol').SnapshotResponse;
    hydrateRuntime(this.runtime, result);
    this.runtime.remoteConnected = true;
    if (typeof window !== 'undefined') {
      window.localStorage.setItem(UNIT_ID_STORAGE_KEY, result.snapshot.unitId);
      window.history.replaceState({}, '', `/workbooks/${encodeURIComponent(result.snapshot.unitId)}`);
    }
    this.activeSheetId = this.runtime.model.activeSheetId;
    this.selectionService.resetForSheet(this.activeSheetId);
    this.notify('XLSX imported');
    this.phase = 'ready';
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

  // Comment helpers omitted for brevity - add reply, resolve, remove, hyperlink
  replyComment(text: string): void {
    if (!text.trim()) return;
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    const comment = cell?.comment;
    if (!comment) return;
    const nextComment: CellComment = { ...structuredClone(comment), resolved: false, replies: [...(comment.replies ?? []), { id: nextId('reply'), author: this.actorId, text: text.trim(), createdAt: new Date().toISOString() }] };
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: { ...cell, comment: nextComment } });
  }
  resolveComment(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!cell?.comment) return;
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: { ...cell, comment: { ...structuredClone(cell.comment), resolved: true, resolvedAt: new Date().toISOString() } } });
  }
  removeComment(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!cell?.comment) return;
    const { comment: _c, ...rest } = cell;
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: rest as CellData });
  }
  setHyperlink(url: string): void {
    if (!url.trim()) return;
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: { ...(cell ?? { value: null }), hyperlink: url.trim() } });
  }
  removeHyperlink(): void {
    const sel = this.selectionService.getState();
    const cell = this.runtime.model.getSheet(this.activeSheetId).cells.get(sel.primaryRowIndex, sel.primaryColumnIndex);
    if (!cell?.hyperlink) return;
    const { hyperlink: _h, ...rest } = cell;
    this.runCommand('sheet.cell.set', { sheetId: this.activeSheetId, row: sel.primaryRowIndex, column: sel.primaryColumnIndex, value: rest as CellData });
  }
}

export { resolveUnitId, resolveActorId };
