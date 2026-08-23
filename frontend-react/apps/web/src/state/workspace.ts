import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  WorkbookModel,
  type CellComment,
  type CellData,
  type CellStyle,
  type ChartModel,
  type ConditionalFormatRule,
  type DataValidationRule,
  type FilterModel,
  type FreezeModel,
  type MergeSpan,
  type PivotFieldDefinition,
  type PivotLayout,
  type PivotModel,
  type PivotChartReference,
  type PivotSlicer,
  type PivotTimeline,
  type PivotResultTree,
  type PivotSourceRowPath,
  type RangeRef,
  type ShapeModel,
  type SparklineModel,
  type TableScalar,
  type TableFieldType,
  type WorkbookTableModel,
  type WorksheetModel,
} from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry } from '@react-sheets/command-runtime';
import {
  copyRangeToClipboardData,
  parseTsv,
  formatTsv,
  registerSheetCommands,
  computeConditionalOverlays,
  computeFilterHiddenRows,
  collectFindReplacements,
  validateDataInput,
  normalizeRangeRef,
  findValidationRule,
  validationList,
  type ConditionalOverlay,
} from '@react-sheets/sheet-features';
import { FormulaEngine, isFormulaError, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import {
  CollabSocketClient,
  WorkbookApiClient,
  type CollaborationMessage,
  type CollaborationMutation,
  type RevisionRecord,
  type SnapshotResponse,
  type TableRowsResponse,
} from '@react-sheets/protocol';
import { buildXlsxArchiveBase64, computePivotResult, getPivotFieldCatalog as buildPivotFieldCatalog, registerProSheetCommands, type PrintLayout } from '@react-sheets/pro-features';
import type { RibbonAction } from '../domain/ribbon-actions';

export type WorkspacePhase = 'empty' | 'error' | 'loading' | 'ready';
export type RibbonTabId = 'data' | 'home' | 'insert' | 'review' | 'view';
export type SidebarPanelId =
  | 'inspector'
  | 'chart'
  | 'pivot'
  | 'shape'
  | 'sparkline'
  | 'conditionalFormat'
  | 'dataValidation'
  | 'print'
  | 'history'
  | 'data';
export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing';

/** 多选区状态:ranges 均已归一化且属于当前活动工作表 */
export interface SelectionState {
  ranges: RangeRef[];
  primaryRowIndex: number;
  primaryColumnIndex: number;
  primaryRangeIndex: number;
}

export function createInitialSelection(): SelectionState {
  return {
    ranges: [],
    primaryRowIndex: 0,
    primaryColumnIndex: 0,
    primaryRangeIndex: 0,
  };
}

export interface SheetCell {
  address: string;
  displayValue?: string;
  formula?: string;
  style?: CellStyle;
  value: string;
  /** 单元格级标记(由视图构建阶段计算) */
  hasComment?: boolean;
  commentText?: string;
  comment?: CellComment;
  invalid?: boolean;
  hyperlink?: string;
  /** 条件格式覆盖 */
  overlay?: ConditionalOverlay;
}

export interface SheetRow {
  cells: SheetCell[];
  rowNumber: number;
  height: number;
}

export interface SheetView {
  columns: string[];
  id: string;
  isEmpty?: boolean;
  occupiedCellCount: number;
  numericAverage?: number;
  getCell: (row: number, column: number) => SheetCell | undefined;
  usedRange: RangeRef;
  name: string;
  rows: SheetRow[];
  charts: ChartModel[];
  pivots: PivotModel[];
  pivotResults: Record<string, PivotResultTree>;
  shapes: ShapeModel[];
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  merges: MergeSpan[];
  freeze: FreezeModel;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  /** 手工隐藏 ∪ 筛选隐藏后的行号列表 */
  hiddenRows: number[];
  /** 开启筛选的列(显示漏斗按钮) */
  filterColumns: number[];
  rowCount: number;
}

export interface PeerCursor {
  actorId: string;
  name: string;
  color: string;
  sheetId: string;
  row: number;
  column: number;
}

export interface WorkspaceState {
  unitId: string;
  workbookName: string;
  selection: SelectionState;
  /** 主单元格地址(派生,供公式栏/状态栏使用) */
  activeCell: string;
  activePanel: SidebarPanelId;
  activeSheetId: string;
  formulaDraft: string;
  /** 正在编辑的单元格;null 表示非编辑态 */
  editingCell: { row: number; column: number } | null;
  selectedFloatingId: string | null;
  notice: string;
  phase: WorkspacePhase;
  ribbonTab: RibbonTabId;
  saveState: SaveState;
  selectedSheet: SheetView;
  sheets: SheetView[];
  zoom: number;
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  tables: readonly WorkbookTableModel[];
  showFunctionWizard: boolean;
  showSortDialog: boolean;
  showFindReplace: boolean;
  findQuery: string;
  showPrintPreview: boolean;
  printLayout: PrintLayout;
  peers: PeerCursor[];
  collabStatus: 'connecting' | 'open' | 'closed';
  actorId: string;
}

export interface WorkspaceActions {
  // 选区
  selectCell: (address: string) => void;
  selectRange: (range: { startRow: number; startColumn: number; endRow: number; endColumn: number }, mode?: 'replace' | 'add') => void;
  movePrimary: (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => void;
  jumpEdge: (direction: 'up' | 'down' | 'left' | 'right', extend?: boolean) => void;
  selectAll: () => void;
  selectRowHeader: (startRow: number, endRow?: number, mode?: 'replace' | 'add') => void;
  selectColumnHeader: (startColumn: number, endColumn?: number, mode?: 'replace' | 'add') => void;
  // 编辑
  beginEdit: (initialText?: string) => void;
  cancelEdit: () => void;
  commitEdit: (moveAfter?: 'down' | 'up' | 'left' | 'right' | 'none') => void;
  setFormulaDraft: (value: string) => void;
  insertRefIntoDraft: (refText: string) => void;
  toggleAbsoluteReference: () => void;
  commitFormula: (overrideValue?: string) => void;
  moveCell: (address: string, direction: 'down' | 'left' | 'right' | 'up') => void;
  notify: (message: string) => void;
  redo: () => void;
  retry: () => void;
  selectSheet: (sheetId: string) => void;
  setActivePanel: (panel: SidebarPanelId) => void;
  setRibbonTab: (tab: RibbonTabId) => void;
  setZoom: (zoom: number) => void;
  undo: () => void;
  handleRibbonAction: (action: RibbonAction, payload?: unknown) => void;
  // Pro 模型
  addChart: (chart: ChartModel) => void;
  updateChartBounds: (id: string, bounds: ChartModel['bounds']) => void;
  removeChart: (id: string) => void;
  addPivot: (pivot: PivotModel) => void;
  updatePivotLayout: (pivotId: string, layout: PivotLayout) => void;
  updatePivotConfiguration: (pivotId: string, patch: { sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[] }) => void;
  refreshPivot: (id: string) => void;
  removePivot: (id: string) => void;
  addShape: (shape: ShapeModel) => void;
  updateShapeBounds: (id: string, bounds: ShapeModel['bounds']) => void;
  removeShape: (id: string) => void;
  addSparkline: (sparkline: SparklineModel) => void;
  removeSparkline: (id: string) => void;
  addConditionalFormat: (rule: ConditionalFormatRule) => void;
  removeConditionalFormat: (id: string) => void;
  addDataValidation: (rule: DataValidationRule) => void;
  removeDataValidation: (id: string) => void;
  // 数据功能
  addComment: (text: string) => void;
  replyComment: (text: string) => void;
  resolveComment: () => void;
  removeComment: () => void;
  setHyperlink: (url: string) => void;
  removeHyperlink: () => void;
  applyFilter: (column: number, patch: { selectedValues?: string[] | null; conditionOperator?: string; conditionValue?: string }) => void;
  clearFilter: () => void;
  findReplace: (params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }) => number;
  // 结构操作
  insertRowsAtPrimary: (count: number) => void;
  deleteRowsAtPrimary: () => void;
  insertColumnsAtPrimary: (count: number) => void;
  deleteColumnsAtPrimary: () => void;
  hideRowsAtPrimary: () => void;
  hideColumnsAtPrimary: () => void;
  unhideAll: () => void;
  toggleBandedRows: () => void;
  transposeSelection: () => void;
  flipSelection: (axis: 'h' | 'v') => void;
  splitByDelimiter: (delimiter: string) => void;
  // 名称与打印与导入
  defineName: (name: string, reference: string) => void;
  removeName: (name: string) => void;
  printWorkbook: (layout: PrintLayout) => void;
  exportPdf: (layout: PrintLayout) => void;
  setShowPrintPreview: (open: boolean) => void;
  importXlsxBase64: (base64: string) => Promise<void>;
  closeFunctionWizard: () => void;
  closeSortDialog: () => void;
  closeFindReplace: () => void;
  sortRange: (criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader: boolean) => void;
  // 面板数据访问
  getRangeMatrix: (range: RangeRef) => CellData[][];
  getRangeNumbers: (range: RangeRef) => number[];
  getValidationForPrimary: () => DataValidationRule | undefined;
  getValidationAt: (row: number, column: number) => string[] | undefined;
  addSheet: () => void;
  renameSheet: (sheetId: string, name: string) => void;
  cut: () => void;
  copy: () => void;
  paste: () => void;
  clearFormats: () => void;
  deleteSheet: (sheetId: string) => void;
  resizeRow: (row: number, heightPx: number) => void;
  resizeColumn: (column: number, widthPx: number) => void;
  fillRange: (targetRange: { startRow: number; endRow: number; startColumn: number; endColumn: number }) => void;
  setSelectedFloatingId: (id: string | null) => void;
  removeFloatingObject: (kind: 'chart' | 'shape', id: string) => void;
  getActiveSheetName: () => string;
  getPivotFieldCatalog: (range: RangeRef) => PivotFieldDefinition[];
  readDataTable: (tableId: string, offset?: number, limit?: number) => Promise<TableRowsResponse>;
  showPivotDetails: (paths: PivotSourceRowPath[]) => void;
}

export interface UseWorkspaceStateOptions {
  initialPhase?: WorkspacePhase;
}

interface RuntimeHandlers {
  onSaveState?: (state: SaveState) => void;
  onNotice?: (message: string) => void;
  onMutationsApplied?: () => void;
}

interface WorkspaceRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  model: WorkbookModel;
  commands: CommandRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
  pendingMutations: CollaborationMutation[];
  detachers: Array<() => void>;
  handlers: RuntimeHandlers;
  ownOperationIds: Set<string>;
  nextClientSequence: number;
  pivotResults: Record<string, PivotResultTree>;
  collab: CollabSocketClient | null;
}

const UNIT_ID_STORAGE_KEY = 'react-sheets:unitId';
const ACTOR_ID_STORAGE_KEY = 'react-sheets:actorId';

const PEER_COLORS = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];

function hashCode(input: string): number {
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
  }
  return hash;
}

function resolveUnitId(): string {
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

function resolveActorId(): string {
  if (typeof window === 'undefined') return 'actor-server';
  const existing = window.localStorage.getItem(ACTOR_ID_STORAGE_KEY);
  if (existing) return existing;
  const generated = typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID().slice(0, 8)
    : 'actor-' + Date.now().toString(36);
  window.localStorage.setItem(ACTOR_ID_STORAGE_KEY, generated);
  return generated;
}

const columns = Array.from({ length: 26 }, (_, index) => columnLabel(index));

function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function parseAddress(address: string): { column: number; row: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  return { column: column - 1, row: Number(match[2]) - 1 };
}

export function cellAddress(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

function toFormulaDisplay(value: FormulaValue): string {
  if (isFormulaError(value)) return value.code;
  if (Array.isArray(value)) {
    return value.length > 0 && Array.isArray(value[0]) ? String(value[0][0]) : String(value[0]);
  }
  return value == null ? '' : String(value);
}

function createWorkspaceRuntime(): WorkspaceRuntime {
  const model = new WorkbookModel(resolveUnitId(), 'Untitled workbook');
  const commands = new CommandRuntime(model);
  registerSheetCommands(commands);
  registerProSheetCommands(commands);
  const runtime: WorkspaceRuntime = {
    api: new WorkbookApiClient(),
    formula: new FormulaEngine({ defaultSheetId: 'sheet-1' }),
    model,
    commands,
    remoteConnected: false,
    remoteRevision: 0,
    pendingMutations: [],
    detachers: [],
    handlers: {},
    ownOperationIds: new Set(),
    nextClientSequence: 0,
    pivotResults: {},
    collab: null,
  };
  attachCoreListeners(runtime);
  return runtime;
}

// ---------- 引擎同步:任何来源的 cell 级变更都镜像到 FormulaEngine ----------

function syncEngineCell(
  engine: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
  data: CellData | undefined,
): void {
  const address = { sheetId, row, column };
  const hasContent = data !== undefined && (data.formula !== undefined || data.value != null);
  if (!hasContent) {
    engine.clearCell(address);
    return;
  }
  if (data!.formula) {
    engine.setFormula(address, data!.formula);
  } else {
    engine.setValue(address, data!.value as never);
  }
}

function attachCoreListeners(runtime: WorkspaceRuntime): void {
  detachCoreListeners(runtime);

  // 1) 公式引擎同步(command/undo/redo/remote 全部来源)
  runtime.detachers.push(
    runtime.commands.onMutation((mutation) => {
      const changedSheet = runtime.model.getSheets().find((sheet) => sheet.id === mutation.sheetId);
      for (const pivot of changedSheet?.pivots ?? []) delete runtime.pivotResults[pivot.id];
      switch (mutation.id) {
        case 'cell.set': {
          const params = mutation.params as { row: number; column: number; value: CellData };
          syncEngineCell(runtime.formula, mutation.sheetId, params.row, params.column, params.value);
          break;
        }
        case 'cell.restore': {
          const params = mutation.params as { row: number; column: number; previous?: CellData };
          syncEngineCell(runtime.formula, mutation.sheetId, params.row, params.column, params.previous);
          break;
        }
        case 'range.set': {
          const params = mutation.params as { startRow: number; startColumn: number; values: CellData[][] };
          params.values.forEach((rowValues, rowOffset) =>
            rowValues.forEach((value, columnOffset) => {
              syncEngineCell(
                runtime.formula,
                mutation.sheetId,
                params.startRow + rowOffset,
                params.startColumn + columnOffset,
                value,
              );
            }),
          );
          break;
        }
        case 'range.clear': {
          const params = mutation.params as {
            range: RangeRef;
            mode?: 'all' | 'contents' | 'formats';
          };
          if (params.mode === 'formats') break;
          for (let r = params.range.startRow; r <= params.range.endRow; r++) {
            for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
              runtime.formula.clearCell({ sheetId: mutation.sheetId, row: r, column: c });
            }
          }
          break;
        }
        case 'rows.inserted':
        case 'rows.deleted': {
          const params = mutation.params as { at: number; count: number };
          runtime.formula.remapStructure(mutation.sheetId, {
            axis: 'row',
            at: params.at,
            count: params.count,
            op: mutation.id === 'rows.inserted' ? 'insert' : 'delete',
          });
          break;
        }
        case 'columns.inserted':
        case 'columns.deleted': {
          const params = mutation.params as { at: number; count: number };
          runtime.formula.remapStructure(mutation.sheetId, {
            axis: 'column',
            at: params.at,
            count: params.count,
            op: mutation.id === 'columns.inserted' ? 'insert' : 'delete',
          });
          break;
        }
        case 'name.set':
        case 'name.remove':
          runtime.formula.setDefinedNames(runtime.model.definedNames);
          break;
        default:
          break;
      }
    }),
  );

  // 2) 正向命令的变更缓冲(undo/redo/remote 不上报协同)
  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (source !== 'command') return;
      runtime.pendingMutations.push({
        id: mutation.id,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: mutation.affectedRanges,
      });
    }),
  );

  // 3) 根命令结束:冲刷为 changeset;并触发筛选重应用
  runtime.detachers.push(
    runtime.commands.onCommand((_commandId, _params, result) => {
      if (runtime.commands.activeDepth > 0) return; // 嵌套子命令不冲刷
      const batch = runtime.pendingMutations;
      runtime.pendingMutations = [];
      runtime.handlers.onMutationsApplied?.();
      if (batch.length === 0) return;
      submitChangeset(runtime, result.operationId, batch);
    }),
  );
}

function detachCoreListeners(runtime: WorkspaceRuntime): void {
  for (const detach of runtime.detachers) detach();
  runtime.detachers = [];
  runtime.pendingMutations = [];
}

function submitChangeset(
  runtime: WorkspaceRuntime,
  operationId: string,
  mutations: CollaborationMutation[],
): void {
  if (!runtime.collab) {
    runtime.handlers.onSaveState?.('offline');
    return;
  }
  runtime.ownOperationIds.add(operationId);
  const changeSet = {
    schema: 'CollaborationChangeSetV1' as const,
    operationId,
    unitId: runtime.model.unitId,
    actorId: resolveActorId(),
    clientSequence: ++runtime.nextClientSequence,
    baseRevision: runtime.remoteRevision,
    mutations,
    createdAt: new Date().toISOString(),
  };
  runtime.handlers.onSaveState?.('saving');
  if (!runtime.collab.send({ type: 'changeset.submit', payload: changeSet })) {
    runtime.handlers.onSaveState?.('syncing');
  }
}

// ---------- 启动恢复 ----------

function rebuildFormulaInto(engine: FormulaEngine, workbook: WorkbookModel): void {
  engine.reset();
  engine.setDefinedNames(workbook.definedNames);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
}

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.activeSheetId });
  engine.setDefinedNames(workbook.definedNames);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
  return engine;
}

function hydrateRuntime(runtime: WorkspaceRuntime, response: SnapshotResponse): void {
  const workbook = WorkbookModel.fromSnapshot(response.snapshot);
  detachCoreListeners(runtime);
  runtime.model = workbook;
  runtime.commands = new CommandRuntime(workbook);
  registerSheetCommands(runtime.commands);
  registerProSheetCommands(runtime.commands);
  runtime.formula = rebuildFormulaEngine(workbook);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  runtime.pivotResults = {};
}

// ---------- 视图构建 ----------

function formatDisplayValue(
  cell: CellData | undefined,
  formula: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
): string {
  if (!cell) return '';
  if (cell.formula) {
    const computed = formula.getCellValue({ sheetId, row, column });
    return toFormulaDisplay(computed);
  }
  if (cell.value == null) return '';
  if (typeof cell.value === 'number') {
    return formatNumberValue(cell.value, cell.numberFormat ?? cell.style?.numberFormat);
  }
  return String(cell.value);
}

function usedRangeOfSheet(sheet: WorksheetModel): RangeRef {
  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxColumn = 0;
  let numericSum = 0;
  let numericCount = 0;
  sheet.cells.forEach((_cell, row, column) => {
    minRow = Math.min(minRow, row);
    minColumn = Math.min(minColumn, column);
    maxRow = Math.max(maxRow, row);
    maxColumn = Math.max(maxColumn, column);
    if (typeof _cell.value === 'number' && Number.isFinite(_cell.value)) {
      numericSum += _cell.value;
      numericCount += 1;
    }
  });
  return {
    sheetId: sheet.id,
    startRow: Number.isFinite(minRow) ? minRow : 0,
    endRow: Number.isFinite(minRow) ? maxRow : 0,
    startColumn: Number.isFinite(minColumn) ? minColumn : 0,
    endColumn: Number.isFinite(minColumn) ? maxColumn : 0,
  };
}

function inferTableFieldType(values: CellData['value'][]): TableFieldType {
  const present = values.filter((value) => value != null && value !== '');
  if (present.length === 0) return 'mixed';
  if (present.every((value) => typeof value === 'number')) return 'number';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value))) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

function toSheetView(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  formula: FormulaEngine,
  showInvalid: boolean,
  cachedPivotResults: Readonly<Record<string, PivotResultTree>> = {},
): SheetView {
  const overlays = computeConditionalOverlays(sheet);
  const filterHidden = computeFilterHiddenRows(sheet);
  const hiddenRows = new Set<number>([...sheet.hiddenRows, ...filterHidden]);
  const filterColumns = sheet.filter ? Object.keys(sheet.filter.criteria).map(Number) : [];
  const viewColumns = Array.from({ length: Math.max(26, sheet.columnCount) }, (_, index) => columnLabel(index));
  const usedRange = usedRangeOfSheet(sheet);

  const getCell = (row: number, column: number): SheetCell | undefined => {
    if (row < 0 || row >= sheet.rowCount || column < 0 || column >= sheet.columnCount) return undefined;
    const modelCell = sheet.cells.get(row, column);
    const value = formatDisplayValue(modelCell, formula, sheet.id, row, column);
    const key = `${row}:${column}`;
    const overlay = overlays.get(key);
    const style = overlay?.style
      ? { ...(modelCell?.style ?? {}), ...overlay.style }
      : modelCell?.style;
    const validation = validateDataInput(sheet, row, column, modelCell?.value ?? null);
    return {
      address: cellAddress(row, column),
      formula: modelCell?.formula,
      style,
      value,
      displayValue: value,
      hasComment: Boolean(modelCell?.comment),
      commentText: modelCell?.comment?.text,
      comment: modelCell?.comment ? structuredClone(modelCell.comment) : undefined,
      invalid: showInvalid && modelCell?.value != null && !validation.valid,
      hyperlink: modelCell?.hyperlink,
      overlay,
    };
  };

  const rows: SheetRow[] = [];

  // Rows are a bounded preview for panels and print UI. Canvas reads other cells through getCell.
  const previewRows = Math.min(Math.max(60, sheet.rowCount), 200);
  for (let row = 0; row < previewRows; row += 1) {
    if (hiddenRows.has(row)) continue;
    const cells: SheetCell[] = [];
    for (let column = 0; column < viewColumns.length; column += 1) {
      const cell = getCell(row, column);
      if (cell) cells.push(cell);
    }
    rows.push({ rowNumber: row + 1, cells, height: sheet.rowHeights[row] ?? 28 });
  }

  const isEmpty = sheet.cells.count() === 0;
  const pivotResults: Record<string, PivotResultTree> = {};
  for (const pivot of sheet.pivots) {
    try {
      pivotResults[pivot.id] = cachedPivotResults[pivot.id] ?? computePivotResult(workbook, pivot);
    } catch {
      // An invalid pivot remains visible in the model so the UI can report it;
      // a failed projection must not prevent the worksheet from rendering.
    }
  }

  return {
    id: sheet.id,
    name: sheet.name,
    columns: viewColumns,
    rows,
    getCell,
    usedRange,
    isEmpty,
    occupiedCellCount: sheet.cells.count(),
    numericAverage: numericCount > 0 ? numericSum / numericCount : undefined,
    charts: [...sheet.charts],
    pivots: [...sheet.pivots],
    pivotResults,
    shapes: [...sheet.shapes],
    sparklines: [...sheet.sparklines],
    conditionalFormats: [...sheet.conditionalFormats],
    dataValidations: [...sheet.dataValidations],
    merges: [...sheet.merges],
    freeze: { ...sheet.freeze },
    rowHeights: { ...sheet.rowHeights },
    columnWidths: { ...sheet.columnWidths },
    hiddenRows: [...hiddenRows].sort((a, b) => a - b),
    filterColumns,
    rowCount: sheet.rowCount,
  };
}

export function getInitialWorkspacePhase(): WorkspacePhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'loading' || queryPhase === 'error' || queryPhase === 'empty'
    ? queryPhase
    : 'ready';
}

// ---------- Hook ----------

export function useWorkspaceState({ initialPhase = 'ready' }: UseWorkspaceStateOptions = {}): {
  actions: WorkspaceActions;
  state: WorkspaceState;
} {
  const runtimeRef = useRef<WorkspaceRuntime | null>(null);
  if (!runtimeRef.current) runtimeRef.current = createWorkspaceRuntime();
  const runtime = runtimeRef.current;

  const [phase, setPhase] = useState<WorkspacePhase>(initialPhase);
  const [selection, setSelection] = useState<SelectionState>(() => ({
    ...createInitialSelection(),
    ranges: [normalizeRangeRef({ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })],
  }));
  const [formulaDraft, setFormulaDraft] = useState('');
  const [editingCell, setEditingCell] = useState<{ row: number; column: number } | null>(null);
  const [ribbonTab, setRibbonTab] = useState<RibbonTabId>('home');
  const [activePanel, setActivePanel] = useState<SidebarPanelId>('inspector');
  const [zoom, setZoomState] = useState(100);
  const [saveState, setSaveState] = useState<SaveState>('saved');
  const [notice, setNotice] = useState('Workbook engine ready');
  const [modelVersion, setModelVersion] = useState(0);
  const [remoteRevisions, setRemoteRevisions] = useState<RevisionRecord[]>([]);

  const [activeSheetId, setActiveSheetId] = useState(runtime.model.activeSheetId);

  const [showFunctionWizard, setShowFunctionWizard] = useState(false);
  const [showSortDialog, setShowSortDialog] = useState(false);
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [showPrintPreview, setShowPrintPreviewState] = useState(false);
  const [printLayout, setPrintLayout] = useState<PrintLayout>({
    paper: 'A4',
    orientation: 'portrait',
    margin: { top: 20, right: 20, bottom: 20, left: 20 },
  });
  const [selectedFloatingId, setSelectedFloatingIdState] = useState<string | null>(null);
  const [peers, setPeers] = useState<PeerCursor[]>([]);
  const [collabStatus, setCollabStatus] = useState<'connecting' | 'open' | 'closed'>('closed');

  const actorId = useMemo(() => resolveActorId(), []);
  const activeSheetIdRef = useRef(activeSheetId);
  const selectionRef = useRef(selection);
  activeSheetIdRef.current = activeSheetId;
  selectionRef.current = selection;

  const refresh = useCallback(() => setModelVersion((version) => version + 1), []);

  // 将回调注入 runtime(每次渲染更新,保持最新闭包)
  useEffect(() => {
    runtime.handlers.onSaveState = setSaveState;
    runtime.handlers.onNotice = setNotice;
    runtime.handlers.onMutationsApplied = () => refresh();
  }, [runtime, refresh]);

  // P7 协同:?collab=1 时建立 WebSocket,接收远端变更(幂等)与光标
  useEffect(() => {
    if (typeof window === 'undefined') return;
    const protocol = window.location.protocol === 'https:' ? 'wss' : 'ws';
    const client = new CollabSocketClient(protocol + '://' + window.location.host + '/api/v1/collab');
    runtime.collab = client;

    const applyRemote = (message: CollaborationMessage) => {
      if (message.type === 'revision.created') {
        if (message.payload.unitId !== runtime.model.unitId) return;
        // 幂等:自己提交的 changeset 经服务器回播时跳过
        if (runtime.ownOperationIds.has(message.payload.operationId)) return;
        runtime.commands.applyRemoteMutations(
          message.payload.mutations.map((mutation) => ({ ...mutation, unitId: runtime.model.unitId })),
        );
        runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
        runtime.handlers.onMutationsApplied?.();
        void runtime.api.listRevisions(runtime.model.unitId).then(setRemoteRevisions).catch(() => undefined);
      } else if (message.type === 'changeset.ack') {
        runtime.ownOperationIds.add(message.operationId);
        runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
        runtime.handlers.onSaveState?.('saved');
        void runtime.api.listRevisions(runtime.model.unitId).then(setRemoteRevisions).catch(() => undefined);
      } else if (message.type === 'changeset.reject') {
        runtime.ownOperationIds.delete(message.operationId);
        runtime.handlers.onSaveState?.('offline');
        runtime.handlers.onNotice?.(`Change rejected: ${message.error.message}`);
        void runtime.api.getSnapshot(runtime.model.unitId).then((snapshot) => {
          hydrateRuntime(runtime, snapshot);
          runtime.handlers.onMutationsApplied?.();
        }).catch(() => undefined);
      } else if (message.type === 'cursor.updated' || message.type === 'presence.updated') {
        if (message.unitId && message.unitId !== runtime.model.unitId) return;
        const cursorState = message.state as { row?: number; column?: number; name?: string; sheetId?: string } | null;
        const color = PEER_COLORS[Math.abs(hashCode(message.actorId)) % PEER_COLORS.length]!;
        setPeers((current) => {
          const others = current.filter((peer) => peer.actorId !== message.actorId);
          return [
            ...others,
            {
              actorId: message.actorId,
              name: cursorState?.name ?? message.actorId.slice(0, 6),
              color,
              sheetId: cursorState?.sheetId ?? runtime.model.activeSheetId,
              row: cursorState?.row ?? 0,
              column: cursorState?.column ?? 0,
            },
          ];
        });
      }
    };
    const detachMessage = client.onMessage(applyRemote);
    const detachStatus = client.onStatus((status: 'connecting' | 'open' | 'closed') => {
      setCollabStatus(status);
      runtime.remoteConnected = status !== 'closed';
    });
    client.open();

    // 本地光标广播
    let lastBroadcast = '';
    const broadcastTimer = window.setInterval(() => {
      const currentSelection = selectionRef.current;
      const currentSheetId = activeSheetIdRef.current;
      const key = currentSheetId + ':' + currentSelection.primaryRowIndex + ':' + currentSelection.primaryColumnIndex;
      if (key === lastBroadcast) return;
      lastBroadcast = key;
      client.send({
        type: 'cursor.updated',
        unitId: runtime.model.unitId,
        actorId,
        state: { row: currentSelection.primaryRowIndex, column: currentSelection.primaryColumnIndex, sheetId: currentSheetId },
      });
    }, 400);

    return () => {
      window.clearInterval(broadcastTimer);
      detachMessage();
      detachStatus();
      client.close();
      runtime.collab = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [runtime]);

  // 启动恢复 / 创建 / 离线回退
  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
        hydrateRuntime(runtime, snapshotResponse);
        void runtime.api.listRevisions(runtime.model.unitId).then(setRemoteRevisions).catch(() => setRemoteRevisions([]));
        runtime.remoteConnected = true;
        if (!active) return;
        runtime.handlers.onSaveState?.('saved');
        runtime.handlers.onNotice?.('Workbook restored from server');
        setPhase('ready');
        setActiveSheetId(runtime.model.activeSheetId);
        refresh();
      } catch {
        try {
          const created = await runtime.api.createWorkbook(runtime.model.snapshot());
          void runtime.api.listRevisions(runtime.model.unitId).then(setRemoteRevisions).catch(() => setRemoteRevisions([]));
          runtime.remoteConnected = true;
          runtime.remoteRevision = Math.max(runtime.remoteRevision, created.revision);
          if (!active) return;
          runtime.handlers.onSaveState?.('saved');
          runtime.handlers.onNotice?.('SQLite sync connected');
          setPhase('ready');
          refresh();
        } catch {
          if (!active) return;
           runtime.handlers.onSaveState?.('offline');
           runtime.handlers.onNotice?.('Running local in-memory engine');
           setPhase('ready');
        }
      }
    })();
    return () => {
      active = false;
    };
  }, [runtime, refresh]);

  const sheets = useMemo(
    () => runtime.model.getSheets().map((sheet) => toSheetView(runtime.model, sheet, runtime.formula, true, runtime.pivotResults)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [modelVersion, runtime],
  );
  const tables = useMemo(() => [...runtime.model.tables.values()].map((table) => structuredClone(table)), [modelVersion, runtime]);

  const selectedSheet = useMemo(
    () => sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0]!,
    [activeSheetId, sheets],
  );

  const activeCell = useMemo(
    () => cellAddress(selection.primaryRowIndex, selection.primaryColumnIndex),
    [selection],
  );

  // ---- 选区 ----

  const clampRow = useCallback(
    (row: number) => Math.max(0, Math.min(selectedSheet.rowCount - 1, row)),
    [selectedSheet.rowCount],
  );
  const clampColumn = useCallback(
    (column: number) => Math.max(0, Math.min(columns.length - 1, column)),
    [],
  );

  const syncDraftFromCell = useCallback(
    (row: number, column: number) => {
      const cell = runtime.model.getSheet(activeSheetId).cells.get(row, column);
      setFormulaDraft(cell?.formula ?? (cell?.value == null ? '' : String(cell.value)));
    },
    [activeSheetId, runtime],
  );

  const setPrimaryAndSync = useCallback(
    (row: number, column: number) => {
      setSelection((previous) => ({
        ranges: [normalizeRangeRef({
          sheetId: activeSheetId,
          startRow: row,
          endRow: row,
          startColumn: column,
          endColumn: column,
        })],
        primaryRowIndex: row,
        primaryColumnIndex: column,
        primaryRangeIndex: 0,
      }));
      syncDraftFromCell(row, column);
    },
    [activeSheetId, syncDraftFromCell],
  );

  const selectCell = useCallback(
    (address: string) => {
      // 编辑态下点击其他单元格 = 向公式草稿插入引用
      if (editingCell) {
        const parsed = parseAddress(address);
        if (parsed) {
          const reference = `${columnLabel(parsed.column)}${parsed.row + 1}`;
          setFormulaDraft((draft) => draft + reference);
        }
        return;
      }
      const parsed = parseAddress(address);
      if (!parsed) return;
      setPrimaryAndSync(clampRow(parsed.row), clampColumn(parsed.column));
    },
    [clampColumn, clampRow, editingCell, setPrimaryAndSync],
  );

  const selectRange = useCallback(
    (range: { startRow: number; startColumn: number; endRow: number; endColumn: number }, mode: 'replace' | 'add' = 'replace') => {
      const normalized = normalizeRangeRef({
        sheetId: activeSheetId,
        startRow: clampRow(Math.min(range.startRow, range.endRow)),
        endRow: clampRow(Math.max(range.startRow, range.endRow)),
        startColumn: clampColumn(Math.min(range.startColumn, range.endColumn)),
        endColumn: clampColumn(Math.max(range.startColumn, range.endColumn)),
      });
      setSelection((previous) => {
        if (mode === 'add' && previous.ranges.length > 0) {
          return { ...previous, ranges: [...previous.ranges, normalized], primaryRangeIndex: previous.ranges.length };
        }
        return {
          ranges: [normalized],
          primaryRowIndex: normalized.startRow,
          primaryColumnIndex: normalized.startColumn,
          primaryRangeIndex: 0,
        };
      });
      syncDraftFromCell(normalized.startRow, normalized.startColumn);
    },
    [activeSheetId, clampColumn, clampRow, syncDraftFromCell],
  );

  const movePrimary = useCallback(
    (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => {
      setSelection((previous) => {
        const targetRow = clampRow(previous.primaryRowIndex + rowDelta);
        const targetColumn = clampColumn(previous.primaryColumnIndex + columnDelta);
        if (opts?.extend && previous.ranges.length > 0) {
          const range = previous.ranges[previous.primaryRangeIndex] ?? previous.ranges[0]!;
          const anchorRow = previous.primaryRowIndex <= (range.startRow + range.endRow) / 2 ? range.endRow : range.startRow;
          const anchorColumn = previous.primaryColumnIndex <= (range.startColumn + range.endColumn) / 2 ? range.endColumn : range.startColumn;
          const next = normalizeRangeRef({
            sheetId: activeSheetId,
            startRow: Math.min(anchorRow, targetRow),
            endRow: Math.max(anchorRow, targetRow),
            startColumn: Math.min(anchorColumn, targetColumn),
            endColumn: Math.max(anchorColumn, targetColumn),
          });
          const ranges = [...previous.ranges];
          ranges[previous.primaryRangeIndex] = next;
          return { ...previous, ranges, primaryRowIndex: targetRow, primaryColumnIndex: targetColumn };
        }
        return {
          ranges: [normalizeRangeRef({
            sheetId: activeSheetId,
            startRow: targetRow,
            endRow: targetRow,
            startColumn: targetColumn,
            endColumn: targetColumn,
          })],
          primaryRowIndex: targetRow,
          primaryColumnIndex: targetColumn,
          primaryRangeIndex: 0,
        };
      });
    },
    [activeSheetId, clampColumn, clampRow],
  );

  const jumpEdge = useCallback(
    (direction: 'up' | 'down' | 'left' | 'right', extend = false) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      let row = selection.primaryRowIndex;
      let column = selection.primaryColumnIndex;
      const step = direction === 'up' ? -1 : direction === 'down' ? 1 : direction === 'left' ? -1 : 1;
      const horizontal = direction === 'left' || direction === 'right';
      let cursor = horizontal ? column : row;
      cursor += step;
      while (
        cursor >= 0 &&
        (horizontal ? cursor < columns.length : cursor < sheet.rowCount)
      ) {
        const cellValue = horizontal
          ? sheet.cells.get(row, cursor)?.value
          : sheet.cells.get(cursor, column)?.value;
        if (cellValue != null && cellValue !== '') break;
        cursor += step;
      }
      // 越界则停在最后一个非空或边界
      if (horizontal) column = Math.max(0, Math.min(columns.length - 1, cursor));
      else row = Math.max(0, Math.min(sheet.rowCount - 1, cursor));

      if (extend) {
        movePrimary(row - selection.primaryRowIndex, column - selection.primaryColumnIndex, { extend: true });
      } else {
        setPrimaryAndSync(row, column);
      }
    },
    [activeSheetId, movePrimary, selection.primaryColumnIndex, selection.primaryRowIndex, runtime.model, setPrimaryAndSync],
  );

  const selectAll = useCallback(() => {
    selectRange({ startRow: 0, startColumn: 0, endRow: selectedSheet.rowCount - 1, endColumn: columns.length - 1 }, 'replace');
  }, [selectRange, selectedSheet.rowCount]);

  const selectRowHeader = useCallback(
    (startRow: number, endRow = startRow, mode: 'replace' | 'add' = 'replace') => {
      selectRange({ startRow, startColumn: 0, endRow, endColumn: columns.length - 1 }, mode);
    },
    [selectRange],
  );

  const selectColumnHeader = useCallback(
    (startColumn: number, endColumn = startColumn, mode: 'replace' | 'add' = 'replace') => {
      selectRange({ startRow: 0, startColumn, endRow: selectedSheet.rowCount - 1, endColumn }, mode);
    },
    [selectRange, selectedSheet.rowCount],
  );

  // ---- 编辑 ----

  const beginEdit = useCallback(
    (initialText?: string) => {
      const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
      setFormulaDraft(initialText ?? cell?.formula ?? (cell?.value == null ? '' : String(cell.value)));
      setEditingCell({ row: selection.primaryRowIndex, column: selection.primaryColumnIndex });
    },
    [activeSheetId, runtime, selection.primaryColumnIndex, selection.primaryRowIndex],
  );

  const cancelEdit = useCallback(() => {
    setEditingCell(null);
    syncDraftFromCell(selection.primaryRowIndex, selection.primaryColumnIndex);
  }, [selection.primaryColumnIndex, selection.primaryRowIndex, syncDraftFromCell]);

  const commitFormula = useCallback(
    (overrideValue?: string) => {
      if (phase !== 'ready') return false;
      const row = overrideTargetRef.current?.row ?? selection.primaryRowIndex;
      const column = overrideTargetRef.current?.column ?? selection.primaryColumnIndex;
      const raw = (overrideValue !== undefined ? overrideValue : formulaDraft).trim();

      const sheet = runtime.model.getSheet(activeSheetId);
      const existingStyle = sheet.cells.get(row, column)?.style;
      const isFormula = raw.startsWith('=');
      const value = isFormula
        ? null
        : raw === ''
          ? null
          : Number.isFinite(Number(raw))
            ? Number(raw)
            : raw;

      // 数据验证拦截
      const candidate: CellData = { value, formula: isFormula ? raw : undefined, style: existingStyle };
      const validation = validateDataInput(sheet, row, column, value);
      if (!validation.valid) {
        if (validation.blocking) {
          runtime.handlers.onNotice?.(validation.message ?? '输入不符合数据验证规则');
          return false;
        }
        runtime.handlers.onNotice?.(`警告: ${validation.message ?? '数据验证未通过'}`);
      }

      const cellData: CellData = isFormula ? { value: null, formula: raw, style: existingStyle } : candidate;
      runtime.commands.execute('sheet.cell.set', {
        sheetId: activeSheetId,
        row,
        column,
        value: cellData,
      });

      refresh();
      return true;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [activeSheetId, formulaDraft, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex],
  );

  const overrideTargetRef = useRef<{ row: number; column: number } | null>(null);

  const commitEdit = useCallback(
    (moveAfter: 'down' | 'up' | 'left' | 'right' | 'none' = 'down') => {
      if (!editingCell) return;
      overrideTargetRef.current = { row: editingCell.row, column: editingCell.column };
      const committed = commitFormula(formulaDraft.trim());
      overrideTargetRef.current = null;
      setEditingCell(null);
      if (!committed) return; // 验证拦截:保持编辑值留在草稿
      const deltas = { down: [1, 0], up: [-1, 0], left: [0, -1], right: [0, 1], none: [0, 0] } as const;
      const [dr, dc] = deltas[moveAfter];
      movePrimary(dr, dc);
      const targetRow = clampRow(editingCell.row + dr);
      const targetColumn = clampColumn(editingCell.column + dc);
      syncDraftFromCell(targetRow, targetColumn);
    },
    [clampColumn, clampRow, commitFormula, editingCell, formulaDraft, movePrimary, syncDraftFromCell],
  );

  const insertRefIntoDraft = useCallback((refText: string) => {
    setFormulaDraft((draft) => {
      const needsSeparator = /[A-Za-z0-9)]$/.test(draft) && /^[A-Za-z0-9]/.test(refText);
      return draft + (needsSeparator ? '' : '') + refText;
    });
  }, []);

  const toggleAbsoluteReference = useCallback(() => {
    setFormulaDraft((draft) =>
      draft.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (_match, dCol, col, dRow, row) => {
        const nextDCol = dCol ? '' : '$';
        const nextDRow = dRow ? '' : '$';
        return `${nextDCol}${col}${nextDRow}${row}`;
      }),
    );
  }, []);

  const moveCell = useCallback(
    (address: string, direction: 'down' | 'left' | 'right' | 'up') => {
      const parsed = parseAddress(address);
      if (!parsed) return;
      const offsets = { left: [0, -1], right: [0, 1], up: [-1, 0], down: [1, 0] } as const;
      const [rowOffset, columnOffset] = offsets[direction];
      const targetRow = clampRow(parsed.row + rowOffset);
      const targetColumn = clampColumn(parsed.column + columnOffset);
      setPrimaryAndSync(targetRow, targetColumn);
    },
    [clampColumn, clampRow, setPrimaryAndSync],
  );

  // ---- 工作表管理 ----

  const selectSheet = useCallback(
    (sheetId: string) => {
      runtime.model.activeSheetId = sheetId;
      setActiveSheetId(sheetId);
      setSelection(createInitialSelection());
      setEditingCell(null);
      setFormulaDraft('');
      refresh();
    },
    [refresh, runtime],
  );

  const undo = useCallback(() => {
    const applied = runtime.commands.undo();
    if (applied) {
      refresh();
      setNotice('Undo applied');
    }
  }, [refresh, runtime]);

  const redo = useCallback(() => {
    const applied = runtime.commands.redo();
    if (applied) {
      refresh();
      setNotice('Redo applied');
    }
  }, [refresh, runtime]);

  const retry = useCallback(() => {
    setPhase('ready');
    setNotice('Workspace ready');
  }, []);

  // ---- Ribbon 动作 ----

  const handleRibbonAction = (action: RibbonAction, payload?: unknown) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      const primaryRange = selection.ranges[selection.primaryRangeIndex]
        ?? normalizeRangeRef({
          sheetId: activeSheetId,
          startRow: selection.primaryRowIndex,
          endRow: selection.primaryRowIndex,
          startColumn: selection.primaryColumnIndex,
          endColumn: selection.primaryColumnIndex,
        });
      const applyStyleToPrimary = (style: Partial<CellStyle>) => {
        runtime.commands.execute('sheet.style.set', { sheetId: activeSheetId, range: primaryRange, style });
        refresh();
      };
      const readStyle = (): CellStyle | undefined =>
        sheet.cells.get(selection.primaryRowIndex, selection.primaryColumnIndex)?.style;

      switch (action) {
        case 'undo': undo(); break;
        case 'redo': redo(); break;
        case 'copy': {
          const data = copyRangeToClipboardData(runtime.model, primaryRange);
          void navigator.clipboard?.writeText(formatTsv(data.values));
          setNotice('Copied to clipboard');
          break;
        }
        case 'cut': {
          const data = copyRangeToClipboardData(runtime.model, primaryRange);
          void navigator.clipboard?.writeText(formatTsv(data.values));
          runtime.commands.execute('sheet.range.clear', { sheetId: activeSheetId, range: primaryRange });
          setFormulaDraft('');
          refresh();
          setNotice('Cut to clipboard');
          break;
        }
        case 'paste': {
          void navigator.clipboard
            ?.readText()
            .then((text) => {
              const parsed = parseTsv(text);
              if (parsed.length === 0) return;
              runtime.commands.execute('sheet.range.set', {
                sheetId: activeSheetId,
                startRow: selection.primaryRowIndex,
                startColumn: selection.primaryColumnIndex,
                values: parsed.map((row) => row.map((cell) => ({ ...cell, style: undefined, comment: undefined }))),
              });
              refresh();
              setNotice('Pasted from clipboard');
            })
            .catch(() => setNotice('Clipboard unavailable'));
          break;
        }
        case 'bold':
          applyStyleToPrimary({ bold: !readStyle()?.bold });
          break;
        case 'italic':
          applyStyleToPrimary({ italic: !readStyle()?.italic });
          break;
        case 'underline':
          applyStyleToPrimary({ underline: !readStyle()?.underline });
          break;
        case 'strikethrough':
          applyStyleToPrimary({ strikethrough: !readStyle()?.strikethrough });
          break;
        case 'align-left':
        case 'align-center':
        case 'align-right':
          applyStyleToPrimary({ horizontalAlignment: action.replace('align-', '') as 'left' | 'center' | 'right' });
          break;
        case 'v-align-top':
        case 'v-align-middle':
        case 'v-align-bottom':
          applyStyleToPrimary({ verticalAlignment: action.replace('v-align-', '') as 'top' | 'middle' | 'bottom' });
          break;
        case 'wrap-text':
          applyStyleToPrimary({ wrapText: !readStyle()?.wrapText });
          break;
        case 'rotate': {
          const degrees = typeof payload === 'number' ? payload : Number(payload);
          if (Number.isFinite(degrees)) applyStyleToPrimary({ textRotate: degrees });
          break;
        }
        case 'font-family':
          if (typeof payload === 'string') applyStyleToPrimary({ fontFamily: payload });
          break;
        case 'font-size': {
          const size = Number(payload);
          if (Number.isFinite(size) && size > 0) applyStyleToPrimary({ fontSize: size });
          break;
        }
        case 'merge-cells': {
          const range = primaryRange;
          if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
            // 无选区时退化为两列合并(保留旧行为语义)
            runtime.commands.execute('sheet.merge.set', {
              sheetId: activeSheetId,
              range: { ...range, endColumn: range.endColumn + 1 },
            });
          } else {
            runtime.commands.execute('sheet.merge.set', { sheetId: activeSheetId, range });
          }
          refresh();
          break;
        }
        case 'unmerge-cells': {
          for (const range of selection.ranges) {
            runtime.commands.execute('sheet.merge.remove', { sheetId: activeSheetId, range });
          }
          refresh();
          break;
        }
        case 'textColor':
          if (typeof payload === 'string') applyStyleToPrimary({ textColor: payload });
          break;
        case 'background':
          if (typeof payload === 'string') applyStyleToPrimary({ background: payload });
          break;
        case 'numberFormat':
          if (typeof payload === 'string') applyStyleToPrimary({ numberFormat: payload });
          break;
        case 'format-currency':
          applyStyleToPrimary({ numberFormat: '$#,##0' });
          break;
        case 'format-percent':
          applyStyleToPrimary({ numberFormat: '0%' });
          break;
        case 'border-all':
          applyStyleToPrimary({
            borders: {
              top: { style: 'thin', color: '#334155' },
              right: { style: 'thin', color: '#334155' },
              bottom: { style: 'thin', color: '#334155' },
              left: { style: 'thin', color: '#334155' },
            },
          });
          break;
        case 'border-outer':
          applyStyleToPrimary({
            borders: {
              top: { style: 'medium', color: '#334155' },
              right: { style: 'medium', color: '#334155' },
              bottom: { style: 'medium', color: '#334155' },
              left: { style: 'medium', color: '#334155' },
            },
          });
          break;
        case 'border-none':
          applyStyleToPrimary({ borders: undefined });
          break;
        case 'clear-range':
          runtime.commands.execute('sheet.range.clear', { sheetId: activeSheetId, range: primaryRange });
          setFormulaDraft('');
          refresh();
          break;
        case 'clear-formats':
          runtime.commands.execute('sheet.range.clear', { sheetId: activeSheetId, range: primaryRange, mode: 'formats' });
          refresh();
          break;
        case 'autosum': {
          let sumStart = selection.primaryRowIndex - 1;
          while (sumStart >= 0) {
            const above = sheet.cells.get(sumStart, selection.primaryColumnIndex);
            if (!above || typeof above.value !== 'number') break;
            sumStart -= 1;
          }
          sumStart += 1;
          if (sumStart < selection.primaryRowIndex) {
            const label = columnLabelOf(selection.primaryColumnIndex);
            const formula = '=SUM(' + label + (sumStart + 1) + ':' + label + selection.primaryRowIndex + ')';
            setFormulaDraft(formula);
            overrideTargetRef.current = { row: selection.primaryRowIndex, column: selection.primaryColumnIndex };
            commitFormula(formula);
            overrideTargetRef.current = null;
          }
          break;
        }
        case 'function-wizard':
          setShowFunctionWizard(true);
          break;
        case 'sort-dialog':
          setShowSortDialog(true);
          break;
        case 'sort-asc':
        case 'sort-desc': {
          const ascending = action === 'sort-asc';
          runtime.commands.execute('sheet.sort.multi', {
            sheetId: activeSheetId,
            range: primaryRange.endRow > primaryRange.startRow || primaryRange.endColumn > primaryRange.startColumn
              ? primaryRange
              : normalizeRangeRef({
                  sheetId: activeSheetId,
                  startRow: 0,
                  endRow: Math.min(sheet.rowCount - 1, 30),
                  startColumn: 0,
                  endColumn: Math.min(columns.length - 1, 6),
                }),
            criteria: [{ column: selection.primaryColumnIndex, ascending }],
            hasHeader: true,
          });
          refresh();
          break;
        }
        case 'insert-row':
          insertRowsAtPrimary(1);
          break;
        case 'insert-column':
          insertColumnsAtPrimary(1);
          break;
        case 'delete-row':
          deleteRowsAtPrimary();
          break;
        case 'delete-column':
          deleteColumnsAtPrimary();
          break;
        case 'hide-row':
          hideRowsAtPrimary();
          break;
        case 'hide-column':
          hideColumnsAtPrimary();
          break;
        case 'unhide-all':
          unhideAll();
          break;
        case 'transpose':
          transposeSelection();
          break;
        case 'flip-h':
          flipSelection('h');
          break;
        case 'flip-v':
          flipSelection('v');
          break;
        case 'split-column':
          if (typeof payload === 'string') splitByDelimiter(payload);
          break;
        case 'banded-toggle':
          toggleBandedRows();
          break;
        case 'freeze-top-row':
          runtime.commands.execute('sheet.freeze.set', {
            sheetId: activeSheetId,
            freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
          });
          refresh();
          break;
        case 'freeze-first-col':
          runtime.commands.execute('sheet.freeze.set', {
            sheetId: activeSheetId,
            freeze: { xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1 },
          });
          refresh();
          break;
        case 'freeze-at-primary':
          runtime.commands.execute('sheet.freeze.set', {
            sheetId: activeSheetId,
            freeze: {
              xSplit: selection.primaryColumnIndex,
              ySplit: selection.primaryRowIndex,
              startRow: selection.primaryRowIndex,
              startColumn: selection.primaryColumnIndex,
            },
          });
          refresh();
          break;
        case 'unfreeze':
          runtime.commands.execute('sheet.freeze.set', {
            sheetId: activeSheetId,
            freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
          });
          refresh();
          break;
        case 'filter-clear':
          clearFilter();
          break;
        case 'apply-filter-selection': {
          const activeRange = selection.ranges[selection.primaryRangeIndex];
          if (activeRange) {
            const sheetModel = runtime.model.getSheet(activeSheetId);
            runtime.commands.execute('sheet.filter.set', {
              sheetId: activeSheetId,
              filter: {
                sheetId: activeSheetId,
                range: {
                  sheetId: activeSheetId,
                  startRow: 0,
                  endRow: Math.max(0, sheetModel.rowCount - 1),
                  startColumn: 0,
                  endColumn: Math.max(0, sheetModel.columnCount - 1),
                },
                criteria: {},
              },
            });
            refresh();
          }
          break;
        }
        case 'export-xlsx': {
          void (async () => {
            try {
              const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
              const base64 = buildXlsxArchiveBase64(snapshotResponse.snapshot);
              const link = document.createElement('a');
              link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
              link.download = (runtime.model.name || 'workbook') + '.xlsx';
              link.click();
              setNotice('Workbook exported as .xlsx');
            } catch {
              setNotice('Export failed');
            }
          })();
          break;
        }
        case 'import-xlsx': {
          const input = document.createElement('input');
          input.type = 'file';
          input.accept = '.xlsx';
          input.onchange = () => {
            const file = input.files?.[0];
            if (!file) return;
            void file.arrayBuffer().then((buffer) => {
              let binary = '';
              const bytes = new Uint8Array(buffer);
              const chunkSize = 0x8000;
              for (let i = 0; i < bytes.length; i += chunkSize) {
                binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
              }
              return btoa(binary);
              }).then((base64) => importXlsxBase64(base64)).catch(() => setNotice('Import failed'));
          };
          input.click();
          break;
        }
        case 'find-replace':
          setFindQuery(typeof payload === 'string' ? payload : '');
          setShowFindReplace(true);
          break;
        case 'zoom-in':
          setZoomState((current) => Math.min(200, current + 10));
          break;
        case 'zoom-out':
          setZoomState((current) => Math.max(50, current - 10));
          break;
        case 'zoom-100':
          setZoomState(100);
          break;
        case 'open-chart':
          setActivePanel('chart');
          break;
        case 'open-pivot':
          setActivePanel('pivot');
          if (sheet.pivots.length === 0) {
            const sourceRange = primaryRange.startRow !== primaryRange.endRow || primaryRange.startColumn !== primaryRange.endColumn
              ? primaryRange
              : usedRangeOfSheet(sheet);
            const fieldCatalog = buildPivotFieldCatalog(runtime.model, { id: 'pivot-source', sheetId: activeSheetId, sourceRange, layout: { rows: [], columns: [], filters: [], values: [], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false } }).fields;
            const rowField = fieldCatalog.find((field) => field.dataType !== 'number')?.id ?? fieldCatalog[0]?.id;
            const valueField = fieldCatalog.find((field) => field.dataType === 'number')?.id ?? fieldCatalog[0]?.id;
            if (rowField && valueField) {
              const summarizeBy = fieldCatalog.find((field) => field.id === valueField)?.dataType === 'number' ? 'sum' : 'count';
              runtime.commands.execute('pro.pivot.add', {
                id: 'pivot-' + Math.random().toString(36).slice(2, 8),
                sheetId: activeSheetId,
                sourceRange,
                layout: { rows: [{ field: rowField }], columns: [], filters: [], values: [{ field: valueField, summarizeBy }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
              });
              refresh();
            }
          }
          break;
        case 'create-data-table': {
          const sourceRange = primaryRange.startRow !== primaryRange.endRow || primaryRange.startColumn !== primaryRange.endColumn
            ? primaryRange
            : usedRangeOfSheet(sheet);
          const fieldNames = new Set<string>();
          const fields = [] as WorkbookTableModel['fields'];
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
          const table: WorkbookTableModel = {
            id: nextId('table'),
            name: `${sheet.name} table`,
            sourceSheetId: activeSheetId,
            rowCount: Math.max(0, sourceRange.endRow - sourceRange.startRow),
            fields,
            blockSize: 4096,
            blocks: [],
            revision: 0,
          };
          void (async () => {
            runtime.handlers.onSaveState?.('saving');
            await runtime.api.createDataTable(runtime.model.unitId, table);
            for (let startRow = sourceRange.startRow + 1; startRow <= sourceRange.endRow; startRow += table.blockSize) {
              const rows: TableScalar[][] = [];
              const endRow = Math.min(sourceRange.endRow, startRow + table.blockSize - 1);
              for (let row = startRow; row <= endRow; row++) {
                rows.push(fields.map((_field, offset) => sheet.cells.get(row, sourceRange.startColumn + offset)?.value ?? null));
              }
              if (rows.length > 0) table.blocks.push(await runtime.api.appendDataBlock(runtime.model.unitId, table.id, startRow - sourceRange.startRow - 1, rows));
            }
            table.revision = table.blocks.length;
            runtime.commands.execute('table.add', table);
            refresh();
            runtime.handlers.onSaveState?.('saved');
            runtime.handlers.onNotice?.(`Data table ${table.name} created`);
          })().catch((error: unknown) => {
            runtime.handlers.onSaveState?.('offline');
            runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Data table creation failed');
          });
          break;
        }
        case 'open-shape':
          setActivePanel('shape');
          break;
        case 'open-sparkline':
          setActivePanel('sparkline');
          break;
        case 'open-conditional-format':
          setActivePanel('conditionalFormat');
          break;
        case 'open-data-validation':
          setActivePanel('dataValidation');
          break;
        case 'open-history':
          setActivePanel('history');
          break;
        case 'open-print':
          setShowPrintPreviewState(true);
          setActivePanel('print');
          break;
        case 'open-comments':
          setActivePanel('inspector');
          setNotice('Comments are shown in the inspector panel');
          break;
        default:
          break;
      }
  };

  // ==== 动作实现(截断恢复) ====

  const columnLabelOf = (column: number): string => {
    let label = '';
    let remaining = column + 1;
    while (remaining > 0) {
      const modulo = (remaining - 1) % 26;
      label = String.fromCharCode(65 + modulo) + label;
      remaining = Math.floor((remaining - 1) / 26);
    }
    return label;
  };

  const singleCellRange = (row: number, column: number): RangeRef => ({
    sheetId: activeSheetId,
    startRow: row,
    endRow: row,
    startColumn: column,
    endColumn: column,
  });

  const cellStyleAt = (row: number, column: number): CellStyle | undefined =>
    phase === 'ready' ? runtime.model.getSheet(activeSheetId).cells.get(row, column)?.style : undefined;

  const primaryRangeOrDefault = (): RangeRef =>
    selection.ranges[selection.primaryRangeIndex] ?? singleCellRange(selection.primaryRowIndex, selection.primaryColumnIndex);

  function getRangeMatrixInternal(range: RangeRef): CellData[][] {
    const sheet = runtime.model.getSheet(range.sheetId);
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

  const copy = useCallback(() => {
    if (phase !== 'ready') return;
    void navigator.clipboard.writeText(formatTsv(getRangeMatrixInternal(primaryRangeOrDefault())));
    setNotice('Range copied');
  }, [phase, selection.primaryRangeIndex, selection.ranges]);

  const cut = useCallback(() => {
    if (phase !== 'ready') return;
    const range = primaryRangeOrDefault();
    void navigator.clipboard.writeText(formatTsv(getRangeMatrixInternal(range)));
    runtime.commands.execute('sheet.range.clear', { sheetId: activeSheetId, range, mode: 'contents' });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const paste = useCallback(() => {
    if (phase !== 'ready') return;
    void navigator.clipboard.readText().then((text) => {
      if (!text) return;
      const values = parseTsv(text);
      runtime.commands.execute('sheet.range.set', {
        sheetId: activeSheetId,
        startRow: selection.primaryRowIndex,
        startColumn: selection.primaryColumnIndex,
        values,
      });
      refresh();
      setNotice('Pasted from clipboard');
    });
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const clearFormats = useCallback(() => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.range.clear', { sheetId: activeSheetId, range: primaryRangeOrDefault(), mode: 'formats' });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const nextId = (prefix: string) => prefix + '-' + Math.random().toString(36).slice(2, 8);

  const addChart = useCallback((chart: ChartModel) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.chart.add', chart);
    refresh();
  }, [phase, refresh, runtime]);

  const addShape = useCallback((shape: ShapeModel) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.shape.add', shape);
    refresh();
  }, [phase, refresh, runtime]);

  const addSparkline = useCallback((sparkline: SparklineModel) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.sparkline.add', sparkline);
    refresh();
  }, [phase, refresh, runtime]);

  const updateChartBounds = useCallback((id: string, bounds: ChartModel['bounds']) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.chart.move', { id, sheetId: activeSheetId, bounds });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const removeChart = useCallback((id: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.chart.remove', id as never);
    refresh();
  }, [phase, refresh, runtime]);

  const updateShapeBounds = useCallback((id: string, bounds: ShapeModel['bounds']) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.shape.move', { id, sheetId: activeSheetId, bounds });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const removeShape = useCallback((id: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.shape.remove', id as never);
    refresh();
  }, [phase, refresh, runtime]);

  const removeSparkline = useCallback((id: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.sparkline.remove', id as never);
    refresh();
  }, [phase, refresh, runtime]);

  const addPivot = useCallback((pivot: PivotModel) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.pivot.add', pivot);
    refresh();
  }, [phase, refresh, runtime]);
  const refreshPivot = useCallback((pivotId: string) => {
    if (phase !== 'ready') return;
    const pivot = runtime.model.getSheet(activeSheetId).pivots.find((entry) => entry.id === pivotId);
    if (!pivot) return;
    runtime.commands.execute('pro.pivot.refresh', { sheetId: activeSheetId, pivotId });
    refresh();
    void runtime.api.calculatePivot(runtime.model.unitId, pivotId)
      .then((response) => {
        runtime.pivotResults[pivotId] = response.result;
        refresh();
      })
      .catch((error: unknown) => runtime.handlers.onNotice?.(error instanceof Error ? error.message : 'Pivot calculation failed'));
  }, [activeSheetId, phase, refresh, runtime]);

  const updatePivotLayout = useCallback((pivotId: string, layout: PivotLayout) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.pivot.update', { sheetId: activeSheetId, pivotId, layout });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const updatePivotConfiguration = useCallback((pivotId: string, patch: { sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[] }) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.pivot.update', { sheetId: activeSheetId, pivotId, ...patch });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const removePivot = useCallback((id: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('pro.pivot.remove', id as never);
    refresh();
  }, [phase, refresh, runtime]);

  const addConditionalFormat = useCallback((rule: ConditionalFormatRule) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.cf.add', { rule });
    refresh();
  }, [phase, refresh, runtime]);

  const removeConditionalFormat = useCallback((ruleId: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.cf.remove', { sheetId: activeSheetId, ruleId });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const addDataValidation = useCallback((rule: DataValidationRule) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.dv.add', { rule });
    refresh();
  }, [phase, refresh, runtime]);

  const removeDataValidation = useCallback((ruleId: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.dv.remove', { sheetId: activeSheetId, ruleId });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const addComment = useCallback((text: string) => {
    if (phase !== 'ready' || !text.trim()) return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    const mentions = [...text.matchAll(/@([\w-]+)/g)].map((match) => match[1]).filter((value): value is string => Boolean(value));
    const comment: CellComment = { id: nextId('cmt'), author: actorId, text: text.trim(), createdAt: new Date().toISOString(), mentions, replies: [], resolved: false };
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: { ...(cell ?? { value: null }), comment },
    });
    refresh();
  }, [actorId, activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const replyComment = useCallback((text: string) => {
    if (phase !== 'ready' || !text.trim()) return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    const comment = cell?.comment;
    if (!comment) return;
    const nextComment: CellComment = {
      ...structuredClone(comment),
      resolved: false,
      replies: [...(comment.replies ?? []), { id: nextId('reply'), author: actorId, text: text.trim(), createdAt: new Date().toISOString() }],
    };
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: { ...cell, comment: nextComment },
    });
    refresh();
  }, [actorId, activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const resolveComment = useCallback(() => {
    if (phase !== 'ready') return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    if (!cell?.comment) return;
    const comment: CellComment = { ...structuredClone(cell.comment), resolved: true, resolvedAt: new Date().toISOString() };
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: { ...cell, comment },
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const removeComment = useCallback(() => {
    if (phase !== 'ready') return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    if (!cell?.comment) return;
    const { comment: _removed, ...rest } = cell;
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: rest as CellData,
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const setHyperlink = useCallback((url: string) => {
    if (phase !== 'ready' || !url.trim()) return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: { ...(cell ?? { value: null }), hyperlink: url.trim() },
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const removeHyperlink = useCallback(() => {
    if (phase !== 'ready') return;
    const cell = runtime.model.getSheet(activeSheetId).cells.get(selection.primaryRowIndex, selection.primaryColumnIndex);
    if (!cell?.hyperlink) return;
    const { hyperlink: _dropped, ...rest } = cell;
    runtime.commands.execute('sheet.cell.set', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      value: rest as CellData,
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const applyFilter = useCallback(
    (column: number, patch: { selectedValues?: string[] | null; conditionOperator?: string; conditionValue?: string }) => {
      if (phase !== 'ready') return;
      const sheet = runtime.model.getSheet(activeSheetId);
      const baseRange = sheet.filter?.range
        ?? normalizeRangeRef({
          sheetId: activeSheetId,
          startRow: 0,
          endRow: Math.max(0, sheet.rowCount - 1),
          startColumn: 0,
          endColumn: Math.max(0, sheet.columnCount - 1),
        });
      const criteria: FilterModel['criteria'] = { ...(sheet.filter?.criteria ?? {}) };
      criteria[column] = {
        column,
        selectedValues: patch.selectedValues ?? undefined,
        conditionOperator: patch.conditionOperator,
        conditionValue: patch.conditionValue,
      };
      runtime.commands.execute('sheet.filter.set', {
        sheetId: activeSheetId,
        filter: { sheetId: activeSheetId, range: baseRange, criteria },
      });
      refresh();
    },
    [activeSheetId, phase, refresh, runtime],
  );

  const clearFilter = useCallback(() => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.filter.remove', { sheetId: activeSheetId });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const findReplace = useCallback(
    (params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }) => {
      if (phase !== 'ready' || !params.find) return 0;
      const patches = collectFindReplacements(runtime.model, params);
      let count = 0;
      for (const patch of patches) {
        count += patch.values[0]!.length;
        runtime.commands.execute('sheet.range.set', {
          sheetId: patch.sheetId,
          startRow: patch.startRow,
          startColumn: patch.startColumn,
          values: patch.values,
        });
      }
      refresh();
      setNotice(count + ' replacement(s) applied');
      return count;
    },
    [phase, refresh, runtime],
  );

  const insertRowsAtPrimary = useCallback((count: number) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.rows.insert', { sheetId: activeSheetId, at: selection.primaryRowIndex, count });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryRowIndex]);

  const deleteRowsAtPrimary = useCallback(() => {
    if (phase !== 'ready') return;
    const range = selection.ranges[selection.primaryRangeIndex];
    runtime.commands.execute('sheet.rows.delete', {
      sheetId: activeSheetId,
      at: range?.startRow ?? selection.primaryRowIndex,
      count: (range?.endRow ?? selection.primaryRowIndex) - (range?.startRow ?? selection.primaryRowIndex) + 1,
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection]);

  const insertColumnsAtPrimary = useCallback((count: number) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.columns.insert', { sheetId: activeSheetId, at: selection.primaryColumnIndex, count });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection.primaryColumnIndex]);

  const deleteColumnsAtPrimary = useCallback(() => {
    if (phase !== 'ready') return;
    const range = selection.ranges[selection.primaryRangeIndex];
    runtime.commands.execute('sheet.columns.delete', {
      sheetId: activeSheetId,
      at: range?.startColumn ?? selection.primaryColumnIndex,
      count: (range?.endColumn ?? selection.primaryColumnIndex) - (range?.startColumn ?? selection.primaryColumnIndex) + 1,
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection]);

  const hideRowsAtPrimary = useCallback(() => {
    if (phase !== 'ready') return;
    const range = selection.ranges[selection.primaryRangeIndex];
    const start = range?.startRow ?? selection.primaryRowIndex;
    const end = range?.endRow ?? selection.primaryRowIndex;
    for (let row = start; row <= end; row++) {
      runtime.commands.execute('sheet.row.hide', { sheetId: activeSheetId, index: row });
    }
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection]);

  const hideColumnsAtPrimary = useCallback(() => {
    if (phase !== 'ready') return;
    const range = selection.ranges[selection.primaryRangeIndex];
    const start = range?.startColumn ?? selection.primaryColumnIndex;
    const end = range?.endColumn ?? selection.primaryColumnIndex;
    for (let column = start; column <= end; column++) {
      runtime.commands.execute('sheet.column.hide', { sheetId: activeSheetId, index: column });
    }
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection]);

  const unhideAll = useCallback(() => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.rows.unhide.all', { sheetId: activeSheetId });
    runtime.commands.execute('sheet.columns.unhide.all', { sheetId: activeSheetId });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const toggleBandedRows = useCallback(() => {
    if (phase !== 'ready') return;
    const sheet = runtime.model.getSheet(activeSheetId);
    const next = sheet.bandedRule
      ? null
      : {
          range: normalizeRangeRef({
            sheetId: activeSheetId,
            startRow: 0,
            endRow: Math.max(0, sheet.rowCount - 1),
            startColumn: 0,
            endColumn: Math.max(0, sheet.columnCount - 1),
          }),
        firstColor: '#ffffff',
        secondColor: '#f1f5f9',
      };
    runtime.commands.execute('sheet.banded.set', { sheetId: activeSheetId, rule: next });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const transposeSelection = useCallback(() => {
    if (phase !== 'ready') return;
    runtime.commands.execute('matrix.transpose', { sheetId: activeSheetId, range: primaryRangeOrDefault() });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const flipSelection = useCallback((direction: 'h' | 'v') => {
    if (phase !== 'ready') return;
    runtime.commands.execute('matrix.flip', {
      sheetId: activeSheetId,
      range: primaryRangeOrDefault(),
      direction: direction === 'h' ? 'horizontal' : 'vertical',
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const splitByDelimiter = useCallback((delimiter: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.splitColumn', {
      sheetId: activeSheetId,
      row: selection.primaryRowIndex,
      column: selection.primaryColumnIndex,
      delimiter,
      maxColumns: Math.min(columns.length - selection.primaryColumnIndex - 1, 8),
    });
    refresh();
  }, [activeSheetId, columns.length, phase, refresh, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const defineName = useCallback((name: string, value: string) => {
    if (phase !== 'ready' || !name.trim()) return;
    runtime.commands.execute('workbook.name.set', { name: name.trim().toUpperCase(), value });
    refresh();
  }, [phase, refresh, runtime]);

  const removeName = useCallback((name: string) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('workbook.name.remove', { name });
    refresh();
  }, [phase, refresh, runtime]);

  const printWorkbook = useCallback((layout: PrintLayout) => {
    setPrintLayout(layout);
    setShowPrintPreviewState(true);
    setNotice('Preparing print preview');
  }, []);

  const exportPdf = useCallback((layout: PrintLayout) => {
    setPrintLayout(layout);
    setShowPrintPreviewState(true);
    setNotice('Choose Save as PDF in the print dialog');
  }, []);

  const importXlsxBase64 = useCallback(async (base64: string) => {
    const response = await fetch('/api/v1/files/import-xlsx', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ base64 }),
    });
    if (!response.ok) throw new Error('Import failed: ' + response.status);
    const result = (await response.json()) as SnapshotResponse;
    hydrateRuntime(runtime, result);
    runtime.remoteConnected = true;
    setActiveSheetId(runtime.model.activeSheetId);
    setSelection(createInitialSelection());
    refresh();
    setNotice('XLSX imported');
    setPhase('ready');
  }, [refresh, runtime]);

  const sortRange = useCallback(
    (criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader: boolean) => {
      if (phase !== 'ready' || criteria.length === 0) return;
      const range = selection.ranges[selection.primaryRangeIndex]
        ?? normalizeRangeRef({
          sheetId: activeSheetId,
          startRow: 0,
          endRow: Math.min(40, selectedSheet.rowCount - 1),
          startColumn: 0,
          endColumn: Math.min(columns.length - 1, 6),
        });
      runtime.commands.execute('sheet.sort.multi', {
        sheetId: activeSheetId,
        range,
        criteria: criteria.map((criterion) => ({ column: criterion.colIdx, ascending: criterion.ascending })),
        hasHeader,
      });
      refresh();
    },
    [activeSheetId, columns.length, phase, refresh, runtime, selectedSheet.rowCount, selection],
  );

  const getValidationForPrimary = useCallback((): DataValidationRule | undefined => {
    if (phase !== 'ready') return undefined;
    const sheet = runtime.model.getSheet(activeSheetId);
    return findValidationRule(sheet, selection.primaryRowIndex, selection.primaryColumnIndex);
  }, [activeSheetId, phase, runtime, selection.primaryColumnIndex, selection.primaryRowIndex]);

  const getValidationAt = useCallback((row: number, column: number): string[] | undefined => {
    if (phase !== 'ready') return undefined;
    const sheet = runtime.model.getSheet(activeSheetId);
    const rule = findValidationRule(sheet, row, column);
    if (!rule) return undefined;
    return validationList(rule);
  }, [activeSheetId, phase, runtime]);

  const getActiveSheetName = useCallback(() => runtime.model.getSheet(activeSheetId).name, [activeSheetId, runtime]);

  const readDataTable = useCallback((tableId: string, offset = 0, limit = 100): Promise<TableRowsResponse> => {
    return runtime.api.readDataRows(runtime.model.unitId, tableId, offset, limit);
  }, [runtime]);

  const getPivotFieldCatalog = useCallback((range: RangeRef): PivotFieldDefinition[] => {
    const pivot: PivotModel = {
      id: 'pivot-field-catalog',
      sheetId: range.sheetId,
      sourceRange: range,
      layout: { rows: [], columns: [], filters: [], values: [], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
    };
    return buildPivotFieldCatalog(runtime.model, pivot).fields;
  }, [runtime]);

  const showPivotDetails = useCallback((paths: PivotSourceRowPath[]) => {
    if (phase !== 'ready' || paths.length === 0) return;
    const first = paths[0]!;
    const source = runtime.model.getSheet(first.sheetId);
    const id = 'sheet-' + Math.random().toString(36).slice(2, 8);
    runtime.commands.execute('sheet.add', { id, name: 'Pivot Details' });
    const values: CellData[][] = [];
    values.push(Array.from({ length: source.columnCount }, (_, column) => structuredClone(source.cells.get(0, column) ?? { value: null })));
    for (const path of paths) {
      const rowSheet = runtime.model.getSheet(path.sheetId);
      values.push(Array.from({ length: source.columnCount }, (_, column) => structuredClone(rowSheet.cells.get(path.row, column) ?? { value: null })));
    }
    if (values.length > 0) runtime.commands.execute('sheet.range.set', { sheetId: id, startRow: 0, startColumn: 0, values });
    setActiveSheetId(id);
    refresh();
  }, [phase, refresh, runtime]);

  const getRangeMatrix = useCallback((range: RangeRef): CellData[][] => {
    const sheet = runtime.model.getSheet(range.sheetId);
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const rowCells: CellData[] = [];
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        rowCells.push(structuredClone(sheet.cells.get(r, c)) ?? { value: null });
      }
      rows.push(rowCells);
    }
    return rows;
  }, [runtime]);

  const getRangeNumbers = useCallback((range: RangeRef): number[] => {
    const numbers: number[] = [];
    for (const row of getRangeMatrix(range)) {
      for (const cell of row) {
        const numeric = typeof cell.value === 'number'
          ? cell.value
          : Number(String(cell.value ?? '').replace(/[$,%]/g, ''));
        if (Number.isFinite(numeric) && cell.value !== '' && cell.value != null) numbers.push(numeric);
      }
    }
    return numbers;
  }, [getRangeMatrix]);

  const setSelectedFloatingId = useCallback((id: string | null) => setSelectedFloatingIdState(id), []);
  const removeFloatingObject = useCallback(
    (kind: 'chart' | 'shape', id: string) => {
      if (kind === 'chart') removeChart(id);
      else removeShape(id);
      setSelectedFloatingIdState(null);
    },
    [removeChart, removeShape],
  );

  const resizeRow = useCallback((row: number, heightPx: number) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.row.resize', { sheetId: activeSheetId, row, height: Math.max(18, heightPx) });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const resizeColumn = useCallback((column: number, widthPx: number) => {
    if (phase !== 'ready') return;
    runtime.commands.execute('sheet.column.resize', { sheetId: activeSheetId, column, width: Math.max(24, widthPx) });
    refresh();
  }, [activeSheetId, phase, refresh, runtime]);

  const fillRange = useCallback((targetRange: { startRow: number; endRow: number; startColumn: number; endColumn: number }) => {
    if (phase !== 'ready') return;
    const primary = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    if (!primary) return;
    runtime.commands.execute('sheet.autofill', {
      sheetId: activeSheetId,
      sourceRange: { sheetId: activeSheetId, startRow: primary.startRow, endRow: primary.endRow, startColumn: primary.startColumn, endColumn: primary.endColumn },
      targetRange: { sheetId: activeSheetId, ...targetRange },
    });
    refresh();
  }, [activeSheetId, phase, refresh, runtime, selection]);

  const addSheet = useCallback(() => {
    if (phase !== 'ready') return;
    const id = 'sheet-' + Math.random().toString(36).slice(2, 8);
    runtime.commands.execute('sheet.add', { id, name: 'Sheet' + (runtime.model.getSheets().length + 1) });
    setActiveSheetId(id);
    refresh();
  }, [phase, refresh, runtime]);

  const renameSheet = useCallback((sheetId: string, name: string) => {
    if (phase !== 'ready' || !name.trim()) return;
    runtime.commands.execute('sheet.rename', { sheetId, name: name.trim() });
    refresh();
  }, [phase, refresh, runtime]);

  const deleteSheet = useCallback((sheetId: string) => {
    if (phase !== 'ready') return;
    try {
      runtime.commands.execute('sheet.remove', { id: sheetId });
      if (activeSheetId === sheetId) {
        const remaining = runtime.model.getSheets()[0];
        if (remaining) setActiveSheetId(remaining.id);
      }
      refresh();
    } catch {
      setNotice('A workbook must keep at least one sheet');
    }
  }, [activeSheetId, phase, refresh, runtime]);


  const closeFunctionWizard = useCallback(() => setShowFunctionWizard(false), []);
  const closeSortDialog = useCallback(() => setShowSortDialog(false), []);
  const closeFindReplace = useCallback(() => {
    setShowFindReplace(false);
    setFindQuery('');
  }, []);
  const setShowPrintPreview = useCallback((open: boolean) => setShowPrintPreviewState(open), []);

  const actions: WorkspaceActions = {
      selectCell,
      selectRange,
      movePrimary,
      jumpEdge,
      selectAll,
      selectRowHeader,
      selectColumnHeader,
      beginEdit,
      cancelEdit,
      commitEdit,
      setFormulaDraft,
      insertRefIntoDraft,
      toggleAbsoluteReference,
      moveCell,
      notify: setNotice as never,
      redo,
      retry,
      selectSheet,
      setActivePanel,
      setRibbonTab,
      undo,
      handleRibbonAction,
      addChart,
      updateChartBounds,
      removeChart,
      addPivot,
      updatePivotLayout,
      updatePivotConfiguration,
      refreshPivot,
      removePivot,
      addShape,
      updateShapeBounds,
      removeShape,
      addSparkline,
      removeSparkline,
      addConditionalFormat,
      removeConditionalFormat,
      addDataValidation,
      removeDataValidation,
      addComment,
      replyComment,
      resolveComment,
      removeComment,
      setHyperlink,
      removeHyperlink,
      applyFilter,
      clearFilter,
      findReplace,
      insertRowsAtPrimary,
      deleteRowsAtPrimary,
      insertColumnsAtPrimary,
      deleteColumnsAtPrimary,
      hideRowsAtPrimary,
      hideColumnsAtPrimary,
      unhideAll,
      toggleBandedRows,
      transposeSelection,
      flipSelection,
      splitByDelimiter,
      defineName,
      removeName,
      printWorkbook,
      exportPdf,
      importXlsxBase64,
      sortRange,
      getRangeMatrix,
      getRangeNumbers,
      getValidationForPrimary,
      getValidationAt,
      cut,
      copy,
      paste,
      clearFormats,
      addSheet,
      renameSheet,
      deleteSheet,
      resizeRow,
      resizeColumn,
      fillRange,
      setSelectedFloatingId,
      removeFloatingObject,
      getActiveSheetName,
      getPivotFieldCatalog,
      readDataTable,
      showPivotDetails,
      setZoom: (nextZoom: number) => setZoomState(Math.max(50, Math.min(200, nextZoom))),
      commitFormula: (overrideValue?: string) => {
        if (overrideValue !== undefined) {
          overrideTargetRef.current = { row: selection.primaryRowIndex, column: selection.primaryColumnIndex };
        }
        commitFormula(overrideValue);
        overrideTargetRef.current = null;
      },
      closeFunctionWizard,
      closeSortDialog,
      closeFindReplace,
      setShowPrintPreview,
  };

  const historyEntries = useMemo(() => runtime.commands.getUndoEntries(), [modelVersion, runtime]);

  const state = useMemo<WorkspaceState>(() => ({
    unitId: runtime.model.unitId,
    workbookName: runtime.model.name,
    selectedFloatingId,
    selection,
    activeCell,
    activePanel,
    activeSheetId,
    formulaDraft,
    editingCell,
    sheets,
    selectedSheet,
    ribbonTab,
    saveState,
    notice,
    phase,
    zoom,
    peers,
    historyEntries,
    remoteRevisions,
    tables,
    showFunctionWizard,
    showSortDialog,
    showFindReplace,
    findQuery,
    showPrintPreview,
    printLayout,
    actorId,
    collabStatus,
  }), [activeCell, activePanel, activeSheetId, actorId, collabStatus, editingCell, findQuery, formulaDraft, historyEntries, modelVersion, notice, peers, phase, printLayout, remoteRevisions, ribbonTab, saveState, selectedFloatingId, selectedSheet, selection, sheets, showFindReplace, showFunctionWizard, showPrintPreview, showSortDialog, tables, zoom]);

  return { actions, state };
}

