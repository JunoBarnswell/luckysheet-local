        row.map((value) => {
          if (value === '') return { value: null };
          const numeric = /^-?\d+(\.\\d+)?\$/.test(value) ? Number(value) : null;
          return { value: numeric ?? value };
        }),
  type CellData,
  type CellStyle,
  type ChartModel,
  type ConditionalFormatRule,
  type DataValidationRule,
  type FilterModel,
  type FreezeModel,
  type MergeSpan,
  type PivotModel,
  type RangeRef,
  type ShapeModel,
  type SparklineModel,
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
import { WorkbookApiClient, type CollaborationMutation, type SnapshotResponse } from '@react-sheets/protocol';
import { buildXlsxArchiveBase64, registerProSheetCommands, type PrintLayout } from '@react-sheets/pro-features';

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
  | 'comments'
  | 'data'
  | 'automations';
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
  name: string;
  rows: SheetRow[];
  charts: ChartModel[];
  pivots: PivotModel[];
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
  showFunctionWizard: boolean;
  showSortDialog: boolean;
  showFindReplace: boolean;
  showPrintPreview: boolean;
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
  selectRowHeader: (row: number, mode?: 'replace' | 'add') => void;
  selectColumnHeader: (column: number, mode?: 'replace' | 'add') => void;
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
  handleRibbonAction: (action: string, payload?: unknown) => void;
  // Pro 模型
  addChart: (chart: ChartModel) => void;
  updateChartBounds: (id: string, bounds: ChartModel['bounds']) => void;
  removeChart: (id: string) => void;
  addPivot: (pivot: PivotModel) => void;
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
}

const UNIT_ID_STORAGE_KEY = 'react-sheets:unitId';
const ACTOR_ID_STORAGE_KEY = 'react-sheets:actorId';

const PEER_COLORS = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];

function resolveUnitId(): string {
  if (typeof window === 'undefined') return 'wb-server-default';
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

function seedWorkbook(runtime: WorkspaceRuntime): void {
  const values: CellData[][] = [
    [
      { value: 'Q3 Growth Plan', style: { bold: true, fontSize: 13, background: '#f1f5f9' } },
      { value: 'Owner', style: { bold: true, background: '#f1f5f9' } },
      { value: 'Status', style: { bold: true, background: '#f1f5f9' } },
      { value: 'Target', style: { bold: true, background: '#f1f5f9' } },
      { value: 'Actual', style: { bold: true, background: '#f1f5f9' } },
      { value: 'Variance', style: { bold: true, background: '#f1f5f9' } },
    ],
    [
      { value: 'User Activation' },
      { value: 'Maya Chen' },
      { value: 'On track' },
      { value: 0.42, numberFormat: '0%' },
      { value: 0.38, numberFormat: '0%' },
      { value: -0.04, numberFormat: '0%' },
    ],
    [
      { value: 'Retention Rate' },
      { value: 'Noah Williams' },
      { value: 'Needs review' },
      { value: 0.68, numberFormat: '0%' },
      { value: 0.64, numberFormat: '0%' },
      { value: -0.04, numberFormat: '0%' },
    ],
    [
      { value: 'Enterprise Expansion' },
      { value: 'Ava Patel' },
      { value: 'On track' },
      { value: 120000, numberFormat: '$#,##0' },
      { value: 132000, numberFormat: '$#,##0' },
      { value: 0.1, numberFormat: '0%' },
    ],
    [
      { value: 'Referral Engine' },
      { value: 'Liam Garcia' },
      { value: 'At risk' },
      { value: 0.16, numberFormat: '0%' },
      { value: 0.11, numberFormat: '0%' },
      { value: -0.05, numberFormat: '0%' },
    ],
    [
      { value: 'Quarter Total', style: { bold: true } },
      { value: 'Team Aggregate' },
      { value: 'Active' },
      { value: null, formula: '=SUM(D2:D5)' },
      { value: null, formula: '=SUM(E2:E5)' },
      { value: null, formula: '=D6-E6' },
    ],
  ];

  runtime.commands.execute('sheet.range.set', {
    sheetId: 'sheet-1',
    startRow: 0,
    startColumn: 0,
    values,
  });

  for (let row = 0; row < values.length; row += 1) {
    for (let column = 0; column < (values[row]?.length ?? 0); column += 1) {
      const cell = values[row]?.[column];
      if (!cell) continue;
      const address = { sheetId: 'sheet-1', row, column };
      if (cell.formula) runtime.formula.setFormula(address, cell.formula);
      else if (cell.value != null) runtime.formula.setValue(address, cell.value);
    }
  }

  runtime.commands.clearHistory();
}

function createWorkspaceRuntime(): WorkspaceRuntime {
  const model = new WorkbookModel(resolveUnitId(), 'Q3 Growth Planning');
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
  };
  attachCoreListeners(runtime);
  seedWorkbook(runtime);
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
  if (!runtime.remoteConnected) return;
  runtime.ownOperationIds.add(operationId);
  const changeSet = {
    schema: 'CollaborationChangeSetV1' as const,
    operationId,
    unitId: runtime.model.unitId,
    actorId: resolveActorId(),
    baseRevision: runtime.remoteRevision,
    mutations,
    createdAt: new Date().toISOString(),
  };
  runtime.handlers.onSaveState?.('saving');
  void runtime.api
    .submitChangeSet(changeSet)
    .then((result) => {
      runtime.remoteRevision = Math.max(runtime.remoteRevision, result.revision);
      runtime.handlers.onSaveState?.('saved');
    })
    .catch(() => {
      runtime.handlers.onSaveState?.('offline');
    });
}

// ---------- 启动恢复 ----------

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.activeSheetId });
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

function toSheetView(
  sheet: WorksheetModel,
  formula: FormulaEngine,
  showInvalid: boolean,
): SheetView {
  const overlays = computeConditionalOverlays(sheet);
  const filterHidden = computeFilterHiddenRows(sheet);
  const hiddenRows = new Set<number>([...sheet.hiddenRows, ...filterHidden]);
  const filterColumns = sheet.filter ? Object.keys(sheet.filter.criteria).map(Number) : [];
  const rows: SheetRow[] = [];

  const totalRows = Math.max(60, sheet.rowCount);
  for (let row = 0; row < totalRows; row += 1) {
    if (hiddenRows.has(row)) continue;
    const cells: SheetCell[] = [];
    for (let column = 0; column < columns.length; column += 1) {
      const modelCell = sheet.cells.get(row, column);
      const value = formatDisplayValue(modelCell, formula, sheet.id, row, column);
      const key = `${row}:${column}`;
      const overlay = overlays.get(key);
      const style = overlay?.style
        ? { ...(modelCell?.style ?? {}), ...overlay.style }
        : modelCell?.style;

      const validation = validateDataInput(sheet, row, column, modelCell?.value ?? null);

      cells.push({
        address: cellAddress(row, column),
        formula: modelCell?.formula,
        style,
        value,
        hasComment: Boolean(modelCell?.comment),
        invalid: showInvalid && modelCell?.value != null && !validation.valid,
        hyperlink: modelCell?.hyperlink,
        overlay,
      });
    }
    rows.push({ rowNumber: row + 1, cells, height: sheet.rowHeights[row] ?? 28 });
  }

  const isEmpty = sheet.cells.count() === 0;

  return {
    id: sheet.id,
    name: sheet.name,
    columns,
    rows,
    isEmpty,
    charts: [...sheet.charts],
    pivots: [...sheet.pivots],
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
    rowCount: totalRows,
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

  const [activeSheetId, setActiveSheetId] = useState(runtime.model.activeSheetId);

  const [showFunctionWizard, setShowFunctionWizard] = useState(false);
  const [showSortDialog, setShowSortDialog] = useState(false);
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [showPrintPreview, setShowPrintPreviewState] = useState(false);
  const [peers, setPeers] = useState<PeerCursor[]>([]);
  const [collabStatus, setCollabStatus] = useState<'connecting' | 'open' | 'closed'>('closed');

  const actorId = useMemo(() => resolveActorId(), []);

  const refresh = useCallback(() => setModelVersion((version) => version + 1), []);

  // 将回调注入 runtime(每次渲染更新,保持最新闭包)
  useEffect(() => {
    runtime.handlers.onSaveState = setSaveState;
    runtime.handlers.onNotice = setNotice;
    runtime.handlers.onMutationsApplied = () => refresh();
  }, [runtime, refresh]);

  // 启动恢复 / 创建 / 离线回退
  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
        hydrateRuntime(runtime, snapshotResponse);
        runtime.remoteConnected = true;
        if (!active) return;
        runtime.handlers.onSaveState?.('saved');
        runtime.handlers.onNotice?.('Workbook restored from server');
        setActiveSheetId(runtime.model.activeSheetId);
        refresh();
      } catch {
        try {
          const created = await runtime.api.createWorkbook(runtime.model.snapshot());
          runtime.remoteConnected = true;
          runtime.remoteRevision = Math.max(runtime.remoteRevision, created.revision);
          if (!active) return;
          runtime.handlers.onSaveState?.('saved');
          runtime.handlers.onNotice?.('SQLite sync connected');
          refresh();
        } catch {
          if (!active) return;
          runtime.handlers.onSaveState?.('offline');
          runtime.handlers.onNotice?.('Running local in-memory engine');
        }
      }
    })();
    return () => {
      active = false;
    };
  }, [runtime, refresh]);

  const sheets = useMemo(
    () => runtime.model.getSheets().map((sheet) => toSheetView(sheet, runtime.formula, true)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [modelVersion, runtime],
  );

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
    (row: number, mode: 'replace' | 'add' = 'replace') => {
      selectRange({ startRow: row, startColumn: 0, endRow: row, endColumn: columns.length - 1 }, mode);
    },
    [selectRange],
  );

  const selectColumnHeader = useCallback(
    (column: number, mode: 'replace' | 'add' = 'replace') => {
      selectRange({ startRow: 0, startColumn: column, endRow: selectedSheet.rowCount - 1, endColumn: column }, mode);
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

  const handleRibbonAction = useCallback(
    (action: string, payload?: unknown) => {
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
          runtime.commands.execute('sheet.style.clear', { sheetId: activeSheetId, range: primaryRange });
          refresh();
          break;
        case 'autosum': {
          const rowLabel = selection.primaryRowIndex + 1;
          const endLabel = columnLabel(Math.max(0, selection.primaryColumnIndex - 1));
          const formula = ` =SUM(A${rowLabel}:${endLabel}${rowLabel})`.trim();
          setFormulaDraft(formula);
          overrideTargetRef.current = { row: selection.primaryRowIndex, column: selection.primaryColumnIndex };
          commitFormula(formula);
          overrideTargetRef.current = null;
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
          actionsProxy.current.insertRowsAtPrimary?.(1);
          break;
        case 'insert-column':
          actionsProxy.current.insertColumnsAtPrimary?.(1);
          break;
        case 'delete-row':
          actionsProxy.current.deleteRowsAtPrimary?.();
          break;
        case 'delete-column':
          actionsProxy.current.deleteColumnsAtPrimary?.();
          break;
        case 'hide-row':
          actionsProxy.current.hideRowsAtPrimary?.();
          break;
        case 'hide-column':
          actionsProxy.current.hideColumnsAtPrimary?.();
          break;
        case 'unhide-all':
          actionsProxy.current.unhideAll?.();
          break;
        case 'transpose':
          actionsProxy.current.transposeSelection?.();
          break;
        case 'flip-h':
          actionsProxy.current.flipSelection?.('h');
          break;
        case 'flip-v':
          actionsProxy.current.flipSelection?.('v');
          break;
        case 'split-column':
          if (typeof payload === 'string') actionsProxy.current.splitByDelimiter?.(payload);
          break;
        case 'banded-toggle':
          actionsProxy.current.toggleBandedRows?.();
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
          actionsProxy.current.clearFilter?.();
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
            }).then((base64) => actionsProxy.current.importXlsxBase64?.(base64));
          };
          input.click();
          break;
        }
        case 'find-replace':
          setShowFindReplace(true);
          break;
        case 'banded':