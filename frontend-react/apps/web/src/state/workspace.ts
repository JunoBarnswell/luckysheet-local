import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  CellMatrix,
  WorkbookModel,
  type CellData,
  type CellStyle,
  type ChartModel,
  type ConditionalFormatRule,
  type DataValidationRule,
  type FreezeModel,
  type MergeSpan,
  type PivotModel,
  type RangeRef,
  type ShapeModel,
  type SheetId,
  type SparklineModel,
  type WorksheetModel,
} from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry } from '@react-sheets/command-runtime';
import {
  copyRangeToClipboardData,
  parseTsv,
  formatTsv,
  registerSheetCommands,
} from '@react-sheets/sheet-features';
import { FormulaEngine, isFormulaError, type FormulaValue } from '@react-sheets/formula-engine';
import { WorkbookApiClient, type CollaborationMutation } from '@react-sheets/protocol';
import {
  exportSnapshotToXlsxXml,
  paginateRange,
  registerProSheetCommands,
  type PrintLayout,
} from '@react-sheets/pro-features';

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
export type SaveState = 'saved' | 'saving' | 'offline';
export type CellTone = 'accent' | 'header' | 'muted' | 'plain' | 'total';

export interface SheetCell {
  address: string;
  displayValue?: string;
  formula?: string;
  style?: CellStyle;
  tone?: CellTone;
  value: string;
}

export interface SheetRow {
  cells: SheetCell[];
  rowNumber: number;
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
}

export interface WorkspaceState {
  activeCell: string;
  activePanel: SidebarPanelId;
  activeSheetId: string;
  formulaDraft: string;
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
}

export interface WorkspaceActions {
  addSheet: () => void;
  deleteSheet: (sheetId: string) => void;
  renameSheet: (sheetId: string, name: string) => void;
  commitFormula: (overrideValue?: string) => void;
  moveCell: (address: string, direction: 'down' | 'left' | 'right' | 'up') => void;
  notify: (message: string) => void;
  redo: () => void;
  retry: () => void;
  selectCell: (address: string) => void;
  selectSheet: (sheetId: string) => void;
  setActivePanel: (panel: SidebarPanelId) => void;
  setFormulaDraft: (value: string) => void;
  setRibbonTab: (tab: RibbonTabId) => void;
  setZoom: (zoom: number) => void;
  undo: () => void;
  handleRibbonAction: (action: string, payload?: unknown) => void;
  addChart: (chart: ChartModel) => void;
  removeChart: (id: string) => void;
  addPivot: (pivot: PivotModel) => void;
  removePivot: (id: string) => void;
  addShape: (shape: ShapeModel) => void;
  removeShape: (id: string) => void;
  addSparkline: (sparkline: SparklineModel) => void;
  removeSparkline: (id: string) => void;
  addConditionalFormat: (rule: ConditionalFormatRule) => void;
  removeConditionalFormat: (id: string) => void;
  addDataValidation: (rule: DataValidationRule) => void;
  removeDataValidation: (id: string) => void;
  printWorkbook: (layout: PrintLayout) => void;
  exportPdf: (layout: PrintLayout) => void;
  closeFunctionWizard: () => void;
  closeSortDialog: () => void;
  sortRange: (colIdx: number, ascending: boolean, hasHeader: boolean) => void;
}

export interface UseWorkspaceStateOptions {
  initialPhase?: WorkspacePhase;
}

interface WorkspaceRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  model: WorkbookModel;
  commands: CommandRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
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
      else runtime.formula.setValue(address, cell.value == null ? null : cell.value);
    }
  }

  runtime.commands.clearHistory();
}

function createWorkspaceRuntime(): WorkspaceRuntime {
  const model = new WorkbookModel(
    typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() : 'wb-default',
    'Q3 Growth Planning',
  );
  const commands = new CommandRuntime(model);
  registerSheetCommands(commands);
  registerProSheetCommands(commands);
  const runtime = {
    api: new WorkbookApiClient(),
    model,
    commands,
    formula: new FormulaEngine({ defaultSheetId: 'sheet-1' }),
    remoteConnected: false,
    remoteRevision: 0,
  };
  seedWorkbook(runtime);
  return runtime;
}

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
  if (typeof cell.value === 'number' && cell.numberFormat === '0%') {
    return `${Math.round(cell.value * 100)}%`;
  }
  if (typeof cell.value === 'number' && cell.numberFormat === '$#,##0') {
    return `$${cell.value.toLocaleString('en-US')}`;
  }
  return String(cell.value);
}

function toSheetView(sheet: WorksheetModel, formula: FormulaEngine): SheetView {
  const rows: SheetRow[] = [];
  for (let row = 0; row < Math.max(30, sheet.rowCount); row += 1) {
    const cells: SheetCell[] = [];
    for (let column = 0; column < columns.length; column += 1) {
      const modelCell = sheet.cells.get(row, column);
      const value = formatDisplayValue(modelCell, formula, sheet.id, row, column);
      const tone: CellTone | undefined =
        row === 0
          ? 'header'
          : modelCell?.value === 'On track'
            ? 'accent'
            : modelCell?.value === 'At risk' || modelCell?.value === 'Needs review'
              ? 'total'
              : undefined;

      cells.push({
        address: cellAddress(row, column),
        formula: modelCell?.formula,
        style: modelCell?.style,
        tone,
        value,
      });
    }
    rows.push({ rowNumber: row + 1, cells });
  }

  return {
    id: sheet.id,
    name: sheet.name,
    columns,
    rows,
    charts: [...sheet.charts],
    pivots: [...sheet.pivots],
    shapes: [...sheet.shapes],
    sparklines: [...sheet.sparklines],
    conditionalFormats: [...sheet.conditionalFormats],
    dataValidations: [...sheet.dataValidations],
    merges: [...sheet.merges],
    freeze: { ...sheet.freeze },
  };
}

export function getInitialWorkspacePhase(): WorkspacePhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'loading' || queryPhase === 'error' || queryPhase === 'empty'
    ? queryPhase
    : 'ready';
}

export function useWorkspaceState({ initialPhase = 'ready' }: UseWorkspaceStateOptions = {}): {
  actions: WorkspaceActions;
  state: WorkspaceState;
} {
  const runtimeRef = useRef<WorkspaceRuntime | null>(null);
  if (!runtimeRef.current) runtimeRef.current = createWorkspaceRuntime();
  const runtime = runtimeRef.current;

  const [phase, setPhase] = useState<WorkspacePhase>(initialPhase);
  const [activeSheetId, setActiveSheetId] = useState(runtime.model.activeSheetId);
  const [activeCell, setActiveCell] = useState('E4');
  const [formulaDraft, setFormulaDraft] = useState('132000');
  const [ribbonTab, setRibbonTab] = useState<RibbonTabId>('home');
  const [activePanel, setActivePanel] = useState<SidebarPanelId>('inspector');
  const [zoom, setZoomState] = useState(100);
  const [saveState, setSaveState] = useState<SaveState>('saved');
  const [notice, setNotice] = useState('Workbook engine ready');
  const [modelVersion, setModelVersion] = useState(0);

  const [showFunctionWizard, setShowFunctionWizard] = useState(false);
  const [showSortDialog, setShowSortDialog] = useState(false);

  const refresh = useCallback(() => setModelVersion((version) => version + 1), []);

  useEffect(() => {
    let active = true;
    void runtime.api
      .createWorkbook(runtime.model.snapshot())
      .then((response) => {
        if (!active) return;
        runtime.remoteConnected = true;
        runtime.remoteRevision = response.revision;
        setNotice('SQLite sync connected');
      })
      .catch(() => {
        if (!active) return;
        setSaveState('offline');
        setNotice('Running local in-memory & WAL engine');
      });
    return () => {
      active = false;
    };
  }, [runtime]);

  const persistMutation = useCallback(
    (operationId: string, mutations: CollaborationMutation[]) => {
      if (!runtime.remoteConnected) {
        return;
      }
      const changeSet = {
        schema: 'CollaborationChangeSetV1' as const,
        operationId,
        unitId: runtime.model.unitId,
        actorId: 'react-sheets-user',
        baseRevision: runtime.remoteRevision,
        mutations,
        createdAt: new Date().toISOString(),
      };
      void runtime.api
        .submitChangeSet(changeSet)
        .then((result) => {
          runtime.remoteRevision = result.revision;
          setSaveState('saved');
        })
        .catch(() => {
          setSaveState('offline');
        });
    },
    [runtime],
  );

  const sheets = useMemo(
    () => runtime.model.getSheets().map((sheet) => toSheetView(sheet, runtime.formula)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [modelVersion, runtime],
  );

  const selectedSheet = useMemo(
    () => sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0]!,
    [activeSheetId, sheets],
  );

  const selectCell = useCallback(
    (address: string) => {
      const parsed = parseAddress(address);
      if (!parsed) return;
      setActiveCell(address);
      const cell = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column);
      setFormulaDraft(cell?.formula ?? (cell?.value == null ? '' : String(cell.value)));
    },
    [activeSheetId, runtime],
  );

  const commitFormula = useCallback(
    (overrideValue?: string) => {
      if (phase !== 'ready') return;
      const parsed = parseAddress(activeCell);
      if (!parsed) return;

      const raw = (overrideValue !== undefined ? overrideValue : formulaDraft).trim();
      const isFormula = raw.startsWith('=');
      const address = { sheetId: activeSheetId, row: parsed.row, column: parsed.column };

      const result = isFormula
        ? runtime.formula.setFormula(address, raw)
        : runtime.formula.setValue(
            address,
            raw === '' ? null : Number.isFinite(Number(raw)) ? Number(raw) : raw,
          );

      const value = isFormula
        ? null
        : raw === ''
          ? null
          : Number.isFinite(Number(raw))
            ? Number(raw)
            : raw;

      const sheet = runtime.model.getSheet(activeSheetId);
      const existingStyle = sheet.cells.get(parsed.row, parsed.column)?.style;

      const commandResult = runtime.commands.execute('sheet.cell.set', {
        sheetId: activeSheetId,
        row: parsed.row,
        column: parsed.column,
        value: {
          value,
          formula: isFormula ? raw : undefined,
          displayValue: toFormulaDisplay(result.value),
          style: existingStyle,
        },
      });

      persistMutation(commandResult.operationId, [
        {
          id: 'cell.set',
          sheetId: activeSheetId,
          params: {
            row: parsed.row,
            column: parsed.column,
            value: {
              value,
              formula: isFormula ? raw : undefined,
              displayValue: toFormulaDisplay(result.value),
              style: existingStyle,
            },
          },
          affectedRanges: [
            {
              sheetId: activeSheetId,
              startRow: parsed.row,
              endRow: parsed.row,
              startColumn: parsed.column,
              endColumn: parsed.column,
            },
          ],
        },
      ]);

      refresh();
      setSaveState('saving');
      window.setTimeout(() => setSaveState('saved'), 200);
    },
    [activeCell, activeSheetId, formulaDraft, phase, persistMutation, refresh, runtime],
  );

  const moveCell = useCallback(
    (address: string, direction: 'down' | 'left' | 'right' | 'up') => {
      const parsed = parseAddress(address);
      if (!parsed) return;
      const offsets = { left: [-1, 0], right: [1, 0], up: [0, -1], down: [0, 1] } as const;
      const [columnOffset, rowOffset] = offsets[direction];
      const column = Math.max(0, Math.min(columns.length - 1, parsed.column + columnOffset));
      const row = Math.max(
        0,
        Math.min(runtime.model.getSheet(activeSheetId).rowCount - 1, parsed.row + rowOffset),
      );
      selectCell(cellAddress(row, column));
    },
    [activeSheetId, runtime, selectCell],
  );

  const addSheet = useCallback(() => {
    const id = `sheet-${runtime.model.getSheets().length + 1}`;
    const name = `Sheet ${runtime.model.getSheets().length + 1}`;
    const commandResult = runtime.commands.execute('sheet.add', { id, name });
    persistMutation(commandResult.operationId, [
      { id: 'sheet.add', sheetId: id, params: { id, name }, affectedRanges: [] },
    ]);
    runtime.model.activeSheetId = id;
    setActiveSheetId(id);
    setActiveCell('A1');
    setFormulaDraft('');
    refresh();
    setNotice('Worksheet added');
  }, [persistMutation, refresh, runtime]);

  const deleteSheet = useCallback(
    (sheetId: string) => {
      if (runtime.model.getSheets().length <= 1) return;
      runtime.commands.execute('sheet.remove', { sheetId });
      const first = runtime.model.getSheets()[0]!;
      setActiveSheetId(first.id);
      refresh();
    },
    [refresh, runtime],
  );

  const renameSheet = useCallback(
    (sheetId: string, name: string) => {
      runtime.commands.execute('sheet.rename', { sheetId, name });
      refresh();
    },
    [refresh, runtime],
  );

  const selectSheet = useCallback(
    (sheetId: SheetId) => {
      const sheet = runtime.model.getSheet(sheetId);
      runtime.model.activeSheetId = sheetId;
      setActiveSheetId(sheetId);
      setActiveCell('A1');
      setFormulaDraft(sheet.cells.get(0, 0)?.formula ?? '');
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

  const handleRibbonAction = useCallback(
    (action: string, payload?: unknown) => {
      const parsed = parseAddress(activeCell);
      if (!parsed) return;

      const activeRange: RangeRef = {
        sheetId: activeSheetId,
        startRow: parsed.row,
        endRow: parsed.row,
        startColumn: parsed.column,
        endColumn: parsed.column,
      };

      if (action === 'undo') undo();
      else if (action === 'redo') redo();
      else if (action === 'bold') {
        const currentStyle = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column)?.style;
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { bold: !currentStyle?.bold },
        });
        refresh();
      } else if (action === 'italic') {
        const currentStyle = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column)?.style;
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { italic: !currentStyle?.italic },
        });
        refresh();
      } else if (action === 'underline') {
        const currentStyle = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column)?.style;
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { underline: !currentStyle?.underline },
        });
        refresh();
      } else if (action === 'strikethrough') {
        const currentStyle = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column)?.style;
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { strikethrough: !currentStyle?.strikethrough },
        });
        refresh();
      } else if (action === 'align-left' || action === 'align-center' || action === 'align-right') {
        const align = action.replace('align-', '') as 'left' | 'center' | 'right';
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { horizontalAlignment: align },
        });
        refresh();
      } else if (action === 'wrap-text') {
        const currentStyle = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column)?.style;
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { wrapText: !currentStyle?.wrapText },
        });
        refresh();
      } else if (action === 'merge-cells') {
        runtime.commands.execute('sheet.merge.set', {
          sheetId: activeSheetId,
          range: {
            sheetId: activeSheetId,
            startRow: parsed.row,
            endRow: parsed.row,
            startColumn: parsed.column,
            endColumn: parsed.column + 1,
          },
        });
        refresh();
      } else if (action === 'textColor' && typeof payload === 'string') {
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { textColor: payload },
        });
        refresh();
      } else if (action === 'background' && typeof payload === 'string') {
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { background: payload },
        });
        refresh();
      } else if (action === 'numberFormat' && typeof payload === 'string') {
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { numberFormat: payload },
        });
        refresh();
      } else if (action === 'format-currency') {
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { numberFormat: '$#,##0' },
        });
        refresh();
      } else if (action === 'format-percent') {
        runtime.commands.execute('sheet.style.set', {
          sheetId: activeSheetId,
          range: activeRange,
          style: { numberFormat: '0%' },
        });
        refresh();
      } else if (action === 'clear-range') {
        runtime.commands.execute('sheet.range.clear', {
          sheetId: activeSheetId,
          range: activeRange,
        });
        runtime.formula.clearCell({ sheetId: activeSheetId, row: parsed.row, column: parsed.column });
        setFormulaDraft('');
        refresh();
      } else if (action === 'autosum') {
        const formula = `=SUM(A${parsed.row + 1}:${columnLabel(Math.max(0, parsed.column - 1))}${parsed.row + 1})`;
        setFormulaDraft(formula);
        commitFormula(formula);
      } else if (action === 'function-wizard') {
        setShowFunctionWizard(true);
      } else if (action === 'sort-dialog') {
        setShowSortDialog(true);
      } else if (action === 'sort-asc') {
        runtime.commands.execute('sheet.sort', {
          sheetId: activeSheetId,
          range: { sheetId: activeSheetId, startRow: 0, endRow: 20, startColumn: 0, endColumn: 5 },
          sortColumn: parsed.column,
          ascending: true,
          hasHeader: true,
        });
        refresh();
      } else if (action === 'sort-desc') {
        runtime.commands.execute('sheet.sort', {
          sheetId: activeSheetId,
          range: { sheetId: activeSheetId, startRow: 0, endRow: 20, startColumn: 0, endColumn: 5 },
          sortColumn: parsed.column,
          ascending: false,
          hasHeader: true,
        });
        refresh();
      } else if (action === 'copy') {
        const data = copyRangeToClipboardData(runtime.model, activeRange);
        if (typeof navigator !== 'undefined' && navigator.clipboard) {
          navigator.clipboard.writeText(formatTsv(data.values));
          setNotice('Copied to clipboard');
        }
      } else if (action === 'paste') {
        if (typeof navigator !== 'undefined' && navigator.clipboard) {
          navigator.clipboard.readText().then((text) => {
            const values = parseTsv(text);
            if (values.length > 0) {
              runtime.commands.execute('sheet.range.set', {
                sheetId: activeSheetId,
                startRow: parsed.row,
                startColumn: parsed.column,
                values,
              });
              refresh();
              setNotice('Pasted from clipboard');
            }
          });
        }
      } else if (action === 'open-chart') setActivePanel('chart');
      else if (action === 'open-pivot') setActivePanel('pivot');
      else if (action === 'open-shape') setActivePanel('shape');
      else if (action === 'open-sparkline') setActivePanel('sparkline');
      else if (action === 'open-conditional-format') setActivePanel('conditionalFormat');
      else if (action === 'open-data-validation') setActivePanel('dataValidation');
      else if (action === 'open-print') setActivePanel('print');
      else if (action === 'open-history') setActivePanel('history');
      else if (action === 'freeze-top-row') {
        runtime.commands.execute('sheet.freeze.set', {
          sheetId: activeSheetId,
          freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
        });
        refresh();
      } else if (action === 'freeze-first-col') {
        runtime.commands.execute('sheet.freeze.set', {
          sheetId: activeSheetId,
          freeze: { xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1 },
        });
        refresh();
      } else if (action === 'unfreeze') {
        runtime.commands.execute('sheet.freeze.set', {
          sheetId: activeSheetId,
          freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
        });
        refresh();
      } else if (action === 'zoom-in') setZoomState((z) => Math.min(125, z + 5));
      else if (action === 'zoom-out') setZoomState((z) => Math.max(75, z - 5));
      else if (action === 'zoom-100') setZoomState(100);
    },
    [activeCell, activeSheetId, commitFormula, redo, refresh, runtime, undo],
  );

  const addChart = useCallback(
    (chart: ChartModel) => {
      runtime.commands.execute('pro.chart.add', chart);
      refresh();
      setNotice('Chart added to sheet');
    },
    [refresh, runtime],
  );

  const removeChart = useCallback(
    (id: string) => {
      runtime.commands.execute('chart.remove', id);
      refresh();
    },
    [refresh, runtime],
  );

  const addPivot = useCallback(
    (pivot: PivotModel) => {
      runtime.commands.execute('pro.pivot.add', pivot);
      refresh();
      setNotice('Pivot table generated');
    },
    [refresh, runtime],
  );

  const removePivot = useCallback(
    (id: string) => {
      runtime.commands.execute('pivot.remove', id);
      refresh();
    },
    [refresh, runtime],
  );

  const addShape = useCallback(
    (shape: ShapeModel) => {
      runtime.commands.execute('pro.shape.add', shape);
      refresh();
      setNotice('Shape added to canvas');
    },
    [refresh, runtime],
  );

  const removeShape = useCallback(
    (id: string) => {
      runtime.commands.execute('shape.remove', id);
      refresh();
    },
    [refresh, runtime],
  );

  const addSparkline = useCallback(
    (sparkline: SparklineModel) => {
      runtime.commands.execute('pro.sparkline.add', sparkline);
      refresh();
      setNotice('Sparkline attached to cell');
    },
    [refresh, runtime],
  );

  const removeSparkline = useCallback(
    (id: string) => {
      runtime.commands.execute('sparkline.remove', id);
      refresh();
    },
    [refresh, runtime],
  );

  const addConditionalFormat = useCallback(
    (rule: ConditionalFormatRule) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      sheet.conditionalFormats.push(rule);
      refresh();
      setNotice('Conditional format rule applied');
    },
    [activeSheetId, refresh, runtime],
  );

  const removeConditionalFormat = useCallback(
    (id: string) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      const idx = sheet.conditionalFormats.findIndex((r) => r.id === id);
      if (idx >= 0) sheet.conditionalFormats.splice(idx, 1);
      refresh();
    },
    [activeSheetId, refresh, runtime],
  );

  const addDataValidation = useCallback(
    (rule: DataValidationRule) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      sheet.dataValidations.push(rule);
      refresh();
      setNotice('Data validation rule added');
    },
    [activeSheetId, refresh, runtime],
  );

  const removeDataValidation = useCallback(
    (id: string) => {
      const sheet = runtime.model.getSheet(activeSheetId);
      const idx = sheet.dataValidations.findIndex((r) => r.id === id);
      if (idx >= 0) sheet.dataValidations.splice(idx, 1);
      refresh();
    },
    [activeSheetId, refresh, runtime],
  );

  const printWorkbook = useCallback(
    (layout: PrintLayout) => {
      if (typeof window !== 'undefined') {
        window.print();
      }
    },
    [],
  );

  const exportPdf = useCallback(
    (layout: PrintLayout) => {
      if (typeof window !== 'undefined') {
        window.print();
      }
    },
    [],
  );

  const sortRange = useCallback(
    (colIdx: number, ascending: boolean, hasHeader: boolean) => {
      runtime.commands.execute('sheet.sort', {
        sheetId: activeSheetId,
        range: { sheetId: activeSheetId, startRow: 0, endRow: 20, startColumn: 0, endColumn: 10 },
        sortColumn: colIdx,
        ascending,
        hasHeader,
      });
      refresh();
    },
    [activeSheetId, refresh, runtime],
  );

  const historyEntries = useMemo(() => runtime.commands.getUndoEntries(), [modelVersion, runtime]);

  const actions = useMemo<WorkspaceActions>(
    () => ({
      addSheet,
      deleteSheet,
      renameSheet,
      commitFormula,
      moveCell,
      notify: setNotice,
      redo,
      retry,
      selectCell,
      selectSheet,
      setActivePanel,
      setFormulaDraft,
      setRibbonTab,
      setZoom: (nextZoom) => setZoomState(Math.max(75, Math.min(125, nextZoom))),
      undo,
      handleRibbonAction,
      addChart,
      removeChart,
      addPivot,
      removePivot,
      addShape,
      removeShape,
      addSparkline,
      removeSparkline,
      addConditionalFormat,
      removeConditionalFormat,
      addDataValidation,
      removeDataValidation,
      printWorkbook,
      exportPdf,
      closeFunctionWizard: () => setShowFunctionWizard(false),
      closeSortDialog: () => setShowSortDialog(false),
      sortRange,
    }),
    [
      addChart,
      addConditionalFormat,
      addDataValidation,
      addPivot,
      addShape,
      addSheet,
      addSparkline,
      commitFormula,
      deleteSheet,
      exportPdf,
      handleRibbonAction,
      moveCell,
      printWorkbook,
      redo,
      removeChart,
      removeConditionalFormat,
      removeDataValidation,
      removePivot,
      removeShape,
      removeSparkline,
      renameSheet,
      retry,
      selectCell,
      selectSheet,
      sortRange,
      undo,
    ],
  );

  const state = useMemo<WorkspaceState>(
    () => ({
      activeCell,
      activePanel,
      activeSheetId,
      formulaDraft,
      notice,
      phase,
      ribbonTab,
      saveState,
      selectedSheet,
      sheets,
      zoom,
      historyEntries,
      showFunctionWizard,
      showSortDialog,
    }),
    [
      activeCell,
      activePanel,
      activeSheetId,
      formulaDraft,
      historyEntries,
      notice,
      phase,
      ribbonTab,
      saveState,
      selectedSheet,
      sheets,
      showFunctionWizard,
      showSortDialog,
      zoom,
    ],
  );

  return { actions, state };
}
