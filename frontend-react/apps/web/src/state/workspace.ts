import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { WorkbookModel, type CellData, type SheetId, type WorksheetModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { FormulaEngine, isFormulaError, type FormulaValue } from '@react-sheets/formula-engine';
import { WorkbookApiClient, type CollaborationMutation } from '@react-sheets/protocol';

export type WorkspacePhase = 'empty' | 'error' | 'loading' | 'ready';
export type RibbonTabId = 'data' | 'home' | 'insert' | 'review' | 'view';
export type SidebarPanelId = 'comments' | 'data' | 'inspector' | 'automations';
export type SaveState = 'saved' | 'saving' | 'offline';
export type CellTone = 'accent' | 'header' | 'muted' | 'plain' | 'total';

export interface SheetCell { address: string; displayValue?: string; formula?: string; tone?: CellTone; value: string; }
export interface SheetRow { cells: SheetCell[]; rowNumber: number; }
export interface SheetView { columns: string[]; id: string; isEmpty?: boolean; name: string; rows: SheetRow[]; }

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
}

export interface WorkspaceActions {
  addSheet: () => void;
  commitFormula: () => void;
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
}

export interface UseWorkspaceStateOptions { initialPhase?: WorkspacePhase; }

interface WorkspaceRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  model: WorkbookModel;
  commands: CommandRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
}

const columns = Array.from({ length: 12 }, (_, index) => columnLabel(index));

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

function parseAddress(address: string): { column: number; row: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  return { column: column - 1, row: Number(match[2]) - 1 };
}

function cellAddress(row: number, column: number): string { return `${columnLabel(column)}${row + 1}`; }

function toFormulaDisplay(value: FormulaValue): string {
  if (isFormulaError(value)) return value.code;
  return value == null ? '' : String(value);
}

function seedWorkbook(runtime: WorkspaceRuntime): void {
  const values: CellData[][] = [
    [{ value: 'Q3 growth plan' }, { value: 'Owner' }, { value: 'Status' }, { value: 'Target' }, { value: 'Actual' }, { value: 'Variance' }],
    [{ value: 'Activation' }, { value: 'Maya Chen' }, { value: 'On track' }, { value: 0.42, numberFormat: '0%' }, { value: 0.38, numberFormat: '0%' }, { value: -0.04, numberFormat: '0%' }],
    [{ value: 'Retention' }, { value: 'Noah Williams' }, { value: 'Needs review' }, { value: 0.68, numberFormat: '0%' }, { value: 0.64, numberFormat: '0%' }, { value: -0.04, numberFormat: '0%' }],
    [{ value: 'Expansion' }, { value: 'Ava Patel' }, { value: 'On track' }, { value: 120000, numberFormat: '$#,##0' }, { value: 132000, numberFormat: '$#,##0' }, { value: 0.1, numberFormat: '0%' }],
    [{ value: 'Referrals' }, { value: 'Liam Garcia' }, { value: 'At risk' }, { value: 0.16, numberFormat: '0%' }, { value: 0.11, numberFormat: '0%' }, { value: -0.05, numberFormat: '0%' }],
    [{ value: 'Blended view' }, { value: 'Team total' }, { value: 'Review' }, { value: null, formula: '=SUM(D2:D5)' }, { value: null, formula: '=SUM(E2:E5)' }, { value: null, formula: '=D6-E6' }],
  ];
  runtime.commands.execute('sheet.range.set', { sheetId: 'sheet-1', startRow: 0, startColumn: 0, values });
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
  const model = new WorkbookModel(crypto.randomUUID(), 'Q3 Growth Planning');
  const commands = new CommandRuntime(model);
  registerSheetCommands(commands);
  const runtime = { api: new WorkbookApiClient(), model, commands, formula: new FormulaEngine({ defaultSheetId: 'sheet-1' }), remoteConnected: false, remoteRevision: 0 };
  seedWorkbook(runtime);
  return runtime;
}

function formatCellValue(cell: CellData | undefined, formula: FormulaEngine, sheetId: string, row: number, column: number): string {
  if (!cell) return '';
  if (cell.formula) return toFormulaDisplay(formula.getCellValue({ sheetId, row, column }));
  if (cell.value == null) return '';
  if (typeof cell.value === 'number' && cell.numberFormat === '0%') return `${Math.round(cell.value * 100)}%`;
  if (typeof cell.value === 'number' && cell.numberFormat === '$#,##0') return `$${cell.value.toLocaleString('en-US')}`;
  return String(cell.value);
}

function toSheetView(sheet: WorksheetModel, formula: FormulaEngine): SheetView {
  const rows: SheetRow[] = [];
  for (let row = 0; row < Math.max(24, sheet.rowCount); row += 1) {
    const cells: SheetCell[] = [];
    for (let column = 0; column < columns.length; column += 1) {
      const modelCell = sheet.cells.get(row, column);
      const value = formatCellValue(modelCell, formula, sheet.id, row, column);
      const tone: CellTone | undefined = row === 0 ? 'header' : modelCell?.value === 'On track' ? 'accent' : modelCell?.value === 'At risk' || modelCell?.value === 'Needs review' ? 'total' : undefined;
      cells.push({ address: cellAddress(row, column), formula: modelCell?.formula, tone, value });
    }
    rows.push({ rowNumber: row + 1, cells });
  }
  return { id: sheet.id, name: sheet.name, columns, rows };
}

export function getInitialWorkspacePhase(): WorkspacePhase {
  if (typeof window === 'undefined') return 'ready';
  const queryPhase = new URLSearchParams(window.location.search).get('state');
  return queryPhase === 'loading' || queryPhase === 'error' || queryPhase === 'empty' ? queryPhase : 'ready';
}

export function useWorkspaceState({ initialPhase = 'ready' }: UseWorkspaceStateOptions = {}): { actions: WorkspaceActions; state: WorkspaceState } {
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
  const [notice, setNotice] = useState('Workbook ready');
  const [modelVersion, setModelVersion] = useState(0);
  const refresh = useCallback(() => setModelVersion((version) => version + 1), []);
  useEffect(() => {
    let active = true;
    void runtime.api.createWorkbook(runtime.model.snapshot()).then((response) => {
      if (!active) return;
      runtime.remoteConnected = true;
      runtime.remoteRevision = response.revision;
      setNotice('Server snapshot connected');
    }).catch(() => {
      if (!active) return;
      setSaveState('offline');
      setNotice('Server unavailable; local WorkbookModel remains active');
    });
    return () => { active = false; };
  }, [runtime]);
  const persistMutation = useCallback((operationId: string, mutations: CollaborationMutation[]) => {
    if (!runtime.remoteConnected) {
      setSaveState('offline');
      setNotice('Mutation is local; server is not connected');
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
    void runtime.api.submitChangeSet(changeSet).then((result) => {
      runtime.remoteRevision = result.revision;
      setSaveState('saved');
      setNotice(`Server revision ${result.revision}`);
    }).catch(() => {
      setSaveState('offline');
      setNotice('Server rejected Mutation; local state is retained');
    });
  }, [runtime]);
  const sheets = useMemo(() => runtime.model.getSheets().map((sheet) => toSheetView(sheet, runtime.formula)), [modelVersion, runtime]);
  const selectedSheet = useMemo(() => sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0]!, [activeSheetId, sheets]);

  const selectCell = useCallback((address: string) => {
    const parsed = parseAddress(address);
    if (!parsed) return;
    setActiveCell(address);
    const cell = runtime.model.getSheet(activeSheetId).cells.get(parsed.row, parsed.column);
    setFormulaDraft(cell?.formula ?? (cell?.value == null ? '' : String(cell.value)));
  }, [activeSheetId, runtime]);

  const commitFormula = useCallback(() => {
    if (phase !== 'ready') return;
    const parsed = parseAddress(activeCell);
    if (!parsed) return;
    const raw = formulaDraft.trim();
    const isFormula = raw.startsWith('=');
    const result = isFormula
      ? runtime.formula.setFormula({ sheetId: activeSheetId, row: parsed.row, column: parsed.column }, raw)
      : runtime.formula.setValue({ sheetId: activeSheetId, row: parsed.row, column: parsed.column }, raw === '' ? null : Number.isFinite(Number(raw)) ? Number(raw) : raw);
    const value = isFormula ? null : raw === '' ? null : Number.isFinite(Number(raw)) ? Number(raw) : raw;
    const commandResult = runtime.commands.execute('sheet.cell.set', { sheetId: activeSheetId, row: parsed.row, column: parsed.column, value: { value, formula: isFormula ? raw : undefined, displayValue: toFormulaDisplay(result.value) } });
    persistMutation(commandResult.operationId, [{ id: 'cell.set', sheetId: activeSheetId, params: { row: parsed.row, column: parsed.column, value: { value, formula: isFormula ? raw : undefined, displayValue: toFormulaDisplay(result.value) } }, affectedRanges: [{ sheetId: activeSheetId, startRow: parsed.row, endRow: parsed.row, startColumn: parsed.column, endColumn: parsed.column }] }]);
    refresh();
    setSaveState('saving');
    setNotice('Mutation applied to WorkbookModel');
    window.setTimeout(() => setSaveState('saved'), 260);
  }, [activeCell, activeSheetId, formulaDraft, phase, persistMutation, refresh, runtime]);

  const moveCell = useCallback((address: string, direction: 'down' | 'left' | 'right' | 'up') => {
    const parsed = parseAddress(address);
    if (!parsed) return;
    const offsets = { left: [-1, 0], right: [1, 0], up: [0, -1], down: [0, 1] } as const;
    const [columnOffset, rowOffset] = offsets[direction];
    const column = Math.max(0, Math.min(columns.length - 1, parsed.column + columnOffset));
    const row = Math.max(0, Math.min(runtime.model.getSheet(activeSheetId).rowCount - 1, parsed.row + rowOffset));
    selectCell(cellAddress(row, column));
  }, [activeSheetId, runtime, selectCell]);

  const addSheet = useCallback(() => {
    const id = `sheet-${runtime.model.getSheets().length + 1}`;
    const commandResult = runtime.commands.execute('sheet.add', { id, name: `Sheet ${runtime.model.getSheets().length + 1}` });
    persistMutation(commandResult.operationId, [{ id: 'sheet.add', sheetId: id, params: { id, name: `Sheet ${runtime.model.getSheets().length + 1}` }, affectedRanges: [] }]);
    runtime.model.activeSheetId = id;
    setActiveSheetId(id);
    setActiveCell('A1');
    setFormulaDraft('');
    refresh();
    setNotice('Worksheet Mutation committed');
  }, [persistMutation, refresh, runtime]);

  const selectSheet = useCallback((sheetId: SheetId) => {
    const sheet = runtime.model.getSheet(sheetId);
    runtime.model.activeSheetId = sheetId;
    setActiveSheetId(sheetId);
    setActiveCell('A1');
    setFormulaDraft(sheet.cells.get(0, 0)?.formula ?? '');
    refresh();
    setNotice(`Viewing ${sheet.name}`);
  }, [refresh, runtime]);

  const undo = useCallback(() => {
    const applied = runtime.commands.undo();
    if (applied) {
      refresh();
      setNotice('Undo Mutation applied');
    } else setNotice('Undo stack is empty');
  }, [refresh, runtime]);
  const redo = useCallback(() => {
    const applied = runtime.commands.redo();
    if (applied) {
      refresh();
      setNotice('Redo Mutation applied');
    } else setNotice('Redo stack is empty');
  }, [refresh, runtime]);
  const retry = useCallback(() => { setPhase('ready'); setNotice('Workspace is ready'); }, []);

  const actions = useMemo<WorkspaceActions>(() => ({
    addSheet, commitFormula, moveCell, notify: setNotice, redo, retry, selectCell, selectSheet,
    setActivePanel, setFormulaDraft, setRibbonTab,
    setZoom: (nextZoom) => setZoomState(Math.max(75, Math.min(125, nextZoom))), undo,
  }), [addSheet, commitFormula, moveCell, redo, retry, selectCell, selectSheet, undo]);
  const state = useMemo<WorkspaceState>(() => ({ activeCell, activePanel, activeSheetId, formulaDraft, notice, phase, ribbonTab, saveState, selectedSheet, sheets, zoom }), [activeCell, activePanel, activeSheetId, formulaDraft, notice, phase, ribbonTab, saveState, selectedSheet, sheets, zoom]);
  return { actions, state };
}
