import type { CellStyle, RangeRef } from '@react-sheets/core-model';
import type { SelectionState } from './selection-service';

export type AppPhase = 'empty' | 'error' | 'loading' | 'ready';
export type RibbonTabId =
  | 'file'
  | 'home'
  | 'insert'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'settings'
  | 'automate'
  | 'pivotAnalyze'
  | 'pivotDesign'
  | 'tableSheetDesign'
  | 'ganttTask'
  | 'ganttProject'
  | 'ganttView'
  | 'ganttFormat'
  | 'reportSheetDesign'
  | 'tableDesign'
  | 'chartDesign'
  | 'chartFormat';
export type SidebarPanelId =
  | 'inspector'
  | 'chart'
  | 'dataChart'
  | 'barcode'
  | 'pivot'
  | 'formulaAudit'
  | 'definedNames'
  | 'shape'
  | 'sparkline'
  | 'conditionalFormat'
  | 'selectionPane'
  | 'dataValidation'
  | 'print'
  | 'query'
  | 'automate'
  | 'extended'
  | 'history'
  | 'data';
export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export type InputMode =
  | 'grid'
  | 'cell-edit'
  | 'formula-edit'
  | 'ribbon'
  | 'ribbon-keytip'
  | 'command-palette'
  | 'context-menu'
  | 'dropdown'
  | 'dialog'
  | 'side-panel';

export type FocusTarget = 'grid' | 'ribbon' | 'formula-bar' | 'command-palette' | 'context-menu' | 'dialog' | 'side-panel';

export interface FocusState {
  mode: InputMode;
  target: FocusTarget;
}

export interface EditSession {
  sheetId: string;
  cell: { row: number; column: number };
  originalValue: unknown;
  originalFormula?: string;
  draftText: string;
  mode: 'value' | 'formula';
  source: 'cell' | 'formulaBar';
}

export interface PanelState {
  active: SidebarPanelId;
  open: boolean;
  width?: number;
  dock: 'left' | 'right';
}

export type DialogId = 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells' | 'create-pivot' | 'create-table' | 'merge-confirm' | 'column-width' | 'command-palette' | 'sheet-dialog' | 'cell-template' | 'cell-editor' | 'insert-picture';
export type CellShiftOperation = 'insert' | 'delete';

export type SheetDialogKind = 'rename' | 'tab-color' | 'delete';

export interface SheetDialogState {
  kind: SheetDialogKind;
  sheetId: string;
  value: string;
}

export interface DialogState {
  active: DialogId | null;
  findQuery: string;
  mergeDiscardCount: number;
  columnWidth: { columns: number[]; defaultMode: boolean } | null;
  sheet: SheetDialogState | null;
  cellShiftOperation: CellShiftOperation;
}

export interface ClipboardState {
  hasContent: boolean;
  mode: 'copy' | 'cut' | null;
  systemStatus: 'unknown' | 'published' | 'reduced' | 'failed';
  systemFormats: readonly string[];
}

export interface UndoRedoState {
  canUndo: boolean;
  canRedo: boolean;
  undoCount: number;
  redoCount: number;
}

export type BackstagePanel = 'info' | 'options';

export interface BackstageState {
  open: boolean;
  panel: BackstagePanel;
}

export interface DesignerState {
  workbook: { unitId: string; name: string };
  selection: SelectionState;
  editSession: EditSession | null;
  activeObject: { kind: string; id: string } | null;
  ribbon: { activeTab: RibbonTabId };
  panels: PanelState;
  dialogs: DialogState;
  clipboard: ClipboardState;
  focus: FocusState;
  inputMode: InputMode;
  undoRedo: UndoRedoState;
  backstage: BackstageState;
}

/** Ephemeral selection context for contextual Ribbon tabs and menus. */
export type ActiveContext =
  | { kind: 'none' }
  | { kind: 'pivot'; sheetId: string; pivotId: string }
  | { kind: 'drawing'; sheetId: string; drawingId: string }
  | { kind: 'table'; sheetId: string; tableId: string }
  | { kind: 'table-sheet'; sheetId: string; viewId: string }
  | { kind: 'gantt-sheet'; sheetId: string; viewId: string }
  | { kind: 'report-sheet'; sheetId: string; tableId?: string };

/** Ephemeral chrome state; these intents never write the workbook model. */
export type UiSessionIntent =
  | { type: 'panel.open'; panel: SidebarPanelId; notice?: string }
  | { type: 'dialog.open'; dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells' | 'create-pivot' | 'create-table' | 'column-width' | 'sheet-rename' | 'sheet-tab-color' | 'sheet-delete' | 'cell-template' | 'cell-editor' | 'insert-picture'; operation?: CellShiftOperation; findQuery?: string; columnWidth?: { columns: number[]; defaultMode: boolean }; sheet?: SheetDialogState }
  | { type: 'dialog.close' }
  | { type: 'dialog.update'; value: string }
  | { type: 'command-palette.open' }
  | { type: 'command-palette.close' }
  | { type: 'backstage.open'; panel: 'info' | 'options' }
  | { type: 'zoom.set'; value: number }
  | { type: 'zoom.adjust'; delta?: number; value?: number }
  | { type: 'notice'; message: string };

export interface PeerCursor {
  actorId: string;
  name: string;
  color: string;
  sheetId: string;
  row: number;
  column: number;
}

/** A derived Home-ribbon value. It is never persisted as workbook state. */
export type HomeSelectionValue<T> =
  | { kind: 'value'; value: T }
  | { kind: 'mixed' }
  | { kind: 'unset' };

export type HomeStyleKey = keyof Pick<
  CellStyle,
  | 'fontFamily'
  | 'textRotate'
  | 'fontSizePx'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'textColor'
  | 'background'
  | 'horizontalAlignment'
  | 'verticalAlignment'
  | 'indent'
  | 'wrapText'
  | 'numberFormat'
  | 'borders'
>;

/**
 * Read-only selection-derived state consumed by every Home entry point.
 * The WorkbookSession recalculates it from the canonical model on snapshot
 * creation; UI components must never keep a competing copy.
 */
export interface HomeRibbonState {
  sheetId: string;
  ranges: readonly RangeRef[];
  activeCell: { row: number; column: number };
  style: Partial<CellStyle>;
  mixedStyleKeys: readonly HomeStyleKey[];
  merge: 'none' | 'full' | 'mixed';
  canFormat: boolean;
  canEdit: boolean;
  canStructure: boolean;
  hasFilter: boolean;
  hasFilterCriteria: boolean;
}

export type { SelectionState, SelectionSnapshot } from './selection-service';
