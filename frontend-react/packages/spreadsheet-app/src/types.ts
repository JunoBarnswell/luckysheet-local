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
  | 'chartFormat'
  | 'pictureFormat'
  | 'shapeFormat'
  | 'sparklineDesign';
export type SidebarPanelId =
  | 'inspector'
  | 'chart'
  | 'barcode'
  | 'pivot'
  | 'slicer'
  | 'formulaAudit'
  | 'definedNames'
  | 'shape'
  | 'textbox'
  | 'formControl'
  | 'picture'
  | 'sparkline'
  | 'conditionalFormat'
  | 'selectionPane'
  | 'dataValidation'
  | 'print'
  | 'query'
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

export interface PanelState {
  active: SidebarPanelId;
  open: boolean;
  width?: number;
  dock: 'left' | 'right';
}

export type LocalObjectDialogKind = 'icon' | 'model3d' | 'smartart' | 'screenshot' | 'wordart' | 'signature-line' | 'embedded-object' | 'equation';
export type DialogId = 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'phonetic-guide' | 'symbol' | 'shift-cells' | 'create-pivot' | 'create-table' | 'recommended-pivots' | 'recommended-charts' | 'merge-confirm' | 'column-width' | 'row-height' | 'command-palette' | 'sheet-dialog' | 'cell-template' | 'cell-editor' | 'insert-picture' | 'hyperlink' | 'local-object';
export type FindDialogMode = 'find' | 'replace';
export type CellShiftOperation = 'insert' | 'delete';
export type MergeOperation = 'center' | 'cells' | 'across' | 'unmerge';
export type FormatCellsTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

export type SheetDialogKind = 'rename' | 'tab-color' | 'delete';

export interface SheetDialogState {
  kind: SheetDialogKind;
  sheetId: string;
  value: string;
}

export interface DialogState {
  active: DialogId | null;
  findQuery: string;
  findMode: FindDialogMode;
  mergeDiscardCount: number;
  mergeOperation: MergeOperation;
  columnWidth: { columns: number[]; defaultMode: boolean } | null;
  rowHeight: { rows: number[] } | null;
  sheet: SheetDialogState | null;
  cellShiftOperation: CellShiftOperation;
  formatCellsTab: FormatCellsTab;
  localObjectKind: LocalObjectDialogKind | null;
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
  | { kind: 'sparkline'; sheetId: string; sparklineId: string }
  | { kind: 'table'; sheetId: string; tableId: string }
  | { kind: 'table-sheet'; sheetId: string; viewId: string }
  | { kind: 'gantt-sheet'; sheetId: string; viewId: string }
  | { kind: 'report-sheet'; sheetId: string; tableId?: string };

/** Ephemeral chrome state; these intents never write the workbook model. */
export type UiSessionIntent =
  | { type: 'panel.open'; panel: SidebarPanelId; notice?: string }
  | { type: 'dialog.open'; dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'phonetic-guide' | 'symbol' | 'shift-cells' | 'create-pivot' | 'create-table' | 'recommended-pivots' | 'recommended-charts' | 'column-width' | 'row-height' | 'sheet-rename' | 'sheet-tab-color' | 'sheet-delete' | 'cell-template' | 'cell-editor' | 'insert-picture' | 'hyperlink' | 'local-object'; localObjectKind?: LocalObjectDialogKind; operation?: CellShiftOperation; findQuery?: string; findMode?: FindDialogMode; formatCellsTab?: FormatCellsTab; columnWidth?: { columns: number[]; defaultMode: boolean }; rowHeight?: { rows: number[] }; sheet?: SheetDialogState }
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
  | 'textOrientation'
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
  | 'shrinkToFit'
  | 'numberFormat'
  | 'borders'
  | 'locked'
  | 'formulaHidden'
>;

export type HomeStyleFieldState =
  | { status: 'uniform'; value: unknown }
  | { status: 'mixed' }
  | { status: 'unset' }
  | { status: 'unsupported'; reason: string };

export type HomeStyleAggregate = { [K in HomeStyleKey]: HomeStyleFieldState };

/**
 * Read-only selection-derived state consumed by every Home entry point.
 * The WorkbookSession recalculates it from the canonical model on snapshot
 * creation; UI components must never keep a competing copy.
 */
export interface HomeRibbonState {
  sheetId: string;
  ranges: readonly RangeRef[];
  activeCell: { row: number; column: number };
  styleAggregate: HomeStyleAggregate;
  style: Partial<CellStyle>;
  mixedStyleKeys: readonly HomeStyleKey[];
  unsupportedStyleKeys: readonly HomeStyleKey[];
  merge: 'none' | 'full' | 'mixed';
  canFormat: boolean;
  canEdit: boolean;
  canStructure: boolean;
  hasFilter: boolean;
  hasFilterCriteria: boolean;
}

export type { SelectionState, SelectionSnapshot } from './selection-service';
