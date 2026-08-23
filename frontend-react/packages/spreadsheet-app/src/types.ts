import type { CellStyle, RangeRef } from '@react-sheets/core-model';

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
  | 'automate'
  | 'pivotAnalyze'
  | 'pivotDesign';
export type SidebarPanelId =
  | 'inspector'
  | 'chart'
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

/** Ephemeral selection context for contextual Ribbon tabs and menus. */
export type ActiveContext =
  | { kind: 'none' }
  | { kind: 'pivot'; sheetId: string; pivotId: string }
  | { kind: 'drawing'; sheetId: string; drawingId: string }
  | { kind: 'table'; sheetId: string; tableId: string };

/** Ephemeral chrome state; these intents never write the workbook model. */
export type UiSessionIntent =
  | { type: 'panel.open'; panel: SidebarPanelId; notice?: string }
  | { type: 'dialog.open'; dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells' | 'create-pivot'; findQuery?: string }
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
  | 'fontSizePx'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'textColor'
  | 'background'
  | 'horizontalAlignment'
  | 'verticalAlignment'
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
