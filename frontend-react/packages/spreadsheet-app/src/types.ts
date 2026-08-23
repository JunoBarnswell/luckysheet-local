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

export type { SelectionState, SelectionSnapshot } from './selection-service';
