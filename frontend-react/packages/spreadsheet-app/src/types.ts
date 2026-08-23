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
  | 'automate';
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
export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export interface PeerCursor {
  actorId: string;
  name: string;
  color: string;
  sheetId: string;
  row: number;
  column: number;
}

export type { SelectionState, SelectionSnapshot } from './selection-service';
