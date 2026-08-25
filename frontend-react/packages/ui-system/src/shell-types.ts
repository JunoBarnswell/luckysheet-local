import type { ReactNode } from 'react';

export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export const DESIGNER_GEOMETRY = {
  viewportWidth: 1280,
  viewportHeight: 720,
  ribbonHeight: 142,
  formulaBarHeight: 37,
  workspaceHeight: 519,
  sheetTabsHeight: 29,
  statusBarHeight: 22,
  ribbonTabHeight: 36,
  ribbonContentHeight: 106,
} as const;

export interface PeerCursor {
  actorId: string;
  name: string;
  color: string;
}

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
  | 'tableSheetDesign';

export const RIBBON_TAB_ORDER: readonly RibbonTabId[] = [
  'file',
  'home',
  'insert',
  'pageLayout',
  'formulas',
  'data',
  'review',
  'view',
  'settings',
  'automate',
] as const;

export type RibbonLayoutMode = 'wide' | 'compact' | 'narrow';

export interface RibbonLayoutState {
  mode: RibbonLayoutMode;
  width: number;
}

export interface DesignerShellProps {
  children: ReactNode;
  formulaBar: ReactNode;
  isBusy: boolean;
  ribbon: ReactNode;
  sheetTabs: ReactNode;
  statusBar: ReactNode;
  floatingOverlay?: ReactNode;
  workspacePhase: string;
}
