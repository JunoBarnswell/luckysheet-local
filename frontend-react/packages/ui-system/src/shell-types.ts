import type { ReactNode } from 'react';

export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export const DESIGNER_GEOMETRY = {
  viewportWidth: 1920,
  viewportHeight: 1080,
  ribbonHeight: 118,
  formulaBarHeight: 48,
  workspaceHeight: 892,
  sheetTabsHeight: 29,
  statusBarHeight: 22,
  ribbonTabHeight: 32,
  ribbonContentHeight: 86,
} as const;

/**
 * The compact Designer Ribbon is a product contract, not a per-tab styling
 * preference.  Hosts and tab renderers consume these values through the
 * shared geometry above; the named values make the control-height rules
 * explicit for visual tests and future surfaces.
 */
export const RIBBON_DENSITY = {
  shellHeight: 118,
  tabStripHeight: 32,
  commandAreaHeight: 86,
  groupControlHeight: 56,
  groupCaptionHeight: 14,
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

export const RIBBON_TAB_ORDER: readonly RibbonTabId[] = [
  'file',
  'home',
  'insert',
  'pageLayout',
  'formulas',
  'data',
  'view',
  'review',
  'settings',
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
