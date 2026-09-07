import type { ReactNode } from 'react';

export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export const DESIGNER_GEOMETRY = {
  viewportWidth: 1920,
  viewportHeight: 1080,
  ribbonHeight: 104,
  formulaBarHeight: 32,
  workspaceHeight: 843,
  sheetTabsHeight: 28,
  statusBarHeight: 22,
  ribbonTabHeight: 28,
  ribbonContentHeight: 76,
  documentBarHeight: 36,
} as const;

/**
 * The Designer Ribbon geometry is a fixed product contract. Hosts and tab
 * renderers consume these values through the shared geometry above; viewport
 * width changes only the scroll position, never the command density.
 */
export const RIBBON_DENSITY = {
  shellHeight: 104,
  tabStripHeight: 28,
  commandAreaHeight: 76,
  groupContentHeight: 70,
  largeCommandHeight: 58,
  groupCaptionHeight: 16,
} as const;

export const RIBBON_DENSITY_CLASSES = {
  shell: 'h-[104px]',
  tabStrip: 'h-[28px]',
  commandArea: 'h-[76px]',
  groupContent: 'h-[70px]',
  groupControls: 'h-[58px]',
  largeCommand: '!h-[58px]',
  groupCaption: 'h-[16px] leading-[16px]',
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

export interface RibbonKeyTipState {
  active: boolean;
  prefix: string;
}

export interface RibbonKeyTipBinding {
  sequence: string;
  target: { kind: 'tab' | 'command'; id: string };
}

export interface DesignerShellProps {
  children: ReactNode;
  documentBar?: ReactNode;
  formulaBar: ReactNode;
  formulaBarVisible?: boolean;
  ribbonVisible?: boolean;
  isBusy: boolean;
  ribbon: ReactNode;
  sheetTabs: ReactNode;
  statusBar: ReactNode;
  floatingOverlay?: ReactNode;
  workspacePhase: string;
}
