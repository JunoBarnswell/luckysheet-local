import type { ReactNode } from 'react';

export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

export const DESIGNER_GEOMETRY = {
  viewportWidth: 1920,
  viewportHeight: 1080,
  ribbonHeight: 167,
  formulaBarHeight: 48,
  workspaceHeight: 843,
  sheetTabsHeight: 29,
  statusBarHeight: 22,
  ribbonTabHeight: 32,
  ribbonContentHeight: 135,
} as const;

/**
 * The Designer Ribbon geometry is a fixed product contract. Hosts and tab
 * renderers consume these values through the shared geometry above; viewport
 * width changes only the scroll position, never the command density.
 */
export const RIBBON_DENSITY = {
  shellHeight: 167,
  tabStripHeight: 32,
  commandAreaHeight: 135,
  groupContentHeight: 127,
  largeCommandHeight: 104,
  groupCaptionHeight: 18,
} as const;

export const RIBBON_DENSITY_CLASSES = {
  shell: 'h-[clamp(118px,15.5vh,167px)]',
  tabStrip: 'h-[clamp(28px,3vh,32px)]',
  commandArea: 'h-[clamp(90px,12.5vh,135px)]',
  groupContent: 'h-[clamp(84px,11.8vh,127px)]',
  groupControls: 'h-[clamp(66px,9.6vh,104px)]',
  largeCommand: '!h-[clamp(66px,9.6vh,104px)]',
  groupCaption: 'h-[18px] leading-[18px]',
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
