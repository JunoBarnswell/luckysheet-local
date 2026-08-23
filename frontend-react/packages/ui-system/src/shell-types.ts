import type { ReactNode } from 'react';

export type SaveState = 'saved' | 'saving' | 'offline' | 'syncing' | 'conflict' | 'calculating' | 'error';

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
  | 'automate';

export const RIBBON_TAB_ORDER: readonly RibbonTabId[] = [
  'file',
  'home',
  'insert',
  'pageLayout',
  'formulas',
  'data',
  'review',
  'view',
  'automate',
] as const;

export type RibbonLayoutMode = 'wide' | 'compact' | 'narrow';

export interface RibbonLayoutState {
  mode: RibbonLayoutMode;
  width: number;
}

export interface AppShellProps {
  children: ReactNode;
  formulaBar: ReactNode;
  isBusy: boolean;
  notice: string;
  onSearch?: (query: string) => void;
  onShare: () => void;
  peers: readonly PeerCursor[];
  workbookMenu?: ReactNode;
  ribbon: ReactNode;
  saveState: SaveState;
  sheetTabs: ReactNode;
  statusBar: ReactNode;
  title: string;
  workspacePhase: string;
  /** Localized shell strings */
  labels: {
    workspaceLabel: string;
    planningWorkbook: string;
    personal: string;
    searchWorkbook: string;
    share: string;
    language: string;
    english: string;
    simplifiedChinese: string;
    saveState: string;
  };
  localeMenuLabel: string;
  onLocaleChange: (locale: 'zh-CN' | 'en-US') => void;
}
