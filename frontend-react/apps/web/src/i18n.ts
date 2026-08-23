import enUS from './locales/en-US.json';
import zhCN from './locales/zh-CN.json';
import type { RibbonTabId, SaveState } from '@react-sheets/ui-system';
import type { RibbonTextKey } from '@react-sheets/spreadsheet-app';

export type Locale = 'zh-CN' | 'en-US';

export type MessageKey =
  | 'workspaceLabel'
  | 'planningWorkbook'
  | 'personal'
  | 'saved'
  | 'saving'
  | 'offline'
  | 'syncing'
  | 'conflict'
  | 'calculating'
  | 'error'
  | 'share'
  | 'language'
  | 'english'
  | 'simplifiedChinese'
  | 'file'
  | 'home'
  | 'insert'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'automate'
  | 'engineConnected'
  | 'searchWorkbook';

const messages: Record<Locale, Record<MessageKey, string>> = {
  'en-US': enUS as Record<MessageKey, string>,
  'zh-CN': zhCN as Record<MessageKey, string>,
};

const localeBundles: Record<Locale, typeof enUS> = {
  'en-US': enUS,
  'zh-CN': zhCN,
};

const textTranslations: Record<string, string> = {
  'Inspect': '检查',
  'Chart': '图表',
  'Pivot': '透视',
  'Shape': '形状',
  'Spark': '迷你图',
  'Format': '格式',
  'Validate': '验证',
  'Print': '打印',
  'Tables': '数据表',
  'History': '历史',
  'Clipboard': '剪贴板',
  'Font': '字体',
  'Alignment': '对齐',
  'Number': '数字',
  'Cells': '单元格',
  'Editing': '编辑',
  'Tables & Pivots': '表格与透视',
  'Charts & Visuals': '图表与可视化',
  'Illustrations': '插图',
  'Functions': '函数',
  'Sort & Filter': '排序与筛选',
  'Data Tools': '数据工具',
  'Find & Transform': '查找与转换',
  'History & Audit': '历史与审计',
  'Freeze Panes': '冻结窗格',
  'Zoom': '缩放',
  'Print Layout': '打印布局',
  'Appearance & Files': '外观与文件',
  'Text Color': '文字颜色',
  'Fill Background': '填充背景',
  'Clear': '清除',
};

export const localeLabels: Record<Locale, string> = {
  'en-US': 'EN / 中文',
  'zh-CN': '中文 / EN',
};

const LOCALE_STORAGE_KEY = 'react-sheets:locale';

export function getInitialLocale(): Locale {
  if (typeof window === 'undefined') return 'en-US';
  const stored = window.localStorage.getItem(LOCALE_STORAGE_KEY);
  if (stored === 'zh-CN' || stored === 'en-US') return stored;
  return window.navigator.language.toLowerCase().startsWith('zh') ? 'zh-CN' : 'en-US';
}

export function persistLocale(locale: Locale): void {
  if (typeof window !== 'undefined') window.localStorage.setItem(LOCALE_STORAGE_KEY, locale);
}

export function translate(locale: Locale, key: MessageKey): string {
  return messages[locale][key];
}

export function translateRibbonTab(locale: Locale, tab: RibbonTabId): string {
  return translate(locale, tab as MessageKey);
}

export function translateRibbonText(locale: Locale, key: RibbonTextKey): string {
  const [section, name] = key.split('.') as ['groups' | 'commands', string];
  const bundle = localeBundles[locale].ribbon[section] as Record<string, string>;
  return bundle[name] ?? key;
}

export function formulaBarLabels(locale: Locale, phase: 'empty' | 'error' | 'loading' | 'ready') {
  const bundle = localeBundles[locale].formulaBar;
  return {
    selectedCell: bundle.selectedCell,
    formulaInput: bundle.formulaInput,
    insertFunction: bundle.insertFunction,
    cancel: bundle.cancel,
    apply: bundle.apply,
    applyHint: bundle.applyHint,
    placeholder: phase === 'empty' ? bundle.placeholderEmpty : bundle.placeholderReady,
  };
}

export function localizeText(locale: Locale, text: string): string {
  return locale === 'zh-CN' ? textTranslations[text] ?? text : text;
}

export function shellLabels(locale: Locale, saveState: SaveState) {
  return {
    workspaceLabel: translate(locale, 'workspaceLabel'),
    planningWorkbook: translate(locale, 'planningWorkbook'),
    personal: translate(locale, 'personal'),
    searchWorkbook: translate(locale, 'searchWorkbook'),
    share: translate(locale, 'share'),
    language: translate(locale, 'language'),
    english: translate(locale, 'english'),
    simplifiedChinese: translate(locale, 'simplifiedChinese'),
    saveState: translate(locale, saveState as MessageKey),
  };
}
