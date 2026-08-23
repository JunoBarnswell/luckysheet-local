export type Locale = 'zh-CN' | 'en-US';

export type MessageKey =
  | 'workspaceLabel'
  | 'planningWorkbook'
  | 'personal'
  | 'saved'
  | 'saving'
  | 'offline'
  | 'syncing'
  | 'share'
  | 'language'
  | 'english'
  | 'simplifiedChinese'
  | 'home'
  | 'insert'
  | 'data'
  | 'review'
  | 'view'
  | 'engineConnected'
  | 'searchWorkbook';

const messages: Record<Locale, Record<MessageKey, string>> = {
  'en-US': {
    workspaceLabel: 'Workspace',
    planningWorkbook: 'Planning workbook',
    personal: 'Personal',
    saved: 'Saved',
    saving: 'Saving',
    offline: 'Offline',
    syncing: 'Syncing',
    share: 'Share',
    language: 'Language',
    english: 'English',
    simplifiedChinese: '中文',
    home: 'Home',
    insert: 'Insert',
    data: 'Data',
    review: 'Review',
    view: 'View',
    engineConnected: 'Engine Connected',
    searchWorkbook: 'Search workbook',
  },
  'zh-CN': {
    workspaceLabel: '工作区',
    planningWorkbook: '规划工作簿',
    personal: '个人',
    saved: '已保存',
    saving: '保存中',
    offline: '离线',
    syncing: '同步中',
    share: '共享',
    language: '语言',
    english: 'English',
    simplifiedChinese: '中文',
    home: '开始',
    insert: '插入',
    data: '数据',
    review: '审阅',
    view: '视图',
    engineConnected: '引擎已连接',
    searchWorkbook: '搜索工作簿',
  },
};

export const localeLabels: Record<Locale, string> = {
  'en-US': 'EN / 中文',
  'zh-CN': '中文 / EN',
};

const LOCALE_STORAGE_KEY = 'react-sheets:locale';

export function getInitialLocale(): Locale {
  if (typeof window === 'undefined') return 'en-US';
  return window.localStorage.getItem(LOCALE_STORAGE_KEY) === 'zh-CN' ? 'zh-CN' : 'en-US';
}

export function persistLocale(locale: Locale): void {
  if (typeof window !== 'undefined') window.localStorage.setItem(LOCALE_STORAGE_KEY, locale);
}

export function translate(locale: Locale, key: MessageKey): string {
  return messages[locale][key];
}
