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
    conflict: 'Conflict',
    calculating: 'Calculating',
    error: 'Error',
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
    conflict: '存在冲突',
    calculating: '计算中',
    error: '错误',
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

export function localizeText(locale: Locale, text: string): string {
  return locale === 'zh-CN' ? textTranslations[text] ?? text : text;
}
