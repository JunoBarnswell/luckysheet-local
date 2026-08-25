import enUS from '../../locales/en-US.json';
import zhCN from '../../locales/zh-CN.json';
import { getInitialLocale, type Locale } from '../../i18n';

export type HomeUiTextKey = keyof typeof enUS.homeUi;

export const HOME_NUMBER_FORMAT_OPTIONS = [
  { value: 'general', labelKey: 'numberPresetGeneral' },
  { value: '$#,##0', labelKey: 'numberPresetCurrency' },
  { value: '0%', labelKey: 'numberPresetPercent' },
  { value: '#,##0', labelKey: 'numberPresetComma' },
  { value: '0.00', labelKey: 'numberPresetNumber' },
] as const satisfies readonly { value: string; labelKey: HomeUiTextKey }[];

export const HOME_CELLS_ACTIONS = [
  { id: 'columnWidth', labelKey: 'columnWidth' },
  { id: 'autoFitColumnWidth', labelKey: 'autoFitColumnWidth' },
  { id: 'hideColumns', labelKey: 'hideColumns' },
  { id: 'unhideColumns', labelKey: 'unhideColumns' },
  { id: 'defaultColumnWidth', labelKey: 'defaultColumnWidth' },
] as const satisfies readonly { id: string; labelKey: HomeUiTextKey }[];

export function resolveHomeLocale(locale?: Locale): Locale {
  return locale ?? getInitialLocale();
}

export function homeText(locale: Locale | undefined, key: HomeUiTextKey): string {
  const activeLocale = resolveHomeLocale(locale);
  const bundle = activeLocale === 'zh-CN' ? zhCN.homeUi : enUS.homeUi;
  return bundle[key] ?? enUS.homeUi[key];
}

export function homeTemplate(locale: Locale | undefined, key: HomeUiTextKey, values: Record<string, string | number>): string {
  return Object.entries(values).reduce(
    (text, [name, value]) => text.replaceAll(`{${name}}`, String(value)),
    homeText(locale, key),
  );
}
