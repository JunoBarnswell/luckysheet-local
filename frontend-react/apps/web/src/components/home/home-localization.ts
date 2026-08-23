import enUS from '../../locales/en-US.json';
import zhCN from '../../locales/zh-CN.json';
import { getInitialLocale, type Locale } from '../../i18n';

export type HomeUiTextKey = keyof typeof enUS.homeUi;

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
