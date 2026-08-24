import { Button, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotTimelineProps {
  fieldLabel: string;
  locale: Locale;
  start?: string;
  end?: string;
  disabled?: boolean;
  onChange: (start: string, end: string) => void;
  onClear: () => void;
}

export function PivotTimeline({ disabled = false, end = '', fieldLabel, locale, onChange, onClear, start = '' }: PivotTimelineProps) {
  return (
    <Stack gap="xs" className="rounded-lg border border-violet-100 bg-violet-50/30 p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'timelineTitle')} · {fieldLabel}</Text>
      <Inline gap="xs"><TextInput type="date" aria-label={pivotText(locale, 'timelineStart')} disabled={disabled} value={start} onChange={(event) => onChange(event.target.value, end)} /><TextInput type="date" aria-label={pivotText(locale, 'timelineEnd')} disabled={disabled} value={end} onChange={(event) => onChange(start, event.target.value)} /></Inline>
      <Button disabled={disabled} size="xs" variant="ghost" onClick={onClear}>{pivotText(locale, 'clearTimeline')}</Button>
    </Stack>
  );
}
