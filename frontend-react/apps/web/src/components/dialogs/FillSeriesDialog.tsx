import React, { useEffect, useState } from 'react';
import type { FillSeriesDateUnit, FillSeriesOptions, FillSeriesType } from '@react-sheets/sheet-features';
import { Button, CheckToggle, Dialog, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale } from '../home/home-localization';

export interface FillSeriesDialogProps {
  open: boolean;
  locale?: Locale;
  onClose: () => void;
  onApply: (options: FillSeriesOptions) => void;
}

export function FillSeriesDialog({ open, locale, onClose, onApply }: FillSeriesDialogProps): React.ReactElement | null {
  const activeLocale = resolveHomeLocale(locale);
  const [seriesIn, setSeriesIn] = useState<FillSeriesOptions['seriesIn']>('columns');
  const [type, setType] = useState<FillSeriesType>('linear');
  const [dateUnit, setDateUnit] = useState<FillSeriesDateUnit>('day');
  const [stepValue, setStepValue] = useState('1');
  const [stopValue, setStopValue] = useState('');
  const [trend, setTrend] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setSeriesIn('columns');
    setType('linear');
    setDateUnit('day');
    setStepValue('1');
    setStopValue('');
    setTrend(false);
    setError(null);
  }, [open]);

  const apply = (): void => {
    const step = Number(stepValue);
    const stop = stopValue.trim() === '' ? undefined : Number(stopValue);
    if (!Number.isFinite(step) || (stop !== undefined && !Number.isFinite(stop))) {
      setError(activeLocale === 'zh-CN' ? '步长和终止值必须是有效数字。' : 'Step and stop values must be finite numbers.');
      return;
    }
    onApply({ seriesIn, type, stepValue: step, ...(stop === undefined ? {} : { stopValue: stop }), ...(type === 'date' ? { dateUnit } : {}), ...(trend ? { trend: true } : {}) });
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'fillSeriesDialog')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="fill-series-dialog"
      maxWidth="sm"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button size="sm" variant="primary" data-testid="fill-series-apply" onClick={apply}>{homeText(activeLocale, 'ok')}</Button>
        </>
      )}
    >
      <Stack gap="md">
        <Stack gap="xs">
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillSeriesIn')}</Text>
          <Select sizeVariant="sm" value={seriesIn} onChange={(event) => setSeriesIn(event.target.value as FillSeriesOptions['seriesIn'])}>
            <option value="rows">{homeText(activeLocale, 'fillSeriesRows')}</option>
            <option value="columns">{homeText(activeLocale, 'fillSeriesColumns')}</option>
          </Select>
        </Stack>
        <Stack gap="xs">
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillSeriesType')}</Text>
          <Select sizeVariant="sm" value={type} onChange={(event) => setType(event.target.value as FillSeriesType)}>
            <option value="linear">{homeText(activeLocale, 'fillSeriesLinear')}</option>
            <option value="growth">{homeText(activeLocale, 'fillSeriesGrowth')}</option>
            <option value="date">{homeText(activeLocale, 'fillSeriesDate')}</option>
            <option value="autofill">{homeText(activeLocale, 'fillSeriesAutoFill')}</option>
          </Select>
        </Stack>
        {type === 'date' ? (
          <Stack gap="xs">
            <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillSeriesDateUnit')}</Text>
            <Select sizeVariant="sm" value={dateUnit} onChange={(event) => setDateUnit(event.target.value as FillSeriesDateUnit)}>
              <option value="day">{homeText(activeLocale, 'fillSeriesDay')}</option>
              <option value="weekday">{homeText(activeLocale, 'fillSeriesWeekday')}</option>
              <option value="month">{homeText(activeLocale, 'fillSeriesMonth')}</option>
              <option value="year">{homeText(activeLocale, 'fillSeriesYear')}</option>
            </Select>
          </Stack>
        ) : null}
        <Stack gap="xs">
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillSeriesStep')}</Text>
          <TextInput aria-label={homeText(activeLocale, 'fillSeriesStep')} value={stepValue} onChange={(event) => setStepValue(event.target.value)} inputMode="decimal" />
        </Stack>
        <Stack gap="xs">
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillSeriesStop')}</Text>
          <TextInput aria-label={homeText(activeLocale, 'fillSeriesStop')} value={stopValue} onChange={(event) => setStopValue(event.target.value)} inputMode="decimal" />
        </Stack>
        <CheckToggle checked={trend} label={homeText(activeLocale, 'fillSeriesTrend')} onChange={(event) => setTrend(event.target.checked)} />
        {error ? <Text size="xs" className="text-rose-700">{error}</Text> : null}
      </Stack>
    </Dialog>
  );
}
