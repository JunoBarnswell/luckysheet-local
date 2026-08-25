import { Button, CheckToggle, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { buildPivotTimelineTiles, type PivotScalar, type PivotTimelineDrawingPayload, type PivotTimelineLevel, type PivotTimelinePeriod } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotTimelineProps {
  fieldLabel: string;
  locale: Locale;
  start?: string;
  end?: string;
  values?: readonly PivotScalar[];
  level: PivotTimelineLevel;
  showHeader: boolean;
  showSelectionLabel: boolean;
  showTimeLevel: boolean;
  showHorizontalScrollbar: boolean;
  bounds: PivotTimelinePeriod;
  caption?: string;
  styleName?: string;
  disabled?: boolean;
  onChange: (start: string, end: string) => void;
  onClear: () => void;
  onLevelChange: (level: PivotTimelineLevel) => void;
  onWindowChange: (bounds: PivotTimelinePeriod) => void;
  onDisplayChange: (display: Pick<PivotTimelineDrawingPayload, 'showHeader' | 'showSelectionLabel' | 'showTimeLevel' | 'showHorizontalScrollbar'>) => void;
  onCaptionChange: (caption: string) => void;
  onStyleChange: (styleName: string) => void;
}

export function PivotTimeline({ bounds, caption, disabled = false, end = '', fieldLabel, level, locale, onCaptionChange, onChange, onClear, onDisplayChange, onLevelChange, onStyleChange, onWindowChange, showHeader, showHorizontalScrollbar, showSelectionLabel, showTimeLevel, start = '', styleName = 'TimelineStyleLight2', values = [] }: PivotTimelineProps) {
  const tiles = buildPivotTimelineTiles(values, level);
  const windowStart = bounds.start;
  const windowEnd = bounds.end;
  const visibleTiles = tiles.filter((tile) => (!windowStart || tile.end >= windowStart) && (!windowEnd || tile.start <= windowEnd));
  const selected = (tile: { start: string; end: string }) => (!start || tile.end >= start) && (!end || tile.start <= end);
  const shiftWindow = (direction: -1 | 1) => {
    if (tiles.length === 0) return;
    const currentStart = windowStart ? Math.max(0, tiles.findIndex((tile) => tile.start >= windowStart)) : 0;
    const currentEnd = windowEnd ? Math.max(currentStart, tiles.findIndex((tile) => tile.end >= windowEnd)) : Math.min(tiles.length - 1, currentStart + 7);
    const width = Math.max(0, currentEnd - currentStart);
    const nextStart = Math.max(0, Math.min(tiles.length - 1 - width, currentStart + direction * Math.max(1, width)));
    const nextEnd = Math.min(tiles.length - 1, nextStart + width);
    onWindowChange({ start: tiles[nextStart]?.start, end: tiles[nextEnd]?.end });
  };
  return (
    <Stack gap="xs" className="rounded-lg border border-violet-100 bg-violet-50/30 p-2">
      {showHeader ? <Text size="xs" weight="semibold">{caption || pivotText(locale, 'timelineTitle')} · {fieldLabel}</Text> : null}
      <Inline gap="xs" className="items-center">
        {showTimeLevel ? <Select aria-label={pivotText(locale, 'timelineLevel')} disabled={disabled} sizeVariant="sm" value={level} onChange={(event) => onLevelChange(event.target.value as PivotTimelineLevel)}><option value="years">{pivotText(locale, 'years')}</option><option value="quarters">{pivotText(locale, 'quarters')}</option><option value="months">{pivotText(locale, 'months')}</option><option value="days">{pivotText(locale, 'days')}</option></Select> : null}
        <Select aria-label={pivotText(locale, 'timelineStyle')} disabled={disabled} sizeVariant="sm" value={styleName} onChange={(event) => onStyleChange(event.target.value)}><option value="TimelineStyleLight2">{pivotText(locale, 'styleLight')}</option><option value="TimelineStyleMedium2">{pivotText(locale, 'styleMedium')}</option><option value="TimelineStyleDark2">{pivotText(locale, 'styleDark')}</option></Select>
        <TextInput aria-label={pivotText(locale, 'timelineCaption')} disabled={disabled} value={caption} placeholder={fieldLabel} onChange={(event) => onCaptionChange(event.target.value)} />
      </Inline>
      <Inline gap="xs" className="items-center">
        {showHorizontalScrollbar ? <Button disabled={disabled || tiles.length === 0} size="xs" variant="ghost" onClick={() => shiftWindow(-1)} aria-label={pivotText(locale, 'timelinePrevious')}>‹</Button> : null}
        <Inline gap="none" className="min-w-0 flex-1 overflow-hidden">
          {visibleTiles.map((tile) => <Button key={tile.key} disabled={disabled} size="xs" variant={selected(tile) ? 'soft' : 'ghost'} className="min-w-[3.5rem] rounded-none px-1 text-[10px]" aria-pressed={selected(tile)} onClick={() => onChange(tile.start, tile.end)}>{tile.label}</Button>)}
        </Inline>
        {showHorizontalScrollbar ? <Button disabled={disabled || tiles.length === 0} size="xs" variant="ghost" onClick={() => shiftWindow(1)} aria-label={pivotText(locale, 'timelineNext')}>›</Button> : null}
      </Inline>
      {showSelectionLabel ? <Text size="xs" tone="muted">{start || end ? `${start || '…'} — ${end || '…'}` : pivotText(locale, 'timelineAllPeriods')}</Text> : null}
      <Inline gap="xs" className="flex-wrap">
        <CheckToggle checked={showHeader} label={pivotText(locale, 'timelineHeader')} onChange={(event) => onDisplayChange({ showHeader: event.target.checked, showSelectionLabel, showTimeLevel, showHorizontalScrollbar })} />
        <CheckToggle checked={showSelectionLabel} label={pivotText(locale, 'timelineSelection')} onChange={(event) => onDisplayChange({ showHeader, showSelectionLabel: event.target.checked, showTimeLevel, showHorizontalScrollbar })} />
      </Inline>
      <Inline gap="xs"><TextInput type="date" aria-label={pivotText(locale, 'timelineStart')} disabled={disabled} value={start} onChange={(event) => onChange(event.target.value, end)} /><TextInput type="date" aria-label={pivotText(locale, 'timelineEnd')} disabled={disabled} value={end} onChange={(event) => onChange(start, event.target.value)} /><Button disabled={disabled} size="xs" variant="ghost" onClick={onClear}>{pivotText(locale, 'clearTimeline')}</Button></Inline>
    </Stack>
  );
}
