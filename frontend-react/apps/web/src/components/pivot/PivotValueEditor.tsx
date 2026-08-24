import { Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PivotAggregateFunction, PivotFieldDefinition, PivotShowAs, PivotValueField } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText, type PivotMessageKey } from './pivot-localization';

export interface PivotValueEditorProps {
  value: PivotValueField;
  fields: readonly PivotFieldDefinition[];
  locale: Locale;
  disabled?: boolean;
  onChange: (value: PivotValueField) => void;
}

type ShowAsKey = 'normal' | 'percent-of-total' | 'percent-of-row' | 'percent-of-column' | 'percent-of-parent' | 'difference-from' | 'percent-difference-from' | 'running-total' | 'rank' | 'index';

function showAsKey(showAs?: PivotShowAs): ShowAsKey {
  switch (showAs?.kind) {
    case 'grand-percentage': return 'percent-of-total';
    case 'row-percentage': return 'percent-of-row';
    case 'column-percentage': return 'percent-of-column';
    case 'parent-percentage': return 'percent-of-parent';
    case 'difference': return 'difference-from';
    case 'percentage-difference': return 'percent-difference-from';
    case 'running-total': return 'running-total';
    case 'rank': return 'rank';
    case 'index': return 'index';
    default: return 'normal';
  }
}

function showAsValue(value: ShowAsKey): PivotShowAs {
  if (value === 'percent-of-total') return { kind: 'grand-percentage' };
  if (value === 'percent-of-row') return { kind: 'row-percentage' };
  if (value === 'percent-of-column') return { kind: 'column-percentage' };
  if (value === 'percent-of-parent') return { kind: 'parent-percentage' };
  if (value === 'difference-from') return { kind: 'difference', base: 'grand' };
  if (value === 'percent-difference-from') return { kind: 'percentage-difference', base: 'grand' };
  if (value === 'running-total') return { kind: 'running-total', axis: 'row' };
  if (value === 'rank') return { kind: 'rank', axis: 'row', direction: 'descending' };
  if (value === 'index') return { kind: 'index' };
  return { kind: 'normal' };
}

const aggregateKeys: Record<PivotAggregateFunction, PivotMessageKey> = { sum: 'sum', count: 'count', 'count-numbers': 'countNumbers', average: 'average', min: 'min', max: 'max', product: 'product', stdev: 'stdev', stdevp: 'stdevp', var: 'variance', varp: 'variancep', 'distinct-count': 'distinctCount' };
const showAsKeys: Record<ShowAsKey, PivotMessageKey> = { normal: 'normal', 'percent-of-total': 'grandPercent', 'percent-of-row': 'rowPercent', 'percent-of-column': 'columnPercent', 'percent-of-parent': 'parentPercent', 'difference-from': 'differenceFrom', 'percent-difference-from': 'percentDifferenceFrom', 'running-total': 'runningTotal', rank: 'rank', index: 'index' };

export function PivotValueEditor({ disabled = false, fields, locale, onChange, value }: PivotValueEditorProps) {
  const showAs = showAsKey(value.showAs);
  const fieldId = value.fieldId;
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/40 p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'valueSettings')} · {fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId}</Text>
      <Select aria-label={pivotText(locale, 'summaryFunction')} disabled={disabled} sizeVariant="sm" value={value.summarizeBy} onChange={(event) => onChange({ ...value, summarizeBy: event.target.value as PivotAggregateFunction })}>
        {(['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const).map((option) => <option key={option} value={option}>{pivotText(locale, aggregateKeys[option])}</option>)}
      </Select>
      <TextInput aria-label={pivotText(locale, 'displayName')} disabled={disabled} placeholder={pivotText(locale, 'displayNamePlaceholder')} value={value.displayName ?? ''} onChange={(event) => onChange({ ...value, displayName: event.target.value || undefined })} />
      <TextInput aria-label={pivotText(locale, 'numberFormat')} disabled={disabled} placeholder={pivotText(locale, 'numberFormatPlaceholder')} value={value.numberFormat ?? ''} onChange={(event) => onChange({ ...value, numberFormat: event.target.value || undefined })} />
      <Select aria-label={pivotText(locale, 'showValueAs')} disabled={disabled} sizeVariant="sm" value={showAs} onChange={(event) => onChange({ ...value, showAs: showAsValue(event.target.value as ShowAsKey) })}>
        {(Object.keys(showAsKeys) as ShowAsKey[]).map((key) => <option key={key} value={key}>{pivotText(locale, showAsKeys[key])}</option>)}
      </Select>
      {['difference-from', 'percent-difference-from'].includes(showAs) ? (
        <><Select aria-label={pivotText(locale, 'baseField')} disabled={disabled} sizeVariant="sm" value={value.baseFieldId ?? ''} onChange={(event) => onChange({ ...value, baseFieldId: event.target.value || undefined })}><option value="">{pivotText(locale, 'baseField')}</option>{fields.map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}</Select><TextInput aria-label={pivotText(locale, 'baseItem')} disabled={disabled} placeholder={pivotText(locale, 'baseItem')} value={typeof value.baseItem === 'string' ? value.baseItem : value.baseItem ? String(value.baseItem.value ?? '') : ''} onChange={(event) => onChange({ ...value, baseItem: event.target.value || undefined })} /></>
      ) : null}
    </Stack>
  );
}
