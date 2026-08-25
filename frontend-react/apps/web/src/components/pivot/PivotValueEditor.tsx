import { Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { createPivotMemberKey, type PivotAggregateFunction, type PivotFieldDefinition, type PivotMemberKey, type PivotShowAs, type PivotShowAsBaseItem, type PivotValueField } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText, type PivotMessageKey } from './pivot-localization';

export interface PivotValueEditorProps {
  value: PivotValueField;
  fields: readonly PivotFieldDefinition[];
  baseFields?: readonly PivotFieldDefinition[];
  locale: Locale;
  disabled?: boolean;
  onChange: (value: PivotValueField) => void;
}

type ShowAsKey = 'normal' | 'percent-of-total' | 'percent-of-row' | 'percent-of-column' | 'percent-of-parent' | 'difference-from' | 'percent-difference-from' | 'running-total' | 'percent-running-total' | 'rank' | 'index';

function showAsKey(showAs?: PivotShowAs): ShowAsKey {
  switch (showAs?.kind) {
    case 'grand-percentage': return 'percent-of-total';
    case 'row-percentage': return 'percent-of-row';
    case 'column-percentage': return 'percent-of-column';
    case 'parent-percentage': return 'percent-of-parent';
    case 'difference': return 'difference-from';
    case 'percentage-difference': return 'percent-difference-from';
    case 'running-total': return 'running-total';
    case 'percentage-running-total': return 'percent-running-total';
    case 'rank': return 'rank';
    case 'index': return 'index';
    default: return 'normal';
  }
}

function showAsValue(value: ShowAsKey, baseFields: readonly PivotFieldDefinition[], previous?: PivotShowAs): PivotShowAs {
  const previousBase = previous && 'baseFieldId' in previous ? previous.baseFieldId : undefined;
  const baseFieldId = previousBase && baseFields.some((field) => field.fieldId === previousBase) ? previousBase : baseFields[0]?.fieldId ?? '';
  const baseField = baseFields.find((field) => field.fieldId === baseFieldId);
  const baseItem = previous && 'baseItem' in previous ? previous.baseItem : createPivotMemberKey(baseField?.values?.[0] ?? null);
  if (value === 'percent-of-total') return { kind: 'grand-percentage' };
  if (value === 'percent-of-row') return { kind: 'row-percentage' };
  if (value === 'percent-of-column') return { kind: 'column-percentage' };
  if (value === 'percent-of-parent') return { kind: 'parent-percentage' };
  if (value === 'difference-from') return { kind: 'difference', baseFieldId, baseItem };
  if (value === 'percent-difference-from') return { kind: 'percentage-difference', baseFieldId, baseItem };
  if (value === 'running-total') return { kind: 'running-total', baseFieldId };
  if (value === 'percent-running-total') return { kind: 'percentage-running-total', baseFieldId };
  if (value === 'rank') return { kind: 'rank', baseFieldId, direction: previous?.kind === 'rank' ? previous.direction : 'descending' };
  if (value === 'index') return { kind: 'index' };
  return { kind: 'normal' };
}

const aggregateKeys: Record<PivotAggregateFunction, PivotMessageKey> = { sum: 'sum', count: 'count', 'count-numbers': 'countNumbers', average: 'average', min: 'min', max: 'max', product: 'product', stdev: 'stdev', stdevp: 'stdevp', var: 'variance', varp: 'variancep', 'distinct-count': 'distinctCount' };
const showAsKeys: Record<ShowAsKey, PivotMessageKey> = { normal: 'normal', 'percent-of-total': 'grandPercent', 'percent-of-row': 'rowPercent', 'percent-of-column': 'columnPercent', 'percent-of-parent': 'parentPercent', 'difference-from': 'differenceFrom', 'percent-difference-from': 'percentDifferenceFrom', 'running-total': 'runningTotal', 'percent-running-total': 'percentRunningTotal', rank: 'rank', index: 'index' };

export function PivotValueEditor({ baseFields = [], disabled = false, fields, locale, onChange, value }: PivotValueEditorProps) {
  const showAs = showAsKey(value.showAs);
  const fieldId = value.fieldId;
  const showAsDefinition = value.showAs;
  const baseFieldId = showAsDefinition && 'baseFieldId' in showAsDefinition ? showAsDefinition.baseFieldId : baseFields[0]?.fieldId ?? '';
  const baseField = baseFields.find((field) => field.fieldId === baseFieldId);
  const baseItem = showAsDefinition && 'baseItem' in showAsDefinition ? showAsDefinition.baseItem : createPivotMemberKey(baseField?.values?.[0] ?? null);
  const baseItemToken = (item: PivotShowAsBaseItem): string => typeof item === 'string' ? item : JSON.stringify(item);
  const parseBaseItem = (token: string): PivotShowAsBaseItem => token === 'previous' || token === 'next' ? token : JSON.parse(token) as PivotMemberKey;
  const baseItemOptions = baseField?.values?.map(createPivotMemberKey) ?? [];
  const changeBaseField = (nextFieldId: string) => {
    const nextField = baseFields.find((field) => field.fieldId === nextFieldId);
    const nextItem = createPivotMemberKey(nextField?.values?.[0] ?? null);
    if (showAsDefinition && 'baseItem' in showAsDefinition) onChange({ ...value, showAs: { ...showAsDefinition, baseFieldId: nextFieldId, baseItem: nextItem } });
    else if (showAsDefinition && 'baseFieldId' in showAsDefinition) onChange({ ...value, showAs: { ...showAsDefinition, baseFieldId: nextFieldId } });
  };
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/40 p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'valueSettings')} · {fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId}</Text>
      <Select aria-label={pivotText(locale, 'summaryFunction')} disabled={disabled} sizeVariant="sm" value={value.summarizeBy} onChange={(event) => onChange({ ...value, summarizeBy: event.target.value as PivotAggregateFunction })}>
        {(['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const).map((option) => <option key={option} value={option}>{pivotText(locale, aggregateKeys[option])}</option>)}
      </Select>
      <TextInput aria-label={pivotText(locale, 'displayName')} disabled={disabled} placeholder={pivotText(locale, 'displayNamePlaceholder')} value={value.displayName ?? ''} onChange={(event) => onChange({ ...value, displayName: event.target.value || undefined })} />
      <TextInput aria-label={pivotText(locale, 'numberFormat')} disabled={disabled} placeholder={pivotText(locale, 'numberFormatPlaceholder')} value={value.numberFormat ?? ''} onChange={(event) => onChange({ ...value, numberFormat: event.target.value || undefined })} />
      <Select aria-label={pivotText(locale, 'showValueAs')} disabled={disabled} sizeVariant="sm" value={showAs} onChange={(event) => onChange({ ...value, showAs: showAsValue(event.target.value as ShowAsKey, baseFields, value.showAs) })}>
        {(Object.keys(showAsKeys) as ShowAsKey[]).map((key) => <option key={key} value={key}>{pivotText(locale, showAsKeys[key])}</option>)}
      </Select>
      {['difference-from', 'percent-difference-from', 'running-total', 'percent-running-total', 'rank'].includes(showAs) ? (
        <Select aria-label={pivotText(locale, 'baseField')} disabled={disabled || baseFields.length === 0} sizeVariant="sm" value={baseFieldId} onChange={(event) => changeBaseField(event.target.value)}>
          <option value="" disabled>{pivotText(locale, 'baseField')}</option>{baseFields.map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}
        </Select>
      ) : null}
      {['difference-from', 'percent-difference-from'].includes(showAs) ? (
        <Select aria-label={pivotText(locale, 'baseItem')} disabled={disabled || !baseField} sizeVariant="sm" value={baseItemToken(baseItem)} onChange={(event) => { if (showAsDefinition && 'baseItem' in showAsDefinition) onChange({ ...value, showAs: { ...showAsDefinition, baseItem: parseBaseItem(event.target.value) } }); }}>
          <option value="previous">Previous</option><option value="next">Next</option>{baseItemOptions.map((item) => <option key={baseItemToken(item)} value={baseItemToken(item)}>{String(item.value ?? '(blank)')}</option>)}
        </Select>
      ) : null}
      {showAs === 'rank' ? (
        <Select aria-label={pivotText(locale, 'rank')} disabled={disabled || !baseField} sizeVariant="sm" value={showAsDefinition?.kind === 'rank' ? showAsDefinition.direction : 'descending'} onChange={(event) => { if (showAsDefinition?.kind === 'rank') onChange({ ...value, showAs: { ...showAsDefinition, direction: event.target.value as 'ascending' | 'descending' } }); }}><option value="ascending">{pivotText(locale, 'ascending')}</option><option value="descending">{pivotText(locale, 'descending')}</option></Select>
      ) : null}
    </Stack>
  );
}
