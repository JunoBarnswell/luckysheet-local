import { Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PivotAggregateFunction, PivotFieldDefinition, PivotShowAs, PivotValueField } from '@react-sheets/core-model';

export interface PivotValueEditorProps {
  value: PivotValueField;
  fields: readonly PivotFieldDefinition[];
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

export function PivotValueEditor({ disabled = false, fields, onChange, value }: PivotValueEditorProps) {
  const showAs = showAsKey(value.showAs);
  const fieldId = value.fieldId;
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/40 p-2">
      <Text size="xs" weight="semibold">Value settings · {fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId ?? 'Unknown field'}</Text>
      <Select aria-label="Summary function" disabled={disabled} sizeVariant="sm" value={value.summarizeBy} onChange={(event) => onChange({ ...value, summarizeBy: event.target.value as PivotAggregateFunction })}>
        {(['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const).map((option) => <option key={option} value={option}>{option}</option>)}
      </Select>
      <TextInput aria-label="Value display name" disabled={disabled} placeholder="Custom display name" value={value.displayName ?? ''} onChange={(event) => onChange({ ...value, displayName: event.target.value || undefined })} />
      <TextInput aria-label="Number format" disabled={disabled} placeholder="Number format, e.g. #,##0.00" value={value.numberFormat ?? ''} onChange={(event) => onChange({ ...value, numberFormat: event.target.value || undefined })} />
      <Select aria-label="Show value as" disabled={disabled} sizeVariant="sm" value={showAs} onChange={(event) => onChange({ ...value, showAs: showAsValue(event.target.value as ShowAsKey) })}>
        <option value="normal">Normal</option><option value="percent-of-total">% of grand total</option><option value="percent-of-row">% of row</option><option value="percent-of-column">% of column</option><option value="percent-of-parent">% of parent</option><option value="difference-from">Difference from</option><option value="percent-difference-from">% difference from</option><option value="running-total">Running total</option><option value="rank">Rank</option><option value="index">Index</option>
      </Select>
      {['difference-from', 'percent-difference-from'].includes(showAs) ? (
        <><Select aria-label="Base field" disabled={disabled} sizeVariant="sm" value={value.baseFieldId ?? ''} onChange={(event) => onChange({ ...value, baseFieldId: event.target.value || undefined })}><option value="">Base field</option>{fields.filter((field) => field.fieldId).map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}</Select><TextInput aria-label="Base item" disabled={disabled} placeholder="Base item" value={typeof value.baseItem === 'string' ? value.baseItem : value.baseItem ? String(value.baseItem.value ?? '') : ''} onChange={(event) => onChange({ ...value, baseItem: event.target.value || undefined })} /></>
      ) : null}
    </Stack>
  );
}
