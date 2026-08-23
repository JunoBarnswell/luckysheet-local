import { Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PivotFieldDefinition, PivotValueDefinition } from './types';

export interface PivotValueEditorProps {
  value: PivotValueDefinition;
  fields: readonly PivotFieldDefinition[];
  disabled?: boolean;
  onChange: (value: PivotValueDefinition) => void;
}

export function PivotValueEditor({ disabled = false, fields, onChange, value }: PivotValueEditorProps) {
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/40 p-2">
      <Text size="xs" weight="semibold">Value settings · {fields.find((field) => field.id === value.fieldId)?.label ?? value.fieldId}</Text>
      <Select aria-label="Summary function" disabled={disabled} sizeVariant="sm" value={value.summary} onChange={(event) => onChange({ ...value, summary: event.target.value as PivotValueDefinition['summary'] })}>
        {(['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'] as const).map((option) => <option key={option} value={option}>{option}</option>)}
      </Select>
      <TextInput aria-label="Value display name" disabled={disabled} placeholder="Custom display name" value={value.displayName} onChange={(event) => onChange({ ...value, displayName: event.target.value })} />
      <TextInput aria-label="Number format" disabled={disabled} placeholder="Number format, e.g. #,##0.00" value={value.numberFormat} onChange={(event) => onChange({ ...value, numberFormat: event.target.value })} />
      <Select aria-label="Show value as" disabled={disabled} sizeVariant="sm" value={value.showAs} onChange={(event) => onChange({ ...value, showAs: event.target.value as PivotValueDefinition['showAs'] })}>
        <option value="normal">Normal</option><option value="percent-of-total">% of grand total</option><option value="percent-of-row">% of row</option><option value="percent-of-column">% of column</option><option value="percent-of-parent">% of parent</option><option value="difference-from">Difference from</option><option value="percent-difference-from">% difference from</option><option value="running-total">Running total</option><option value="rank">Rank</option><option value="index">Index</option>
      </Select>
      {['difference-from', 'percent-difference-from'].includes(value.showAs) ? (
        <><Select aria-label="Base field" disabled={disabled} sizeVariant="sm" value={value.baseFieldId ?? ''} onChange={(event) => onChange({ ...value, baseFieldId: event.target.value || undefined })}><option value="">Base field</option>{fields.map((field) => <option key={field.id} value={field.id}>{field.label}</option>)}</Select><TextInput aria-label="Base item" disabled={disabled} placeholder="Base item" value={value.baseItem ?? ''} onChange={(event) => onChange({ ...value, baseItem: event.target.value || undefined })} /></>
      ) : null}
    </Stack>
  );
}
