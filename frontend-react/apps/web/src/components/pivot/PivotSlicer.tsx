import { Button, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState } from 'react';
import type { PivotFieldDefinition } from '@react-sheets/core-model';

export interface PivotSlicerProps {
  field: PivotFieldDefinition;
  selectedValues: readonly string[];
  disabled?: boolean;
  onChange: (values: string[]) => void;
}

export function PivotSlicer({ disabled = false, field, onChange, selectedValues }: PivotSlicerProps) {
  const [search, setSearch] = useState('');
  const values = (field.values ?? []).map(String);
  const visibleValues = useMemo(() => values.filter((value) => value.toLowerCase().includes(search.toLowerCase())), [search, values]);
  const allSelected = selectedValues.length === 0 || selectedValues.length === values.length;
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/30 p-2">
      <Text size="xs" weight="semibold">Slicer · {field.name}</Text>
      <Button disabled={disabled} size="xs" variant="ghost" className="justify-start" onClick={() => onChange(allSelected ? [] : [...values])}>
        {allSelected ? 'Clear filter' : 'Select all'}
      </Button>
      <TextInput aria-label={`Search ${field.name}`} placeholder="Search items" value={search} onChange={(event) => setSearch(event.target.value)} />
      <Stack gap="xs" className="max-h-40 overflow-auto">
        {visibleValues.map((value) => {
          const selected = allSelected || selectedValues.includes(value);
          const currentSelection = selectedValues.length === 0 ? [...values] : [...selectedValues];
          return <Button key={value} disabled={disabled} aria-pressed={selected} size="xs" variant={selected ? 'soft' : 'ghost'} className="justify-start" onClick={() => onChange(selected ? currentSelection.filter((item) => item !== value) : [...currentSelection, value])}>{value}</Button>;
        })}
      </Stack>
    </Stack>
  );
}
