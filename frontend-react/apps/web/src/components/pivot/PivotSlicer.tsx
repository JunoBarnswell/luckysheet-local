import { Button, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState } from 'react';
import { createPivotMemberKey, pivotMemberKey, pivotMemberKeyEquals, type PivotFieldDefinition, type PivotMemberKey, type PivotScalar } from '@react-sheets/core-model';
import type { PivotFilterMode } from './pivot-contract';

export interface PivotSlicerProps {
  field: PivotFieldDefinition;
  mode: PivotFilterMode;
  memberKeys: readonly PivotMemberKey[];
  disabled?: boolean;
  onChange: (state: { mode: PivotFilterMode; memberKeys: PivotMemberKey[] }) => void;
}

function keyFor(value: PivotScalar): PivotMemberKey {
  return createPivotMemberKey(value);
}

export function PivotSlicer({ disabled = false, field, memberKeys, mode, onChange }: PivotSlicerProps) {
  const [search, setSearch] = useState('');
  const values = (field.values ?? []).map(String);
  const visibleValues = useMemo(() => values.filter((value) => value.toLowerCase().includes(search.toLowerCase())), [search, values]);
  const allMembers = (field.values ?? []).map(keyFor);
  const selected = (value: PivotScalar): boolean => {
    const member = keyFor(value);
    if (mode === 'all') return true;
    const matched = memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, member));
    return mode === 'include' ? matched : !matched;
  };
  const allSelected = mode === 'all' || values.every((value) => selected(value));
  const currentSelection = (): PivotMemberKey[] => allMembers.filter((member) => selected(member.value as PivotScalar));
  const setAll = (next: boolean) => onChange(next ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: [] });
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/30 p-2">
      <Text size="xs" weight="semibold">Slicer · {field.name}</Text>
      <Button disabled={disabled} size="xs" variant="ghost" className="justify-start" onClick={() => setAll(!allSelected)}>
        {allSelected ? 'Clear filter' : 'Select all'}
      </Button>
      <TextInput aria-label={`Search ${field.name}`} placeholder="Search items" value={search} onChange={(event) => setSearch(event.target.value)} />
      <Stack gap="xs" className="max-h-40 overflow-auto">
        {visibleValues.map((value) => {
          const rawValue = (field.values ?? []).find((candidate) => String(candidate) === value) ?? value;
          const isSelected = selected(rawValue);
          return <Button key={pivotMemberKey(keyFor(rawValue))} disabled={disabled} aria-pressed={isSelected} size="xs" variant={isSelected ? 'soft' : 'ghost'} className="justify-start" onClick={() => { const next = isSelected ? currentSelection().filter((candidate) => !pivotMemberKeyEquals(candidate, keyFor(rawValue))) : [...currentSelection(), keyFor(rawValue)]; onChange(next.length === allMembers.length ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: next }); }}>{value}</Button>;
        })}
      </Stack>
    </Stack>
  );
}
