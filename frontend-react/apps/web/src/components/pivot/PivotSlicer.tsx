import { Button, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState } from 'react';
import { createPivotMemberKey, formatPivotMember, pivotMemberKey, pivotMemberKeyEquals, type PivotFieldDefinition, type PivotMemberKey, type PivotScalar } from '@react-sheets/core-model';
import type { PivotFilterMode } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotSlicerProps {
  field: PivotFieldDefinition;
  memberOptions?: readonly PivotSlicerItem[];
  locale: Locale;
  mode: PivotFilterMode;
  memberKeys: readonly PivotMemberKey[];
  disabled?: boolean;
  onChange: (state: { mode: PivotFilterMode; memberKeys: PivotMemberKey[] }) => void;
}

export interface PivotSlicerItem {
  key: PivotMemberKey;
  value: PivotScalar;
  label: string;
}

function keyFor(value: PivotScalar): PivotMemberKey { return createPivotMemberKey(value); }

export function buildPivotSlicerItems(values: readonly PivotScalar[]): PivotSlicerItem[] {
  const seen = new Set<string>();
  return values.flatMap((value) => {
    const key = keyFor(value);
    const identity = pivotMemberKey(key);
    if (seen.has(identity)) return [];
    seen.add(identity);
    return [{ key, value, label: formatPivotMember(value) }];
  });
}

export function PivotSlicer({ disabled = false, field, locale, memberKeys, mode, memberOptions, onChange }: PivotSlicerProps) {
  const [search, setSearch] = useState('');
  const items = useMemo(() => memberOptions ?? buildPivotSlicerItems(field.values ?? []), [field.values, memberOptions]);
  const visibleItems = useMemo(() => items.filter((item) => item.label.toLocaleLowerCase().includes(search.toLocaleLowerCase())), [items, search]);
  const allMembers = items.map((item) => item.key);
  const selected = (member: PivotMemberKey): boolean => {
    if (mode === 'all') return true;
    const matched = memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, member));
    return mode === 'include' ? matched : !matched;
  };
  const allSelected = mode === 'all' || (allMembers.length > 0 && allMembers.every((member) => selected(member)));
  const currentSelection = (): PivotMemberKey[] => allMembers.filter((member) => selected(member));
  const setAll = (next: boolean) => onChange(next ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: [] });
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/30 p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'slicerTitle')} · {field.name}</Text>
      <Button disabled={disabled} size="xs" variant="ghost" className="justify-start" onClick={() => setAll(!allSelected)}>
        {pivotText(locale, allSelected ? 'clearFilter' : 'selectAll')}
      </Button>
      <TextInput aria-label={`${pivotText(locale, 'searchItems')} ${field.name}`} placeholder={pivotText(locale, 'searchItems')} value={search} onChange={(event) => setSearch(event.target.value)} />
      <Stack gap="xs" className="max-h-40 overflow-auto">
        {visibleItems.map((item) => {
          const isSelected = selected(item.key);
          return <Button key={pivotMemberKey(item.key)} disabled={disabled} aria-pressed={isSelected} size="xs" variant={isSelected ? 'soft' : 'ghost'} className="justify-start" onClick={() => { const next = isSelected ? currentSelection().filter((candidate) => !pivotMemberKeyEquals(candidate, item.key)) : [...currentSelection(), item.key]; onChange(next.length === allMembers.length ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: next }); }}>{item.label}</Button>;
        })}
      </Stack>
    </Stack>
  );
}
