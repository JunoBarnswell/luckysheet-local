import { Box, Button, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState } from 'react';
import { createPivotMemberKey, formatPivotMember, pivotMemberKey, pivotMemberKeyEquals, type PivotFieldDefinition, type PivotMemberKey, type PivotScalar, type PivotSlicerItemProjection, type PivotSlicerSettings } from '@react-sheets/core-model';
import type { PivotFilterMode } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotSlicerProps {
  field: PivotFieldDefinition;
  memberOptions?: readonly PivotSlicerItem[];
  itemProjection?: readonly PivotSlicerItemProjection[];
  settings?: PivotSlicerSettings;
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
  hasData?: boolean;
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

const gridClasses: Record<number, string> = { 1: 'grid-cols-1', 2: 'grid-cols-2', 3: 'grid-cols-3', 4: 'grid-cols-4' };

export function PivotSlicer({ disabled = false, field, itemProjection, locale, memberKeys, mode, memberOptions, onChange, settings }: PivotSlicerProps) {
  const [search, setSearch] = useState('');
  const items = useMemo(() => itemProjection?.length ? itemProjection : memberOptions ?? buildPivotSlicerItems(field.values ?? []), [field.values, itemProjection, memberOptions]);
  const sortedItems = useMemo(() => {
    const ordered = [...items].sort((left, right) => left.label.localeCompare(right.label));
    if (settings?.sort === 'descending') ordered.reverse();
    if (settings?.noDataItemsLast) ordered.sort((left, right) => Number(right.hasData !== false) - Number(left.hasData !== false));
    return ordered;
  }, [items, settings?.noDataItemsLast, settings?.sort]);
  const visibleItems = useMemo(() => sortedItems.filter((item) => (settings?.showNoDataItems !== false || item.hasData !== false) && item.label.toLocaleLowerCase().includes(search.toLocaleLowerCase())), [search, settings?.showNoDataItems, sortedItems]);
  const allMembers = items.map((item) => item.key);
  const selected = (member: PivotMemberKey): boolean => {
    if (mode === 'all') return true;
    const matched = memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, member));
    return mode === 'include' ? matched : !matched;
  };
  const allSelected = mode === 'all' || (allMembers.length > 0 && allMembers.every((member) => selected(member)));
  const currentSelection = (): PivotMemberKey[] => allMembers.filter((member) => selected(member));
  const setAll = (next: boolean) => onChange(next ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: [] });
  const columns = settings?.columnCount ?? 1;
  return (
    <Stack gap="xs" className="rounded-lg border border-blue-100 bg-blue-50/30 p-2">
      {settings?.showHeader !== false ? <Text size="xs" weight="semibold">{settings?.caption || pivotText(locale, 'slicerTitle')} · {field.name}</Text> : null}
      <Button disabled={disabled} size="xs" variant="ghost" className="justify-start" onClick={() => setAll(!allSelected)}>
        {pivotText(locale, allSelected ? 'clearFilter' : 'selectAll')}
      </Button>
      <TextInput aria-label={`${pivotText(locale, 'searchItems')} ${field.name}`} placeholder={pivotText(locale, 'searchItems')} value={search} onChange={(event) => setSearch(event.target.value)} />
      <Stack gap="xs" className="max-h-40 overflow-auto">
        <Box className={`grid ${gridClasses[Math.min(4, Math.max(1, columns))] ?? 'grid-cols-1'} gap-1`}>
        {visibleItems.map((item) => {
          const isSelected = selected(item.key);
          return <Button key={pivotMemberKey(item.key)} disabled={disabled} aria-pressed={isSelected} size="xs" variant={isSelected ? 'soft' : 'ghost'} className={`justify-start ${item.hasData === false && settings?.showNoDataStyle !== false ? 'opacity-50' : ''}`} onClick={() => {
            if (settings?.multiSelect === false) { onChange({ mode: 'include', memberKeys: [item.key] }); return; }
            const next = isSelected ? currentSelection().filter((candidate) => !pivotMemberKeyEquals(candidate, item.key)) : [...currentSelection(), item.key];
            onChange(next.length === allMembers.length ? { mode: 'all', memberKeys: [] } : { mode: 'include', memberKeys: next });
          }}>{item.label}</Button>;
        })}
        </Box>
      </Stack>
    </Stack>
  );
}
