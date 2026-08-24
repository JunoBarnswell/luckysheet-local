import { Box, Button, CheckToggle, DropdownMenu, FieldDropZone, Icon, Inline, Select, Stack, Text } from '@react-sheets/ui-system';
import type { DragEvent, ReactNode } from 'react';
import {
  createPivotMemberKey,
  pivotMemberKey,
  pivotMemberKeyEquals,
  type PivotFieldDefinition,
  type PivotFieldPlacement,
  type PivotGroup,
  type PivotMemberKey,
  type PivotSort,
  type PivotScalar,
} from '@react-sheets/core-model';
import type { PivotFieldArea as Area, PivotManualFilterState } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

interface AreaItem extends PivotFieldDefinition {
  id: string;
  fieldId: string;
  index: number;
  placement?: PivotFieldPlacement;
}

export interface PivotFieldAreaProps {
  area: Area;
  fields: readonly PivotFieldDefinition[];
  fieldIds: readonly string[];
  disabled?: boolean;
  onDrop: (event: DragEvent<HTMLElement>, index?: number) => void;
  onRemove: (fieldId: string, index: number) => void;
  onMoveByKeyboard: (fieldId: string, index: number, direction: -1 | 1) => void;
  placements?: ReadonlyMap<string, PivotFieldPlacement>;
  filterStates?: Readonly<Record<string, PivotManualFilterState>>;
  onFilter?: (fieldId: string, filter: PivotManualFilterState) => void;
  onSort?: (fieldId: string, sort: PivotSort | undefined) => void;
  onGroup?: (fieldId: string, group: PivotGroup | undefined) => void;
  locale: Locale;
}

const labels: Record<Area, string> = { filters: 'FILTERS', columns: 'COLUMNS', rows: 'ROWS', values: 'VALUES' };
const icons: Record<Area, 'filter' | 'columns' | 'rows' | 'calculator'> = { filters: 'filter', columns: 'columns', rows: 'rows', values: 'calculator' };

function keyFor(value: PivotScalar): PivotMemberKey {
  return createPivotMemberKey(value);
}

function selectedMembers(field: AreaItem, state: PivotManualFilterState): PivotMemberKey[] {
  const all = (field.values ?? []).map(keyFor);
  if (state.mode === 'all') return all;
  if (state.mode === 'include') return state.memberKeys.filter((key) => all.some((candidate) => pivotMemberKeyEquals(candidate, key)));
  return all.filter((key) => !state.memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, key)));
}

function filterWithValue(field: AreaItem, state: PivotManualFilterState, value: PivotScalar, checked: boolean): PivotManualFilterState {
  const all = (field.values ?? []).map(keyFor);
  const target = keyFor(value);
  const selected = selectedMembers(field, state);
  const nextSelected = checked
    ? [...selected, target].filter((key, index, keys) => keys.findIndex((candidate) => pivotMemberKeyEquals(candidate, key)) === index)
    : selected.filter((key) => !pivotMemberKeyEquals(key, target));
  if (nextSelected.length >= all.length) return { mode: 'all', memberKeys: [] };
  return { mode: 'include', memberKeys: nextSelected };
}

function filterForMode(field: AreaItem, state: PivotManualFilterState, mode: PivotManualFilterState['mode']): PivotManualFilterState {
  if (mode === 'all') return { mode, memberKeys: [] };
  const selected = selectedMembers(field, state);
  if (mode === 'include') return { mode, memberKeys: selected };
  const all = (field.values ?? []).map(keyFor);
  return { mode, memberKeys: all.filter((key) => !selected.some((candidate) => pivotMemberKeyEquals(candidate, key))) };
}

function groupOptions(field: AreaItem, onGroup?: PivotFieldAreaProps['onGroup']): ReactNode {
  if (!onGroup) return null;
  if (field.dataType === 'date') {
    return (
      <Stack gap="xs" className="border-t border-slate-100 pt-1">
        <Text size="xs" weight="semibold">Group by</Text>
        <Inline gap="xs" className="flex-wrap">
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'year' })}>Years</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'quarter' })}>Quarters</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'month' })}>Months</Button>
        </Inline>
        <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>Clear group</Button>
      </Stack>
    );
  }
  if (field.dataType === 'number') {
    return (
      <Stack gap="xs" className="border-t border-slate-100 pt-1">
        <Text size="xs" weight="semibold">Group numbers</Text>
        <Inline gap="xs" className="flex-wrap">
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'number', interval: 10 })}>By 10</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'number', interval: 100 })}>By 100</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>Clear group</Button>
        </Inline>
      </Stack>
    );
  }
  const values = (field.values ?? []).map(keyFor);
  return values.length > 0
    ? <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'manual', groups: [{ groupId: `group:${field.fieldId}:all`, name: 'Group 1', items: values }] })}>Group all values</Button>
    : null;
}

export function PivotFieldArea({ area, disabled = false, fieldIds, fields, filterStates = {}, locale, onDrop, onFilter, onGroup, onMoveByKeyboard, onRemove, onSort, placements }: PivotFieldAreaProps) {
  const items: AreaItem[] = fieldIds.map((fieldId, index) => {
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    return {
      ...(field ?? { name: fieldId, dataType: 'text' as const, ordinal: index }),
      id: fieldId,
      fieldId,
      index,
      placement: placements?.get(fieldId),
    };
  });
  return (
    <Box as="section" aria-label={`${labels[area]} field area`} className="min-w-0 border-[#bdbdbd] bg-white">
      <Inline gap="xs" className="h-8 px-2">
        <Icon name={icons[area]} size="xs" className="text-accent" />
        <Text size="sm" weight="medium">{pivotText(locale, area)}</Text>
      </Inline>
      <FieldDropZone<AreaItem>
        disabled={disabled}
        emptyLabel=""
        className="h-[92px] overflow-auto rounded-none border-0 bg-white p-1"
        items={items}
        onDropItem={onDrop}
        renderItem={(field) => (
          <Inline
            draggable={!disabled}
            gap="xs"
            className="group min-h-8 cursor-grab rounded-md border border-blue-100 bg-white px-2 py-1 shadow-sm active:cursor-grabbing"
            onDragStart={(event) => event.dataTransfer.setData('application/x-pivot-field', field.fieldId)}
          >
            <Icon name="menu" size="xs" className="text-slate-300" />
            <Text size="xs" weight="medium" className="min-w-0 flex-1 truncate">{field.name}</Text>
            <DropdownMenu
              align="right"
              trigger={<Button aria-label={`Keyboard menu for ${field.name}`} icon="more-horizontal" iconOnly size="xs" variant="ghost" />}
            >
              {({ close }) => (
                <Inline gap="xs" className="p-1">
                   <Stack gap="xs" className="min-w-48 p-1">
                     <Inline gap="xs">
                       <Button disabled={field.index === 0} icon="arrow-up" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.fieldId, field.index, -1); close(); }} />
                       <Button disabled={field.index === items.length - 1} icon="arrow-down" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.fieldId, field.index, 1); close(); }} />
                       <Button icon="trash" iconOnly size="xs" variant="danger" onClick={() => { onRemove(field.fieldId, field.index); close(); }} />
                     </Inline>
                     {onSort ? <Stack gap="xs" className="border-t border-slate-100 pt-1"><Text size="xs" weight="semibold">Sort by</Text><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'ascending', by: 'label' }); close(); }}>A–Z</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'descending', by: 'label' }); close(); }}>Z–A</Button></Inline><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'ascending', by: 'value' }); close(); }}>Value ↑</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'descending', by: 'value' }); close(); }}>Value ↓</Button></Inline><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, undefined); close(); }}>Clear sort</Button></Stack> : null}
                     {groupOptions(field, onGroup)}
                     {onFilter && field.values?.length ? (() => {
                       const state = filterStates[field.fieldId] ?? { mode: 'all' as const, memberKeys: [] };
                       return <Stack gap="xs" className="border-t border-slate-100 pt-1"><Text size="xs" weight="semibold">Filter values</Text><Select aria-label={`${field.name} filter mode`} sizeVariant="sm" disabled={disabled} value={state.mode} onChange={(event) => onFilter(field.fieldId, filterForMode(field, state, event.target.value as PivotManualFilterState['mode']))}><option value="all">Show all</option><option value="include">Include selected</option><option value="exclude">Exclude selected</option></Select>{field.values.map((value) => { const textValue = String(value); const selected = selectedMembers(field, state).some((key) => pivotMemberKeyEquals(key, keyFor(value))); return <CheckToggle key={pivotMemberKey(keyFor(value))} label={textValue} checked={selected} onChange={(event) => onFilter(field.fieldId, filterWithValue(field, state, value, event.target.checked))} />; })}</Stack>;
                     })() : null}
                   </Stack>
                </Inline>
              )}
            </DropdownMenu>
          </Inline>
        )}
        onDragOverItem={(event) => { event.dataTransfer.dropEffect = 'move'; }}
      />
    </Box>
  );
}
