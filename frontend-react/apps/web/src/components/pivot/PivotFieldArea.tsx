import { Box, Button, CheckToggle, DropdownMenu, FieldDropZone, Icon, Inline, Select, Stack, Text } from '@react-sheets/ui-system';
import type { DragEvent, ReactNode } from 'react';
import {
  createPivotMemberKey,
  type PivotAggregateFunction,
  pivotMemberKey,
  pivotMemberKeyEquals,
  type PivotFieldDefinition,
  type PivotFieldPlacement,
  type PivotGroup,
  type PivotMemberKey,
  type PivotSort,
  type PivotScalar,
  type PivotSubtotalDefinition,
  type PivotValueField,
} from '@react-sheets/core-model';
import type { PivotFieldArea as Area, PivotManualFilterState } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';
import { PivotValueEditor } from './PivotValueEditor';

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
  onSubtotal?: (fieldId: string, subtotal: PivotSubtotalDefinition) => void;
  valueFields?: readonly PivotValueField[];
  onValueChange?: (value: PivotValueField) => void;
  locale: Locale;
  className?: string;
}

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

function filterOptions(locale: Locale, field: AreaItem, disabled: boolean, state: PivotManualFilterState, onFilter: NonNullable<PivotFieldAreaProps['onFilter']>): ReactNode {
  return (
    <Stack gap="xs" className="max-h-[18rem] min-w-[13rem] overflow-auto p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'filterValues')}</Text>
      <Select aria-label={`${field.name} ${pivotText(locale, 'filterMode')}`} sizeVariant="sm" disabled={disabled} value={state.mode} onChange={(event) => onFilter(field.fieldId, filterForMode(field, state, event.target.value as PivotManualFilterState['mode']))}>
        <option value="all">{pivotText(locale, 'showAll')}</option>
        <option value="include">{pivotText(locale, 'includeSelected')}</option>
        <option value="exclude">{pivotText(locale, 'excludeSelected')}</option>
      </Select>
      {(field.values ?? []).map((value) => {
        const selected = selectedMembers(field, state).some((key) => pivotMemberKeyEquals(key, keyFor(value)));
        return <CheckToggle key={pivotMemberKey(keyFor(value))} label={String(value)} checked={selected} onChange={(event) => onFilter(field.fieldId, filterWithValue(field, state, value, event.target.checked))} />;
      })}
    </Stack>
  );
}

function groupOptions(locale: Locale, field: AreaItem, onGroup?: PivotFieldAreaProps['onGroup']): ReactNode {
  if (!onGroup) return null;
  if (field.dataType === 'date') {
    return (
      <Stack gap="xs" className="border-t border-slate-100 pt-1">
        <Text size="xs" weight="semibold">{pivotText(locale, 'groupBy')}</Text>
        <Inline gap="xs" className="flex-wrap">
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'year' })}>{pivotText(locale, 'years')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'quarter' })}>{pivotText(locale, 'quarters')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'month' })}>{pivotText(locale, 'months')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'week' })}>{pivotText(locale, 'weeks')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'date', unit: 'day' })}>{pivotText(locale, 'days')}</Button>
        </Inline>
        <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>{pivotText(locale, 'clearGroup')}</Button>
      </Stack>
    );
  }
  if (field.dataType === 'number') {
    return (
      <Stack gap="xs" className="border-t border-slate-100 pt-1">
        <Text size="xs" weight="semibold">{pivotText(locale, 'groupNumbers')}</Text>
        <Inline gap="xs" className="flex-wrap">
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'number', interval: 10 })}>{pivotText(locale, 'by10')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'number', interval: 100 })}>{pivotText(locale, 'by100')}</Button>
          <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>{pivotText(locale, 'clearGroup')}</Button>
        </Inline>
      </Stack>
    );
  }
  const values = (field.values ?? []).map(keyFor);
  return values.length > 0
    ? <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'manual', groups: [{ groupId: `group:${field.fieldId}:all`, name: pivotText(locale, 'groupOne'), items: values }] })}>{pivotText(locale, 'groupAll')}</Button>
    : null;
}

const subtotalFunctions: PivotSubtotalDefinition['mode'][] = ['automatic', 'none', 'custom'];
const subtotalAggregates: PivotAggregateFunction[] = ['sum', 'average', 'count', 'count-numbers', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'];

function subtotalOptions(locale: Locale, field: AreaItem, placement: PivotFieldPlacement | undefined, onSubtotal?: PivotFieldAreaProps['onSubtotal']): ReactNode {
  if (!onSubtotal) return null;
  const current = placement?.subtotal ?? { mode: 'automatic' as const };
  const custom = current.mode === 'custom' ? current.functions : ['sum' as const];
  return (
    <Stack gap="xs" className="border-t border-slate-100 pt-1">
      <Text size="xs" weight="semibold">{pivotText(locale, 'subtotal')}</Text>
      <Select aria-label={`${field.name} ${pivotText(locale, 'subtotal')}`} sizeVariant="sm" value={current.mode} onChange={(event) => {
        const mode = event.target.value as PivotSubtotalDefinition['mode'];
        onSubtotal(field.fieldId, mode === 'custom' ? { mode, functions: custom.length ? custom : ['sum'] } : { mode });
      }}>
        {subtotalFunctions.map((mode) => <option key={mode} value={mode}>{pivotText(locale, mode === 'automatic' ? 'subtotalAutomatic' : mode === 'none' ? 'subtotalNone' : 'subtotalCustom')}</option>)}
      </Select>
      {current.mode === 'custom' ? <Inline gap="xs" className="flex-wrap">{subtotalAggregates.map((aggregate) => <CheckToggle key={aggregate} label={pivotText(locale, aggregate === 'count-numbers' ? 'countNumbers' : aggregate === 'distinct-count' ? 'distinctCount' : aggregate === 'var' ? 'variance' : aggregate === 'varp' ? 'variancep' : aggregate)} checked={custom.includes(aggregate)} onChange={(event) => {
        const next = event.target.checked ? [...custom, aggregate] : custom.filter((item) => item !== aggregate);
        onSubtotal(field.fieldId, { mode: 'custom', functions: next.length ? next : ['sum'] });
      }} />)}</Inline> : null}
    </Stack>
  );
}

export function PivotFieldArea({ area, className, disabled = false, fieldIds, fields, filterStates = {}, locale, onDrop, onFilter, onGroup, onMoveByKeyboard, onRemove, onSort, onSubtotal, onValueChange, placements, valueFields = [] }: PivotFieldAreaProps) {
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
    <Box as="section" aria-label={`${pivotText(locale, area)} field area`} className={`flex min-h-0 min-w-0 flex-1 flex-col border-[#bdbdbd] bg-white ${className ?? ''}`}>
      <Inline gap="xs" className="h-8 shrink-0 px-2">
        <Icon name={icons[area]} size="xs" className="text-accent" />
        <Text size="sm" weight="medium">{pivotText(locale, area)}</Text>
      </Inline>
      <FieldDropZone<AreaItem>
        disabled={disabled}
        emptyLabel={pivotText(locale, 'dragFieldsHere')}
        className="min-h-20 flex-1 overflow-auto rounded-none border-0 bg-white p-1"
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
            {area !== 'values' && onFilter && field.values?.length ? (
              <DropdownMenu
                align="right"
                trigger={<Button aria-label={`${pivotText(locale, 'filterValues')}: ${field.name}`} icon="filter" iconOnly size="xs" variant={filterStates[field.fieldId]?.mode && filterStates[field.fieldId]?.mode !== 'all' ? 'soft' : 'ghost'} />}
              >
                {filterOptions(locale, field, disabled, filterStates[field.fieldId] ?? { mode: 'all', memberKeys: [] }, onFilter)}
              </DropdownMenu>
            ) : null}
            <DropdownMenu
              align="right"
              trigger={<Button aria-label={`${pivotText(locale, 'fieldMenu')}: ${field.name}`} icon="more-horizontal" iconOnly size="xs" variant="ghost" />}
            >
              {({ close }) => (
                <Inline gap="xs" className="p-1">
                   <Stack gap="xs" className="min-w-48 p-1">
                     <Text size="xs" weight="semibold">{area === 'values' ? pivotText(locale, 'valueSettings') : pivotText(locale, 'fieldSettings')}</Text>
                     <Inline gap="xs">
                        <Button aria-label={pivotText(locale, 'moveUp')} disabled={field.index === 0} icon="arrow-up" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.fieldId, field.index, -1); close(); }} />
                        <Button aria-label={pivotText(locale, 'moveDown')} disabled={field.index === items.length - 1} icon="arrow-down" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.fieldId, field.index, 1); close(); }} />
                        <Button aria-label={pivotText(locale, 'remove')} icon="trash" iconOnly size="xs" variant="danger" onClick={() => { onRemove(field.fieldId, field.index); close(); }} />
                     </Inline>
                      {onSort ? <Stack gap="xs" className="border-t border-slate-100 pt-1"><Text size="xs" weight="semibold">{pivotText(locale, 'sortBy')}</Text><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'ascending', by: 'label' }); close(); }}>{pivotText(locale, 'ascending')}</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'descending', by: 'label' }); close(); }}>{pivotText(locale, 'descending')}</Button></Inline><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'ascending', by: 'value' }); close(); }}>{pivotText(locale, 'valueAscending')}</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'descending', by: 'value' }); close(); }}>{pivotText(locale, 'valueDescending')}</Button></Inline><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, undefined); close(); }}>{pivotText(locale, 'clearSort')}</Button></Stack> : null}
                      {groupOptions(locale, field, onGroup)}
                      {area !== 'values' && area !== 'filters' ? subtotalOptions(locale, field, field.placement, onSubtotal) : null}
                      {area === 'values' && onValueChange ? (() => {
                        const value = valueFields.find((entry) => entry.fieldId === field.fieldId);
                        return value ? <PivotValueEditor locale={locale} fields={fields} value={value} disabled={disabled} onChange={onValueChange} /> : null;
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
