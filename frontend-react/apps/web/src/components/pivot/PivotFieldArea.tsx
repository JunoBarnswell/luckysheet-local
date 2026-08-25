import { Box, Button, CheckToggle, DropdownMenu, FieldDropZone, Icon, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DragEvent, ReactNode } from 'react';
import { useState } from 'react';
import {
  createPivotMemberKey,
  formatPivotMember,
  PIVOT_MEMBER_DISPLAY_LIMIT,
  type PivotAggregateFunction,
  pivotMemberKey,
  pivotMemberKeyEquals,
  type PivotFieldDefinition,
  type PivotFieldPlacement,
  type PivotGroup,
  type PivotDateGroupUnit,
  type PivotMemberKey,
  type PivotSort,
  type PivotScalar,
  type PivotSubtotalDefinition,
  type PivotValueField,
} from '@react-sheets/core-model';
import type { PivotFieldArea as Area, PivotManualFilterState } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText, type PivotMessageKey } from './pivot-localization';
import { PivotValueEditor } from './PivotValueEditor';
import { buildPivotGroupedFilterMembers, type PivotGroupedFilterMember } from '@react-sheets/spreadsheet-app';
import { applyPivotManualMemberDelta, convertPivotManualFilterMode, pivotManualMemberSelected } from './pivot-member-filter';

interface AreaItem extends PivotFieldDefinition {
  id: string;
  fieldId: string;
  index: number;
  placement?: PivotFieldPlacement;
  groupedMembers?: readonly PivotGroupedFilterMember[];
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

function filterMembers(field: AreaItem): readonly PivotGroupedFilterMember[] {
  return field.groupedMembers ?? (field.values ?? []).map((value) => ({ key: keyFor(value), value, label: formatPivotMember(value) }));
}

function filterWithValue(_field: AreaItem, state: PivotManualFilterState, target: PivotMemberKey, checked: boolean): PivotManualFilterState {
  return applyPivotManualMemberDelta(state, [target], checked);
}

function filterForMode(field: AreaItem, state: PivotManualFilterState, mode: PivotManualFilterState['mode']): PivotManualFilterState {
  return convertPivotManualFilterMode(state, mode, filterMembers(field).map((item) => item.key));
}

function FilterOptions({ locale, field, disabled, state, onFilter }: { locale: Locale; field: AreaItem; disabled: boolean; state: PivotManualFilterState; onFilter: NonNullable<PivotFieldAreaProps['onFilter']> }): ReactNode {
  const [query, setQuery] = useState('');
  const members = filterMembers(field);
  const visibleMembers = members.filter((item) => item.label.toLocaleLowerCase().includes(query.toLocaleLowerCase())).slice(0, PIVOT_MEMBER_DISPLAY_LIMIT);
  return (
    <Stack gap="xs" className="max-h-[18rem] min-w-[13rem] overflow-auto p-2">
      <Text size="xs" weight="semibold">{pivotText(locale, 'filterValues')}</Text>
      <Select aria-label={`${field.name} ${pivotText(locale, 'filterMode')}`} sizeVariant="sm" disabled={disabled} value={state.mode} onChange={(event) => onFilter(field.fieldId, filterForMode(field, state, event.target.value as PivotManualFilterState['mode']))}>
        <option value="all">{pivotText(locale, 'showAll')}</option>
        <option value="include">{pivotText(locale, 'includeSelected')}</option>
        <option value="exclude">{pivotText(locale, 'excludeSelected')}</option>
      </Select>
      <TextInput aria-label={`${field.name} ${pivotText(locale, 'search')}`} placeholder={pivotText(locale, 'search')} value={query} onChange={(event) => setQuery(event.target.value)} />
      {visibleMembers.map((item) => {
        const selected = pivotManualMemberSelected(state, item.key);
        return <CheckToggle key={pivotMemberKey(item.key)} label={item.label} checked={selected} onChange={(event) => onFilter(field.fieldId, filterWithValue(field, state, item.key, event.target.checked))} />;
      })}
    </Stack>
  );
}

const dateUnits: readonly PivotDateGroupUnit[] = ['year', 'quarter', 'month', 'week', 'day'];
const dateUnitKeys: Record<PivotDateGroupUnit, PivotMessageKey> = { year: 'years', quarter: 'quarters', month: 'months', week: 'weeks', day: 'days' };

function GroupingOptions({ field, locale, onGroup }: { field: AreaItem; locale: Locale; onGroup: NonNullable<PivotFieldAreaProps['onGroup']> }): ReactNode {
  const current = field.placement?.group;
  if (field.dataType === 'date') {
    const date = current?.kind === 'date' ? current : undefined;
    const [units, setUnits] = useState<PivotDateGroupUnit[]>(date?.units?.length ? [...date.units] : date ? [date.unit] : ['year']);
    const [start, setStart] = useState(date?.start === undefined ? '' : String(date.start));
    const [end, setEnd] = useState(date?.end === undefined ? '' : String(date.end));
    const [startOfWeek, setStartOfWeek] = useState(String(date?.startOfWeek ?? 0));
    const apply = () => {
      const nextUnits = dateUnits.filter((unit) => units.includes(unit));
      if (!nextUnits.length) return;
      onGroup(field.fieldId, { kind: 'date', unit: nextUnits[0]!, units: nextUnits, ...(start === '' ? {} : { start: Number.isNaN(Number(start)) ? start : Number(start) }), ...(end === '' ? {} : { end: Number.isNaN(Number(end)) ? end : Number(end) }), startOfWeek: Number(startOfWeek) as 0 | 1 | 2 | 3 | 4 | 5 | 6 });
    };
    return <Stack gap="xs" className="border-t border-slate-100 pt-1">
      <Text size="xs" weight="semibold">{pivotText(locale, 'groupBy')}</Text>
      <Inline gap="xs" className="flex-wrap">{dateUnits.map((unit) => <CheckToggle key={unit} label={pivotText(locale, dateUnitKeys[unit])} checked={units.includes(unit)} onChange={(event) => setUnits((currentUnits) => event.target.checked ? [...currentUnits, unit] : currentUnits.filter((candidate) => candidate !== unit))} />)}</Inline>
      <Inline gap="xs"><TextInput aria-label={pivotText(locale, 'groupStart')} placeholder={pivotText(locale, 'groupStart')} value={start} onChange={(event) => setStart(event.target.value)} /><TextInput aria-label={pivotText(locale, 'groupEnd')} placeholder={pivotText(locale, 'groupEnd')} value={end} onChange={(event) => setEnd(event.target.value)} /></Inline>
      <Select aria-label={pivotText(locale, 'weekStarts')} sizeVariant="sm" value={startOfWeek} onChange={(event) => setStartOfWeek(event.target.value)}><option value="0">{pivotText(locale, 'sunday')}</option><option value="1">{pivotText(locale, 'monday')}</option><option value="6">{pivotText(locale, 'saturday')}</option></Select>
      <Inline gap="xs"><Button size="xs" variant="soft" disabled={units.length === 0} onClick={apply}>{pivotText(locale, 'applyGroup')}</Button><Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>{pivotText(locale, 'clearGroup')}</Button></Inline>
    </Stack>;
  }
  if (field.dataType === 'number') {
    const numeric = current?.kind === 'number' ? current : undefined;
    const [interval, setInterval] = useState(numeric ? String(numeric.interval) : '10');
    const [start, setStart] = useState(numeric?.start === undefined ? '' : String(numeric.start));
    const [end, setEnd] = useState(numeric?.end === undefined ? '' : String(numeric.end));
    const apply = () => {
      const parsedInterval = Number(interval);
      const parsedStart = start === '' ? undefined : Number(start);
      const parsedEnd = end === '' ? undefined : Number(end);
      if (!Number.isFinite(parsedInterval) || parsedInterval <= 0 || (parsedStart !== undefined && !Number.isFinite(parsedStart)) || (parsedEnd !== undefined && !Number.isFinite(parsedEnd))) return;
      onGroup(field.fieldId, { kind: 'number', interval: parsedInterval, ...(parsedStart === undefined ? {} : { start: parsedStart }), ...(parsedEnd === undefined ? {} : { end: parsedEnd }) });
    };
    return <Stack gap="xs" className="border-t border-slate-100 pt-1">
      <Text size="xs" weight="semibold">{pivotText(locale, 'groupNumbers')}</Text>
      <Inline gap="xs"><TextInput type="number" aria-label={pivotText(locale, 'groupStart')} placeholder={pivotText(locale, 'groupStart')} value={start} onChange={(event) => setStart(event.target.value)} /><TextInput type="number" aria-label={pivotText(locale, 'groupEnd')} placeholder={pivotText(locale, 'groupEnd')} value={end} onChange={(event) => setEnd(event.target.value)} /></Inline>
      <TextInput type="number" aria-label={pivotText(locale, 'groupInterval')} placeholder={pivotText(locale, 'groupInterval')} value={interval} onChange={(event) => setInterval(event.target.value)} />
      <Inline gap="xs"><Button size="xs" variant="soft" onClick={apply}>{pivotText(locale, 'applyGroup')}</Button><Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>{pivotText(locale, 'clearGroup')}</Button></Inline>
    </Stack>;
  }
  const allValues = field.values ?? [];
  const [query, setQuery] = useState('');
  const values = allValues.filter((item) => formatPivotMember(item).toLocaleLowerCase().includes(query.toLocaleLowerCase())).slice(0, PIVOT_MEMBER_DISPLAY_LIMIT).map(keyFor);
  const manual = current?.kind === 'manual' ? current : undefined;
  const [selected, setSelected] = useState<PivotMemberKey[]>([]);
  const [name, setName] = useState(pivotText(locale, 'groupOne'));
  if (!values.length) return null;
  const create = () => {
    if (selected.length < 2) return;
    const selectedSet = new Set(selected.map(pivotMemberKey));
    const retained = (manual?.groups ?? []).map((group) => ({ ...group, items: group.items.filter((item) => !selectedSet.has(pivotMemberKey(item))) })).filter((group) => group.items.length > 0);
    const groupId = `group:${field.fieldId}:${Date.now().toString(36)}`;
    onGroup(field.fieldId, { kind: 'manual', groups: [...retained, { groupId, name: name.trim() || pivotText(locale, 'groupOne'), items: selected }] });
    setSelected([]);
  };
  return <Stack gap="xs" className="border-t border-slate-100 pt-1">
      <Text size="xs" weight="semibold">{pivotText(locale, 'manualGrouping')}</Text>
      <Text size="xs" tone="muted">{pivotText(locale, 'selectItemsToGroup')}</Text>
      <TextInput aria-label={`${field.name} ${pivotText(locale, 'search')}`} placeholder={pivotText(locale, 'search')} value={query} onChange={(event) => setQuery(event.target.value)} />
      <Stack gap="xs" className="max-h-32 overflow-auto">{values.map((item) => <CheckToggle key={pivotMemberKey(item)} label={String(item.value ?? '')} checked={selected.some((candidate) => pivotMemberKeyEquals(candidate, item))} onChange={(event) => setSelected((currentItems) => event.target.checked ? [...currentItems, item] : currentItems.filter((candidate) => !pivotMemberKeyEquals(candidate, item)))} />)}</Stack>
    <TextInput aria-label={pivotText(locale, 'groupName')} value={name} onChange={(event) => setName(event.target.value)} />
    <Button size="xs" variant="soft" disabled={selected.length < 2} onClick={create}>{pivotText(locale, 'createGroup')}</Button>
    {manual?.groups.map((group) => <Inline key={group.groupId} gap="xs"><Text size="xs" className="min-w-0 flex-1 truncate">{group.name}</Text><Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, { kind: 'manual', groups: manual.groups.filter((candidate) => candidate.groupId !== group.groupId) })}>{pivotText(locale, 'ungroup')}</Button></Inline>)}
    {manual ? <Button size="xs" variant="ghost" onClick={() => onGroup(field.fieldId, undefined)}>{pivotText(locale, 'clearGroup')}</Button> : null}
  </Stack>;
}

function groupOptions(locale: Locale, field: AreaItem, onGroup?: PivotFieldAreaProps['onGroup']): ReactNode {
  return onGroup ? <GroupingOptions field={field} locale={locale} onGroup={onGroup} /> : null;
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
  const [valueSortFieldIds, setValueSortFieldIds] = useState<Record<string, string>>({});
  const items: AreaItem[] = fieldIds.map((fieldId, index) => {
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    return {
      ...(field ?? { name: fieldId, dataType: 'text' as const, ordinal: index }),
      id: fieldId,
      fieldId,
      index,
      placement: placements?.get(fieldId),
      groupedMembers: placements?.get(fieldId)?.group ? buildPivotGroupedFilterMembers(field?.values ?? [], placements.get(fieldId)!.group!) : undefined,
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
                <FilterOptions locale={locale} field={field} disabled={disabled} state={filterStates[field.fieldId] ?? { mode: 'all', memberKeys: [] }} onFilter={onFilter} />
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
                      {onSort ? <Stack gap="xs" className="border-t border-slate-100 pt-1"><Text size="xs" weight="semibold">{pivotText(locale, 'sortBy')}</Text><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'ascending', by: 'label' }); close(); }}>{pivotText(locale, 'ascending')}</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, { direction: 'descending', by: 'label' }); close(); }}>{pivotText(locale, 'descending')}</Button></Inline><Select aria-label={`${field.name} ${pivotText(locale, 'sortField')}`} sizeVariant="sm" disabled={disabled || valueFields.length === 0} value={valueSortFieldIds[field.fieldId] ?? field.placement?.sort?.valueFieldId ?? ''} onChange={(event) => setValueSortFieldIds((current) => ({ ...current, [field.fieldId]: event.target.value }))}><option value="" disabled>{pivotText(locale, 'sortField')}</option>{valueFields.map((value) => <option key={value.fieldId} value={value.fieldId}>{value.displayName ?? value.fieldId}</option>)}</Select><Inline gap="xs"><Button size="xs" variant="ghost" disabled={!valueSortFieldIds[field.fieldId] && !field.placement?.sort?.valueFieldId} onClick={() => { const valueFieldId = valueSortFieldIds[field.fieldId] ?? field.placement?.sort?.valueFieldId; if (valueFieldId) { onSort(field.fieldId, { direction: 'ascending', by: 'value', valueFieldId }); close(); } }}>{pivotText(locale, 'valueAscending')}</Button><Button size="xs" variant="ghost" disabled={!valueSortFieldIds[field.fieldId] && !field.placement?.sort?.valueFieldId} onClick={() => { const valueFieldId = valueSortFieldIds[field.fieldId] ?? field.placement?.sort?.valueFieldId; if (valueFieldId) { onSort(field.fieldId, { direction: 'descending', by: 'value', valueFieldId }); close(); } }}>{pivotText(locale, 'valueDescending')}</Button></Inline><Button size="xs" variant="ghost" onClick={() => { onSort(field.fieldId, undefined); close(); }}>{pivotText(locale, 'clearSort')}</Button></Stack> : null}
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
