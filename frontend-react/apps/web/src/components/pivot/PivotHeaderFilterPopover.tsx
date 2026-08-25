import { useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Icon, Inline, Panel, ScrollArea, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import {
  createPivotMemberKey,
  formatPivotMember,
  pivotMemberKey,
  pivotMemberKeyEquals,
  type PivotFieldDefinition,
  type PivotFilter,
  type PivotMemberKey,
  type PivotScalar,
  type PivotSort,
} from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotTemplate, pivotText } from './pivot-localization';

type FilterMode = 'values' | 'label' | 'value';

export interface PivotValueSortOption {
  fieldId: string;
  label: string;
}

export interface PivotHeaderFilterPopoverProps {
  currentFilter?: PivotFilter;
  currentSort?: PivotSort;
  field: PivotFieldDefinition;
  locale: Locale;
  scope?: 'report' | 'field';
  valueFields: readonly PivotValueSortOption[];
  x: number;
  y: number;
  onApply: (filter: PivotFilter | undefined, sort: PivotSort | undefined) => void;
  onClose: () => void;
}

function member(value: PivotScalar): PivotMemberKey { return createPivotMemberKey(value); }

export function PivotHeaderFilterPopover({ currentFilter, currentSort, field, locale, onApply, onClose, scope = 'field', valueFields, x, y }: PivotHeaderFilterPopoverProps) {
  const values = field.values ?? [];
  const [query, setQuery] = useState('');
  const [mode, setMode] = useState<FilterMode>(currentFilter?.kind === 'condition' ? (field.dataType === 'number' ? 'value' : 'label') : 'values');
  const [operator, setOperator] = useState<Extract<PivotFilter, { kind: 'condition' }>['operator']>(currentFilter?.kind === 'condition' ? currentFilter.operator : 'equals');
  const [conditionValue, setConditionValue] = useState<PivotScalar>(currentFilter?.kind === 'condition' ? currentFilter.value : '');
  const [sort, setSort] = useState<PivotSort | undefined>(currentSort);
  const [showSortOptions, setShowSortOptions] = useState(false);
  const initialSelected = useMemo(() => {
    if (currentFilter?.kind !== 'manual' || currentFilter.mode === 'all') return values.map(member);
    if (currentFilter.mode === 'include') return currentFilter.memberKeys;
    return values.map(member).filter((candidate) => !currentFilter.memberKeys.some((item) => pivotMemberKeyEquals(candidate, item)));
  }, [currentFilter, values]);
  const [selected, setSelected] = useState<PivotMemberKey[]>(() => [...initialSelected]);
  const visibleValues = useMemo(() => values.filter((value) => formatPivotMember(value).toLocaleLowerCase().includes(query.toLocaleLowerCase())), [query, values]);
  const selectedHas = (value: PivotScalar) => selected.some((candidate) => pivotMemberKeyEquals(candidate, member(value)));
  const setVisible = (checked: boolean) => {
    const visible = visibleValues.map(member);
    setSelected((current) => checked
      ? [...current, ...visible].filter((candidate, index, entries) => entries.findIndex((entry) => pivotMemberKeyEquals(candidate, entry)) === index)
      : current.filter((candidate) => !visible.some((entry) => pivotMemberKeyEquals(candidate, entry))));
  };
  const apply = () => {
    if (sort?.by === 'value' && !sort.valueFieldId) return;
    if (mode === 'values') {
      const filter: PivotFilter | undefined = selected.length === values.length
        ? undefined
        : { kind: 'manual', scope, fieldId: field.fieldId, mode: 'include', memberKeys: [...selected] };
      onApply(filter, sort);
    } else {
      onApply({ kind: 'condition', scope, fieldId: field.fieldId, operator, value: conditionValue }, sort);
    }
  };

  return (
    <Panel className="absolute z-40 w-[300px] rounded-none border border-[#a9a9a9] bg-white shadow-[2px_4px_10px_rgba(0,0,0,0.2)]" style={{ left: x, top: y }}>
      <Stack gap="none">
        <Button icon="sort" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setSort({ direction: 'ascending', by: 'label' })}>{pivotText(locale, 'ascending')}</Button>
        <Button icon="sort" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setSort({ direction: 'descending', by: 'label' })}>{pivotText(locale, 'descending')}</Button>
        <Button icon="sliders" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setShowSortOptions((visible) => !visible)}>{pivotText(locale, 'otherSortOptions')}</Button>
        {showSortOptions ? <Stack gap="xs" className="border-b border-[#d7d7d7] px-3 py-2"><Inline gap="xs"><Select aria-label={pivotText(locale, 'sortField')} sizeVariant="sm" value={sort?.by ?? 'label'} onChange={(event) => setSort(event.target.value === 'value' ? { direction: sort?.direction ?? 'ascending', by: 'value' } : { direction: sort?.direction ?? 'ascending', by: 'label' })}><option value="label">{pivotText(locale, 'labelFilter')}</option><option value="value" disabled={valueFields.length === 0}>{pivotText(locale, 'valueFilter')}</option></Select><Button size="xs" variant="ghost" onClick={() => setSort(undefined)}>{pivotText(locale, 'clearSort')}</Button></Inline>{sort?.by === 'value' ? <Select aria-label={pivotText(locale, 'sortField')} sizeVariant="sm" value={sort.valueFieldId ?? ''} onChange={(event) => setSort({ direction: sort.direction, by: 'value', valueFieldId: event.target.value })}><option value="" disabled>{pivotText(locale, 'sortField')}</option>{valueFields.map((value) => <option key={value.fieldId} value={value.fieldId}>{value.label}</option>)}</Select> : null}</Stack> : null}
        <Button icon="filter" disabled={!currentFilter} size="sm" variant="ghost" className="!h-9 !justify-start rounded-none border-b border-[#d7d7d7] px-4 text-[13px]" onClick={() => { setMode('values'); setSelected(values.map(member)); }}>{pivotTemplate(locale, 'clearFieldFilter', { field: field.name })}</Button>
        <Button size="sm" variant={mode === 'label' ? 'soft' : 'ghost'} className="!h-9 !justify-between rounded-none px-4 text-[13px]" onClick={() => setMode('label')}><Text as="span" size="sm">{pivotText(locale, 'labelFilter')}</Text><Icon name="chevron-right" size="xs" /></Button>
        <Button size="sm" variant={mode === 'value' ? 'soft' : 'ghost'} className="!h-9 !justify-between rounded-none border-b border-[#d7d7d7] px-4 text-[13px]" onClick={() => setMode('value')}><Text as="span" size="sm">{pivotText(locale, 'valueFilter')}</Text><Icon name="chevron-right" size="xs" /></Button>
        {mode !== 'values' ? (
          <Stack gap="xs" className="border-b border-[#d7d7d7] p-3">
            <Select aria-label={pivotText(locale, 'filterMode')} sizeVariant="sm" value={operator} onChange={(event) => setOperator(event.target.value as typeof operator)}>
              <option value="equals">=</option><option value="not-equals">≠</option><option value="contains">{locale === 'zh-CN' ? '包含' : 'Contains'}</option><option value="greater-than">&gt;</option><option value="greater-or-equal">≥</option><option value="less-than">&lt;</option><option value="less-or-equal">≤</option>
            </Select>
            <TextInput aria-label={pivotText(locale, 'filterValues')} value={String(conditionValue ?? '')} onChange={(event) => setConditionValue(field.dataType === 'number' && event.target.value !== '' ? Number(event.target.value) : event.target.value)} />
          </Stack>
        ) : (
          <Stack gap="xs" className="p-3">
            <TextInput aria-label={pivotText(locale, 'search')} className="!h-8 rounded-none border-[#b8b8b8] text-[13px]" placeholder={pivotText(locale, 'search')} value={query} onChange={(event) => setQuery(event.target.value)} />
            <Inline gap="sm" className="h-6">
              <Button icon="check" size="xs" variant="ghost" className="!h-6 px-1 text-[12px]" onClick={() => setVisible(true)}>{pivotText(locale, 'selectAll')}</Button>
              <Button icon="x" size="xs" variant="ghost" className="!h-6 px-1 text-[12px]" onClick={() => setVisible(false)}>{locale === 'zh-CN' ? '取消全选' : 'Deselect all'}</Button>
            </Inline>
            <ScrollArea className="h-[210px] border border-[#c8c8c8] p-2">
              <Stack gap="none">
                {visibleValues.map((value) => <Box key={pivotMemberKey(member(value))} className="py-1"><CheckToggle checkedTone="dark" className="text-[13px]" label={formatPivotMember(value)} checked={selectedHas(value)} onChange={(event) => setSelected((current) => event.target.checked ? [...current, member(value)] : current.filter((candidate) => !pivotMemberKeyEquals(candidate, member(value))))} /></Box>)}
              </Stack>
            </ScrollArea>
          </Stack>
        )}
        <Inline gap="sm" className="justify-end border-t border-[#d7d7d7] p-2">
          <Button size="sm" variant="outline" disabled={sort?.by === 'value' && !sort.valueFieldId} className="min-w-[86px] rounded-none border-[#0078d4] text-slate-900" onClick={apply}>{pivotText(locale, 'confirm')}</Button>
          <Button size="sm" variant="outline" className="min-w-[86px] rounded-none" onClick={onClose}>{pivotText(locale, 'cancel')}</Button>
        </Inline>
      </Stack>
    </Panel>
  );
}
