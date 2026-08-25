import { useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Icon, Inline, Panel, ScrollArea, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import {
  createPivotMemberKey,
  formatPivotMember,
  pivotMemberKey,
  pivotMemberKeyEquals,
  type PivotFieldDefinition,
  type PivotDateFilterOperator,
  type PivotFilter,
  type PivotFilterFamily,
  type PivotLabelFilterOperator,
  type PivotMemberKey,
  type PivotScalar,
  type PivotSort,
} from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotTemplate, pivotText } from './pivot-localization';

type FilterMode = 'values' | 'label' | 'date' | 'value';

export interface PivotValueSortOption {
  fieldId: string;
  label: string;
}

export interface PivotFilterMemberOption {
  key: PivotMemberKey;
  value: PivotScalar;
  label: string;
}

export interface PivotHeaderFilterPopoverProps {
  currentFilters: readonly PivotFilter[];
  currentSort?: PivotSort;
  field: PivotFieldDefinition;
  memberOptions?: readonly PivotFilterMemberOption[];
  locale: Locale;
  scope?: 'report' | 'field';
  valueFields: readonly PivotValueSortOption[];
  x: number;
  y: number;
  onApply: (filter: PivotFilter | undefined, sort: PivotSort | undefined, family: PivotFilterFamily | 'all') => void;
  onClose: () => void;
}

function member(value: PivotScalar): PivotMemberKey { return createPivotMemberKey(value); }

export function PivotHeaderFilterPopover({ currentFilters, currentSort, field, locale, memberOptions, onApply, onClose, scope = 'field', valueFields, x, y }: PivotHeaderFilterPopoverProps) {
  const values = field.values ?? [];
  const options = useMemo(() => memberOptions ?? values.map((value) => ({ key: member(value), value, label: formatPivotMember(value) })), [memberOptions, values]);
  const manualFilter = currentFilters.find((filter) => filter.kind === 'manual' && filter.family === 'manual');
  const conditionFilter = currentFilters.find((filter) => filter.kind === 'condition');
  const [query, setQuery] = useState('');
  const [mode, setMode] = useState<FilterMode>(conditionFilter?.family === 'value' || field.dataType === 'number' && conditionFilter ? 'value' : conditionFilter?.family === 'date' || field.dataType === 'date' && conditionFilter ? 'date' : conditionFilter ? 'label' : 'values');
  const [operator, setOperator] = useState<Extract<PivotFilter, { kind: 'condition' }>['operator']>(conditionFilter?.kind === 'condition' ? conditionFilter.operator : 'equals');
  const [conditionValue, setConditionValue] = useState<PivotScalar>(conditionFilter?.kind === 'condition' ? conditionFilter.value : '');
  const [conditionValue2, setConditionValue2] = useState<PivotScalar>(conditionFilter?.kind === 'condition' ? conditionFilter.value2 ?? '' : '');
  const [dynamicDate, setDynamicDate] = useState<Extract<PivotFilter, { kind: 'condition' }>['dynamic']>(conditionFilter?.kind === 'condition' ? conditionFilter.dynamic : undefined);
  const [sort, setSort] = useState<PivotSort | undefined>(currentSort);
  const [showSortOptions, setShowSortOptions] = useState(false);
  const initialSelected = useMemo(() => {
    if (!manualFilter || manualFilter.mode === 'all') return options.map((option) => option.key);
    if (manualFilter.mode === 'include') return manualFilter.memberKeys;
    return options.map((option) => option.key).filter((candidate) => !manualFilter.memberKeys.some((item) => pivotMemberKeyEquals(candidate, item)));
  }, [manualFilter, options]);
  const [selected, setSelected] = useState<PivotMemberKey[]>(() => [...initialSelected]);
  const visibleValues = useMemo(() => options.filter((option) => option.label.toLocaleLowerCase().includes(query.toLocaleLowerCase())), [query, options]);
  const selectedHas = (option: PivotFilterMemberOption) => selected.some((candidate) => pivotMemberKeyEquals(candidate, option.key));
  const setVisible = (checked: boolean) => {
    const visible = visibleValues.map((option) => option.key);
    setSelected((current) => checked
      ? [...current, ...visible].filter((candidate, index, entries) => entries.findIndex((entry) => pivotMemberKeyEquals(candidate, entry)) === index)
      : current.filter((candidate) => !visible.some((entry) => pivotMemberKeyEquals(candidate, entry))));
  };
  const apply = () => {
    if (sort?.by === 'value' && !sort.valueFieldId) return;
    if (mode === 'values') {
      const filter: PivotFilter | undefined = selected.length === options.length
        ? undefined
        : { kind: 'manual', family: 'manual', scope, fieldId: field.fieldId, mode: 'include', memberKeys: [...selected] };
      onApply(filter, sort, 'manual');
    } else {
      const family = mode === 'value' ? 'value' : mode === 'date' ? 'date' : 'label';
      const common = { kind: 'condition' as const, scope, fieldId: field.fieldId, value: conditionValue, ...((operator === 'between' || operator === 'not-between') && dynamicDate === undefined ? { value2: conditionValue2 } : {}), ...(family === 'date' && dynamicDate ? { dynamic: dynamicDate } : {}) };
      if (family === 'date') onApply({ ...common, family, operator: operator as PivotDateFilterOperator }, sort, family);
      else if (family === 'label') onApply({ ...common, family, operator: operator as PivotLabelFilterOperator }, sort, family);
      else onApply({ ...common, family, operator: operator as Extract<PivotFilter, { kind: 'condition'; family: 'value' }>['operator'] }, sort, family);
    }
  };

  return (
    <Panel className="absolute z-40 w-[300px] rounded-none border border-[#a9a9a9] bg-white shadow-[2px_4px_10px_rgba(0,0,0,0.2)]" style={{ left: x, top: y }}>
      <Stack gap="none">
        <Button icon="sort" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setSort({ direction: 'ascending', by: 'label' })}>{pivotText(locale, 'ascending')}</Button>
        <Button icon="sort" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setSort({ direction: 'descending', by: 'label' })}>{pivotText(locale, 'descending')}</Button>
        <Button icon="sliders" size="sm" variant="ghost" className="!h-9 !justify-start rounded-none px-4 text-[13px]" onClick={() => setShowSortOptions((visible) => !visible)}>{pivotText(locale, 'otherSortOptions')}</Button>
        {showSortOptions ? <Stack gap="xs" className="border-b border-[#d7d7d7] px-3 py-2"><Inline gap="xs"><Select aria-label={pivotText(locale, 'sortField')} sizeVariant="sm" value={sort?.by ?? 'label'} onChange={(event) => setSort(event.target.value === 'value' ? { direction: sort?.direction ?? 'ascending', by: 'value' } : { direction: sort?.direction ?? 'ascending', by: 'label' })}><option value="label">{pivotText(locale, 'labelFilter')}</option><option value="value" disabled={valueFields.length === 0}>{pivotText(locale, 'valueFilter')}</option></Select><Button size="xs" variant="ghost" onClick={() => setSort(undefined)}>{pivotText(locale, 'clearSort')}</Button></Inline>{sort?.by === 'value' ? <Select aria-label={pivotText(locale, 'sortField')} sizeVariant="sm" value={sort.valueFieldId ?? ''} onChange={(event) => setSort({ direction: sort.direction, by: 'value', valueFieldId: event.target.value })}><option value="" disabled>{pivotText(locale, 'sortField')}</option>{valueFields.map((value) => <option key={value.fieldId} value={value.fieldId}>{value.label}</option>)}</Select> : null}</Stack> : null}
        <Button icon="filter" disabled={currentFilters.length === 0} size="sm" variant="ghost" className="!h-9 !justify-start rounded-none border-b border-[#d7d7d7] px-4 text-[13px]" onClick={() => onApply(undefined, sort, 'all')}>{pivotTemplate(locale, 'clearFieldFilter', { field: field.name })}</Button>
        {field.dataType === 'date' ? <Button size="sm" variant={mode === 'date' ? 'soft' : 'ghost'} className="!h-9 !justify-between rounded-none px-4 text-[13px]" onClick={() => setMode('date')}><Text as="span" size="sm">{pivotText(locale, 'dateFilter')}</Text><Icon name="chevron-right" size="xs" /></Button> : <Button size="sm" variant={mode === 'label' ? 'soft' : 'ghost'} className="!h-9 !justify-between rounded-none px-4 text-[13px]" onClick={() => setMode('label')}><Text as="span" size="sm">{pivotText(locale, 'labelFilter')}</Text><Icon name="chevron-right" size="xs" /></Button>}
        {field.dataType === 'number' ? <Button size="sm" variant={mode === 'value' ? 'soft' : 'ghost'} className="!h-9 !justify-between rounded-none border-b border-[#d7d7d7] px-4 text-[13px]" onClick={() => setMode('value')}><Text as="span" size="sm">{pivotText(locale, 'valueFilter')}</Text><Icon name="chevron-right" size="xs" /></Button> : null}
        {mode !== 'values' ? (
          <Stack gap="xs" className="border-b border-[#d7d7d7] p-3">
            <Select aria-label={pivotText(locale, 'filterMode')} sizeVariant="sm" value={operator} onChange={(event) => { const next = event.target.value as typeof operator; setOperator(next); if (mode === 'date' && next !== 'equals' && next !== 'between') setDynamicDate(undefined); }}>
              {mode === 'label' ? <><option value="equals">=</option><option value="not-equals">≠</option><option value="begins-with">{pivotText(locale, 'beginsWith')}</option><option value="not-begins-with">{pivotText(locale, 'notBeginsWith')}</option><option value="ends-with">{pivotText(locale, 'endsWith')}</option><option value="not-ends-with">{pivotText(locale, 'notEndsWith')}</option><option value="contains">{locale === 'zh-CN' ? '包含' : 'Contains'}</option><option value="not-contains">{pivotText(locale, 'notContains')}</option><option value="between">{pivotText(locale, 'between')}</option><option value="not-between">{pivotText(locale, 'notBetween')}</option><option value="greater-than">&gt;</option><option value="greater-or-equal">≥</option><option value="less-than">&lt;</option><option value="less-or-equal">≤</option></> : mode === 'date' ? <><option value="equals">=</option><option value="not-equals">≠</option><option value="before">{locale === 'zh-CN' ? '之前' : 'Before'}</option><option value="after">{locale === 'zh-CN' ? '之后' : 'After'}</option><option value="between">{pivotText(locale, 'between')}</option><option value="not-between">{pivotText(locale, 'notBetween')}</option></> : <><option value="equals">=</option><option value="not-equals">≠</option><option value="greater-than">&gt;</option><option value="greater-or-equal">≥</option><option value="less-than">&lt;</option><option value="less-or-equal">≤</option><option value="between">{pivotText(locale, 'between')}</option><option value="not-between">{pivotText(locale, 'notBetween')}</option></>}
            </Select>
            {mode === 'date' ? <Select aria-label={pivotText(locale, 'dynamicDate')} sizeVariant="sm" value={dynamicDate ?? ''} onChange={(event) => setDynamicDate((event.target.value || undefined) as typeof dynamicDate)}><option value="">{pivotText(locale, 'between')}</option>{(['today', 'yesterday', 'tomorrow', 'this-week', 'last-week', 'next-week', 'this-month', 'last-month', 'next-month', 'this-quarter', 'last-quarter', 'next-quarter', 'this-year', 'last-year', 'next-year', 'year-to-date'] as const).map((item) => <option key={item} value={item}>{pivotText(locale, item === 'this-week' ? 'thisWeek' : item === 'last-week' ? 'lastWeek' : item === 'next-week' ? 'nextWeek' : item === 'this-month' ? 'thisMonth' : item === 'last-month' ? 'lastMonth' : item === 'next-month' ? 'nextMonth' : item === 'this-quarter' ? 'thisQuarter' : item === 'last-quarter' ? 'lastQuarter' : item === 'next-quarter' ? 'nextQuarter' : item === 'this-year' ? 'thisYear' : item === 'last-year' ? 'lastYear' : item === 'next-year' ? 'nextYear' : item === 'year-to-date' ? 'yearToDate' : item)}</option>)}</Select> : null}
            <TextInput type={mode === 'date' ? 'date' : field.dataType === 'number' ? 'number' : 'text'} aria-label={pivotText(locale, 'filterValues')} value={String(conditionValue ?? '')} disabled={mode === 'date' && dynamicDate !== undefined} onChange={(event) => setConditionValue(field.dataType === 'number' && event.target.value !== '' ? Number(event.target.value) : event.target.value)} />
            {(operator === 'between' || operator === 'not-between') && dynamicDate === undefined ? <TextInput type={mode === 'date' ? 'date' : field.dataType === 'number' ? 'number' : 'text'} aria-label={pivotText(locale, 'between')} value={String(conditionValue2 ?? '')} onChange={(event) => setConditionValue2(field.dataType === 'number' && event.target.value !== '' ? Number(event.target.value) : event.target.value)} /> : null}
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
                {visibleValues.map((option) => <Box key={pivotMemberKey(option.key)} className="py-1"><CheckToggle checkedTone="dark" className="text-[13px]" label={option.label} checked={selectedHas(option)} onChange={(event) => setSelected((current) => event.target.checked ? [...current, option.key] : current.filter((candidate) => !pivotMemberKeyEquals(candidate, option.key)))} /></Box>)}
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
