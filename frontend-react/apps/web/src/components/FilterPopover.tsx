import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Select, Stack, TextInput, Text, VirtualList } from '@react-sheets/ui-system';
import type { DateGroupItem, FilterCriterion, FilterScalar } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import type { FilterDomainDescriptor, FilterFamily } from '@react-sheets/sheet-features';
import { filterText, type FilterUiTextKey, type Locale } from '../i18n';

export interface FilterPatch { criterion?: FilterCriterion; }

export interface FilterPopoverProps {
  locale: Locale;
  column: number;
  x: number;
  y: number;
  sheet: CanvasSheetSnapshot;
  onApply: (patch: FilterPatch) => void;
  onSort: (ascending: boolean) => void;
  onClose: () => void;
}

type FilterMode = FilterFamily;
type CustomOperator = 'equals' | 'notEquals' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'contains' | 'notContains' | 'beginsWith' | 'endsWith';
type FilterJoin = 'and' | 'or';
type DateMode = 'condition' | 'dynamic';
type DynamicType = 'today' | 'yesterday' | 'tomorrow' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth' | 'thisQuarter' | 'lastQuarter' | 'nextQuarter' | 'thisYear' | 'lastYear' | 'nextYear' | 'yearToDate';

const FILTER_MODES: readonly FilterMode[] = ['values', 'text', 'number', 'date', 'color', 'icon'];
const TEXT_OPERATORS: readonly CustomOperator[] = ['equals', 'notEquals', 'contains', 'notContains', 'beginsWith', 'endsWith'];
const ORDERED_OPERATORS: readonly CustomOperator[] = ['equals', 'notEquals', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual'];
const DYNAMIC_TYPES: readonly DynamicType[] = ['today', 'yesterday', 'tomorrow', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth', 'thisQuarter', 'lastQuarter', 'nextQuarter', 'thisYear', 'lastYear', 'nextYear', 'yearToDate'];
const modeTextKeys: Record<FilterMode, FilterUiTextKey> = { values: 'values', text: 'text', number: 'number', date: 'date', color: 'color', icon: 'icon' };
const operatorTextKeys: Record<CustomOperator, FilterUiTextKey> = { equals: 'equals', notEquals: 'notEquals', lessThan: 'lessThan', lessThanOrEqual: 'lessThanOrEqual', greaterThan: 'greaterThan', greaterThanOrEqual: 'greaterThanOrEqual', contains: 'contains', notContains: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' };

export interface FilterValueOption { key: string; label: string; value: FilterScalar; }

export function filterScalarKey(value: FilterScalar): string { return JSON.stringify(value); }
export function filterScalarLabel(value: FilterScalar): string { return value == null || value === '' ? '' : String(value); }
export function filterModeOptions(descriptor: FilterDomainDescriptor): FilterMode[] { return FILTER_MODES.filter((mode) => descriptor.supportedFamilies.includes(mode)); }
export function filterOperatorsFor(mode: FilterMode): readonly CustomOperator[] {
  if (mode === 'text') return TEXT_OPERATORS;
  if (mode === 'number' || mode === 'date') return ORDERED_OPERATORS;
  return [];
}

function currentMode(descriptor: FilterDomainDescriptor, criterion: FilterCriterion | undefined): FilterMode {
  if (descriptor.currentFamily && descriptor.supportedFamilies.includes(descriptor.currentFamily)) return descriptor.currentFamily;
  const first = filterModeOptions(descriptor)[0];
  if (!first) throw new Error('Filter domain has no supported family');
  return first;
}

function currentCriterionValues(criterion: FilterCriterion | undefined): Set<string> {
  return new Set(criterion?.kind === 'values' ? criterion.values.map(filterScalarKey) : []);
}

type DateGroupUnit = 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second';
type DateGroupNode = { key: string; group: DateGroupItem; unit: DateGroupUnit; depth: number; label: string };
const dateGroupUnits: readonly DateGroupUnit[] = ['year', 'month', 'day', 'hour', 'minute', 'second'];

function dateGroupKey(group: DateGroupItem): string { return dateGroupUnits.filter((unit) => group[unit] !== undefined).map((unit) => `${unit}=${group[unit]}`).join('|'); }
function dateGroupLabel(group: DateGroupItem, unit: DateGroupUnit): string {
  if (unit === 'year') return `${group.year}`;
  if (unit === 'month') return `${group.year}-${String(group.month).padStart(2, '0')}`;
  if (unit === 'day') return `${group.year}-${String(group.month).padStart(2, '0')}-${String(group.day).padStart(2, '0')}`;
  if (unit === 'hour') return `${dateGroupLabel(group, 'day')} ${String(group.hour).padStart(2, '0')}:00`;
  if (unit === 'minute') return `${dateGroupLabel(group, 'hour')}:${String(group.minute).padStart(2, '0')}`;
  return `${dateGroupLabel(group, 'minute')}:${String(group.second).padStart(2, '0')}`;
}
function buildDateGroupNodes(entries: Array<{ group: DateGroupItem }>, active: DateGroupItem[]): DateGroupNode[] {
  const groups = new Map<string, DateGroupItem>();
  const add = (source: DateGroupItem): void => {
    for (let index = 0; index < dateGroupUnits.length; index += 1) {
      const unit = dateGroupUnits[index]!;
      if (unit !== 'year' && source[dateGroupUnits[index - 1]!] === undefined) break;
      if (source[unit] === undefined && unit !== 'year') break;
      const group = Object.fromEntries(dateGroupUnits.slice(0, index + 1).filter((key) => source[key] !== undefined).map((key) => [key, source[key]])) as unknown as DateGroupItem;
      groups.set(dateGroupKey(group), group);
    }
  };
  entries.forEach((entry) => add(entry.group));
  active.forEach(add);
  return [...groups.entries()].map(([key, group]) => {
    const unit = [...dateGroupUnits].reverse().find((candidate) => group[candidate] !== undefined) ?? 'year';
    return { key, group, unit, depth: dateGroupUnits.indexOf(unit), label: dateGroupLabel(group, unit) };
  }).sort((left, right) => left.key.localeCompare(right.key));
}

function dynamicLabel(locale: Locale, type: DynamicType): string { return filterText(locale, type); }
export interface FilterPopoverViewport { width: number; height: number; }
export interface FilterPopoverSize { width: number; height: number; }
export function clampFilterPopoverSize(size: FilterPopoverSize, viewport: FilterPopoverViewport): FilterPopoverSize {
  const maxWidth = Math.max(280, viewport.width - 16);
  const maxHeight = Math.max(360, viewport.height - 16);
  return { width: Math.min(maxWidth, Math.max(280, size.width)), height: Math.min(maxHeight, Math.max(360, size.height)) };
}
export function clampFilterPopoverPosition(x: number, y: number, size: FilterPopoverSize, viewport: FilterPopoverViewport): { left: number; top: number } {
  return { left: Math.min(Math.max(4, x), Math.max(4, viewport.width - size.width - 4)), top: Math.min(Math.max(4, y), Math.max(4, viewport.height - size.height - 4)) };
}
function viewportBounds(width: number, height: number): { maxLeft: number; maxTop: number } {
  if (typeof window === 'undefined') return { maxLeft: Number.POSITIVE_INFINITY, maxTop: Number.POSITIVE_INFINITY };
  return { maxLeft: Math.max(4, window.innerWidth - width - 4), maxTop: Math.max(4, window.innerHeight - height - 4) };
}

/** Excel-style filter task surface. Draft and resize state remain transient; only Apply emits a canonical command payload. */
export function FilterPopover({ locale, column, x, y, sheet, onApply, onSort, onClose }: FilterPopoverProps): React.ReactElement {
  const descriptor = useMemo(() => sheet.getFilterDomainDescriptor(column), [column, sheet]);
  const currentCriterion = sheet.getFilterCriterion(column);
  const modes = descriptor ? filterModeOptions(descriptor) : [];
  const initialMode = descriptor && modes.length > 0 ? currentMode(descriptor, currentCriterion) : 'values';
  const [mode, setMode] = useState<FilterMode>(initialMode);
  const [selected, setSelected] = useState<Set<string>>(() => currentCriterionValues(currentCriterion));
  const [includeBlank, setIncludeBlank] = useState(currentCriterion?.kind === 'values' ? currentCriterion.includeBlank : false);
  const currentDateGroups = currentCriterion?.kind === 'values' ? currentCriterion.dateGroups ?? [] : [];
  const [selectedDateGroups, setSelectedDateGroups] = useState<Set<string>>(() => new Set(currentDateGroups.map(dateGroupKey)));
  const [dateGroupsDirty, setDateGroupsDirty] = useState(false);
  const [search, setSearch] = useState('');
  const [operator, setOperator] = useState<CustomOperator>(currentCriterion?.kind === 'custom' ? currentCriterion.conditions[0]?.operator as CustomOperator ?? 'contains' : 'contains');
  const [operand, setOperand] = useState(currentCriterion?.kind === 'custom' ? String(currentCriterion.conditions[0]?.value ?? '') : '');
  const [useSecondCondition, setUseSecondCondition] = useState(Boolean(currentCriterion?.kind === 'custom' && currentCriterion.conditions[1]));
  const [secondOperator, setSecondOperator] = useState<CustomOperator>(currentCriterion?.kind === 'custom' ? currentCriterion.conditions[1]?.operator as CustomOperator ?? 'contains' : 'contains');
  const [secondOperand, setSecondOperand] = useState(currentCriterion?.kind === 'custom' ? String(currentCriterion.conditions[1]?.value ?? '') : '');
  const [join, setJoin] = useState<FilterJoin>(currentCriterion?.kind === 'custom' ? currentCriterion.join : 'and');
  const [numberMode, setNumberMode] = useState<'condition' | 'top10'>(currentCriterion?.kind === 'top10' ? 'top10' : 'condition');
  const [top, setTop] = useState(currentCriterion?.kind === 'top10' ? currentCriterion.top : true);
  const [percent, setPercent] = useState(currentCriterion?.kind === 'top10' ? currentCriterion.percent : false);
  const [rank, setRank] = useState(String(currentCriterion?.kind === 'top10' ? currentCriterion.rank : 10));
  const [dynamicType, setDynamicType] = useState<DynamicType>(currentCriterion?.kind === 'dynamic' ? currentCriterion.type : 'today');
  const [dateMode, setDateMode] = useState<DateMode>(currentCriterion?.kind === 'dynamic' ? 'dynamic' : 'condition');
  const [colorTarget, setColorTarget] = useState<'cell' | 'font'>(currentCriterion?.kind === 'color' ? currentCriterion.target : 'cell');
  const [color, setColor] = useState(currentCriterion?.kind === 'color' ? currentCriterion.style?.background ?? currentCriterion.style?.textColor ?? '' : '');
  const [icon, setIcon] = useState(currentCriterion?.kind === 'icon' ? `${currentCriterion.iconSet}:${currentCriterion.iconId}` : '');
  const [size, setSize] = useState({ width: 320, height: 560 });
  const resizeRef = useRef<{ pointerId: number; x: number; y: number; width: number; height: number } | null>(null);

  const valueOptions = useMemo<FilterValueOption[]>(() => (descriptor?.values ?? []).map((value) => ({ key: filterScalarKey(value), label: value === null || value === '' ? filterText(locale, 'blanks') : filterScalarLabel(value), value })), [descriptor, locale]);
  const dateDomain = descriptor?.dateDomain ?? [];
  const dateGroupNodes = useMemo(() => buildDateGroupNodes(dateDomain, currentDateGroups), [currentDateGroups, dateDomain]);
  const visibleValues = useMemo(() => valueOptions.filter((option) => option.label.toLocaleLowerCase().includes(search.toLocaleLowerCase())), [search, valueOptions]);
  const availableColors = useMemo(() => (descriptor?.colorDomain ?? []).filter((entry) => entry.target === colorTarget), [colorTarget, descriptor]);
  const allVisibleSelected = visibleValues.length > 0 && visibleValues.every((option) => selected.has(option.key));
  const activeMode = modes.includes(mode) ? mode : modes[0];
  const operators = activeMode ? filterOperatorsFor(activeMode) : [];
  const activeOperator = operators.includes(operator) ? operator : operators[0];
  const activeSecondOperator = operators.includes(secondOperator) ? secondOperator : operators[0];
  const { maxLeft, maxTop } = viewportBounds(size.width, size.height);
  const left = Math.min(Math.max(4, x), maxLeft);
  const topPosition = Math.min(Math.max(4, y), maxTop);

  useEffect(() => {
    if (!descriptor) return;
    const nextValues = currentCriterionValues(currentCriterion);
    setSelected(nextValues.size > 0 ? nextValues : new Set(descriptor.values.map(filterScalarKey)));
    setIncludeBlank(currentCriterion?.kind === 'values' ? currentCriterion.includeBlank : false);
    setSelectedDateGroups(new Set(currentDateGroups.map(dateGroupKey)));
    setDateGroupsDirty(false);
  }, [descriptor, currentCriterion, currentDateGroups]);

  const beginResize = (event: React.PointerEvent<HTMLElement>): void => {
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    resizeRef.current = { pointerId: event.pointerId, x: event.clientX, y: event.clientY, width: size.width, height: size.height };
  };
  const resize = (event: React.PointerEvent<HTMLElement>): void => {
    const active = resizeRef.current;
    if (!active || active.pointerId !== event.pointerId) return;
    const viewport = typeof window === 'undefined' ? { width: 736, height: 736 } : { width: window.innerWidth, height: window.innerHeight };
    setSize(clampFilterPopoverSize({ width: active.width + event.clientX - active.x, height: active.height + event.clientY - active.y }, viewport));
  };
  const endResize = (event: React.PointerEvent<HTMLElement>): void => { if (resizeRef.current?.pointerId === event.pointerId) resizeRef.current = null; };

  const apply = (): void => {
    if (!descriptor || !activeMode) return;
    if (activeMode === 'number' && numberMode === 'top10') {
      const numericRank = Number(rank);
      if (!Number.isSafeInteger(numericRank) || numericRank <= 0) return;
      onApply({ criterion: { kind: 'top10', top, percent, rank: numericRank } });
      return;
    }
    if (activeMode === 'date' && dateMode === 'dynamic') {
      if (!DYNAMIC_TYPES.includes(dynamicType)) return;
      onApply({ criterion: { kind: 'dynamic', type: dynamicType } });
      return;
    }
    if (activeMode === 'color') {
      if (!color || !availableColors.some((entry) => entry.color === color)) return;
      onApply({ criterion: { kind: 'color', target: colorTarget, dxfId: -1, style: colorTarget === 'cell' ? { background: color } : { textColor: color } } });
      return;
    }
    if (activeMode === 'icon') {
      const [iconSet, iconIdText] = icon.split(':');
      const iconId = Number(iconIdText);
      if (!iconSet || !Number.isSafeInteger(iconId) || !descriptor.iconDomain.some((entry) => entry.iconSet === iconSet && entry.iconId === iconId)) return;
      onApply({ criterion: { kind: 'icon', iconSet, iconId } });
      return;
    }
    if (activeMode !== 'values') {
      if (!activeOperator || !operators.includes(activeOperator) || ((activeMode === 'date' || activeMode === 'number') && !operand.trim())) return;
      const conditions: [{ operator: CustomOperator; value: string }, { operator: CustomOperator; value: string }?] = [{ operator: activeOperator, value: operand }];
      if (useSecondCondition) {
        if (!activeSecondOperator || !operators.includes(activeSecondOperator)) return;
        conditions.push({ operator: activeSecondOperator, value: secondOperand });
      }
      onApply({ criterion: { kind: 'custom', join, conditions } });
      return;
    }
    const selectedValues = valueOptions.filter((option) => selected.has(option.key)).map((option) => option.value).filter((value) => value !== null && value !== '');
    const literalOptions = valueOptions.filter((option) => option.value !== null && option.value !== '');
    const allLiteralsSelected = selectedValues.length === literalOptions.length && literalOptions.every((option) => selected.has(option.key));
    const dateGroups = dateGroupsDirty ? dateGroupNodes.filter((node) => selectedDateGroups.has(node.key)).map((node) => structuredClone(node.group)) : currentDateGroups.map((group) => structuredClone(group));
    const criterion = allLiteralsSelected && !includeBlank && dateGroups.length === 0 ? undefined : { kind: 'values' as const, values: selectedValues, includeBlank, ...(dateGroups.length ? { dateGroups } : {}) };
    onApply({ criterion });
  };

  return (
    <Box className="fixed z-50 min-w-0 max-w-[calc(100vw-1rem)] overflow-auto rounded-lg border border-slate-200 bg-white p-3 shadow-xl" style={{ left, top: topPosition, width: size.width, height: size.height }} role="dialog" aria-label={filterText(locale, 'title').replace('{column}', sheet.columns[column] ?? '')} onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); onClose(); } }} tabIndex={-1}>
      <Stack gap="sm">
        <Text size="xs" weight="semibold">{filterText(locale, 'title').replace('{column}', sheet.columns[column] ?? '')}</Text>
        <Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => onSort(true)}>{filterText(locale, 'sortAscending')}</Button><Button size="xs" variant="ghost" onClick={() => onSort(false)}>{filterText(locale, 'sortDescending')}</Button></Inline>
        {!descriptor || modes.length === 0 ? <Text size="sm" tone="subtle">{filterText(locale, 'noModes')}</Text> : (
          <>
            <Inline gap="xs" className="flex-wrap">{modes.map((candidate) => <Button key={candidate} size="xs" variant={activeMode === candidate ? 'soft' : 'ghost'} onClick={() => setMode(candidate)} aria-label={filterText(locale, modeTextKeys[candidate])}>{filterText(locale, modeTextKeys[candidate])}</Button>)}</Inline>
            {activeMode === 'color' ? (
              <Stack gap="xs">
                <Select sizeVariant="sm" aria-label={filterText(locale, 'filterColor')} value={colorTarget} onChange={(event) => setColorTarget(event.target.value as 'cell' | 'font')}><option value="cell">{filterText(locale, 'cellColor')}</option><option value="font">{filterText(locale, 'fontColor')}</option></Select>
                <Select sizeVariant="sm" aria-label={filterText(locale, 'filterColor')} value={color} onChange={(event) => setColor(event.target.value)} disabled={availableColors.length === 0}>{availableColors.map((entry) => <option key={`${entry.target}:${entry.color}`} value={entry.color}>{entry.color}</option>)}</Select>
                {availableColors.length === 0 ? <Text size="xs" tone="subtle">{filterText(locale, 'noColors')}</Text> : null}
              </Stack>
            ) : activeMode === 'icon' ? (
              <><Select sizeVariant="sm" aria-label={filterText(locale, 'filterIcon')} value={icon} onChange={(event) => setIcon(event.target.value)} disabled={descriptor.iconDomain.length === 0}>{descriptor.iconDomain.map((entry) => <option key={`${entry.iconSet}:${entry.iconId}`} value={`${entry.iconSet}:${entry.iconId}`}>{entry.iconSet} {entry.iconId}</option>)}</Select>{descriptor.iconDomain.length === 0 ? <Text size="xs" tone="subtle">{filterText(locale, 'noIcons')}</Text> : null}</>
            ) : activeMode === 'values' ? (
              <>
                <TextInput aria-label={filterText(locale, 'searchValues')} placeholder={filterText(locale, 'searchValues')} value={search} onChange={(event) => setSearch(event.target.value)} />
                <CheckToggle checked={allVisibleSelected} label={filterText(locale, 'selectAll')} onChange={(event) => { const next = new Set(selected); visibleValues.forEach((option) => event.target.checked ? next.add(option.key) : next.delete(option.key)); if (visibleValues.some((option) => option.value == null || option.value === '')) setIncludeBlank(event.target.checked); setSelected(next); }} />
                <VirtualList className="rounded border border-slate-100 p-2" height={224} itemHeight={28} items={visibleValues} itemKey={(option) => option.key} renderItem={(option) => <CheckToggle checked={selected.has(option.key)} label={option.label} onChange={(event) => { const next = new Set(selected); if (event.target.checked) next.add(option.key); else next.delete(option.key); if (option.value == null || option.value === '') setIncludeBlank(event.target.checked); setSelected(next); }} />} />
                {dateGroupNodes.length > 0 ? <Stack gap="xs" className="rounded border border-slate-100 p-2"><Text size="xs" weight="semibold">{filterText(locale, 'dateGroups')}</Text>{dateGroupNodes.map((node) => <CheckToggle key={node.key} checked={selectedDateGroups.has(node.key)} label={`${'\u00a0'.repeat(node.depth * 2)}${node.label}`} onChange={(event) => { const next = new Set(selectedDateGroups); if (event.target.checked) next.add(node.key); else next.delete(node.key); setSelectedDateGroups(next); setDateGroupsDirty(true); }} />)}</Stack> : null}
              </>
            ) : (
              <Stack gap="xs">
                {activeMode === 'number' ? <Select sizeVariant="sm" aria-label={filterText(locale, 'number')} value={numberMode} onChange={(event) => setNumberMode(event.target.value as 'condition' | 'top10')}><option value="condition">{filterText(locale, 'numberCondition')}</option><option value="top10">{filterText(locale, 'topBottom')}</option></Select> : null}
                {activeMode === 'date' ? <Select sizeVariant="sm" aria-label={filterText(locale, 'date')} value={dateMode} onChange={(event) => setDateMode(event.target.value as DateMode)}><option value="condition">{filterText(locale, 'dateCondition')}</option><option value="dynamic">{filterText(locale, 'dynamicDate')}</option></Select> : null}
                {activeMode === 'number' && numberMode === 'top10' ? <Inline gap="xs"><Select sizeVariant="sm" aria-label={filterText(locale, 'topBottom')} value={top ? 'top' : 'bottom'} onChange={(event) => setTop(event.target.value === 'top')}><option value="top">{filterText(locale, 'top')}</option><option value="bottom">{filterText(locale, 'bottom')}</option></Select><TextInput aria-label={filterText(locale, 'rank')} value={rank} onChange={(event) => setRank(event.target.value)} /><CheckToggle checked={percent} label={filterText(locale, 'percent')} onChange={(event) => setPercent(event.target.checked)} /></Inline> : activeMode === 'date' && dateMode === 'dynamic' ? <Select sizeVariant="sm" aria-label={filterText(locale, 'dynamicDate')} value={dynamicType} onChange={(event) => setDynamicType(event.target.value as DynamicType)}>{DYNAMIC_TYPES.map((type) => <option key={type} value={type}>{dynamicLabel(locale, type)}</option>)}</Select> : <>
                  <Select sizeVariant="sm" aria-label={filterText(locale, 'customValue')} value={activeOperator ?? ''} onChange={(event) => setOperator(event.target.value as CustomOperator)}>{operators.map((candidate) => <option key={candidate} value={candidate}>{filterText(locale, operatorTextKeys[candidate])}</option>)}</Select>
                  <TextInput aria-label={filterText(locale, 'customValue')} placeholder={filterText(locale, 'customValue')} value={operand} onChange={(event) => setOperand(event.target.value)} />
                  <CheckToggle checked={useSecondCondition} label={filterText(locale, 'secondCondition')} onChange={(event) => setUseSecondCondition(event.target.checked)} />
                  {useSecondCondition ? <><Select sizeVariant="sm" aria-label={filterText(locale, 'customValue')} value={activeSecondOperator ?? ''} onChange={(event) => setSecondOperator(event.target.value as CustomOperator)}>{operators.map((candidate) => <option key={candidate} value={candidate}>{filterText(locale, operatorTextKeys[candidate])}</option>)}</Select><TextInput aria-label={filterText(locale, 'customValue')} placeholder={filterText(locale, 'customValue')} value={secondOperand} onChange={(event) => setSecondOperand(event.target.value)} /><Select sizeVariant="sm" aria-label={filterText(locale, 'join')} value={join} onChange={(event) => setJoin(event.target.value as FilterJoin)}><option value="and">{filterText(locale, 'and')}</option><option value="or">{filterText(locale, 'or')}</option></Select></> : null}
                </>}
              </Stack>
            )}
          </>
        )}
        <Inline gap="sm" className="justify-end"><Button size="sm" variant="ghost" onClick={onClose}>{filterText(locale, 'cancel')}</Button><Button size="sm" variant="ghost" onClick={() => onApply({ criterion: undefined })}>{filterText(locale, 'clear')}</Button><Button size="sm" variant="primary" disabled={!descriptor || modes.length === 0} onClick={apply}>{filterText(locale, 'apply')}</Button></Inline>
        <Box role="separator" aria-label={filterText(locale, 'resize')} aria-orientation="horizontal" className="h-2 w-full cursor-se-resize rounded bg-slate-100" onPointerDown={beginResize} onPointerMove={resize} onPointerUp={endResize} onPointerCancel={endResize} />
      </Stack>
    </Box>
  );
}
