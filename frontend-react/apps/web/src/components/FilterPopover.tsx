import React, { useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Select, Stack, TextInput, Text, VirtualList } from '@react-sheets/ui-system';
import type { FilterCriterion } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';

export interface FilterPatch {
  criterion?: FilterCriterion;
}

export interface FilterPopoverProps {
  column: number;
  x: number;
  y: number;
  sheet: CanvasSheetSnapshot;
  onApply: (patch: FilterPatch) => void;
  onSort: (ascending: boolean) => void;
  onClose: () => void;
}

type FilterMode = 'values' | 'text' | 'number' | 'date' | 'color' | 'icon';
type CustomOperator = 'equals' | 'notEquals' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'contains' | 'notContains' | 'beginsWith' | 'endsWith';
type FilterJoin = 'and' | 'or';
type DateMode = 'condition' | 'dynamic';
type DynamicType = 'today' | 'yesterday' | 'tomorrow' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth' | 'thisQuarter' | 'lastQuarter' | 'nextQuarter' | 'thisYear' | 'lastYear' | 'nextYear' | 'yearToDate';

function criterionValues(criterion: FilterCriterion | undefined): string[] {
  return criterion?.kind === 'values' ? criterion.values.map((value) => String(value ?? '')) : [];
}

/** Excel-style filter task surface. UI draft state is local; only OK emits a command payload. */
export function FilterPopover({ column, x, y, sheet, onApply, onSort, onClose }: FilterPopoverProps): React.ReactElement {
  const values = useMemo(() => sheet.getFilterValueDomain(column), [column, sheet]);
  const colors = useMemo(() => sheet.getFilterColorDomain(column), [column, sheet]);
  const icons = useMemo(() => sheet.getFilterIconDomain(column), [column, sheet]);
  const currentCriterion = sheet.getFilterCriterion(column);
  const [mode, setMode] = useState<FilterMode>(currentCriterion?.kind === 'custom' ? 'text' : currentCriterion?.kind === 'dynamic' ? 'date' : currentCriterion?.kind === 'top10' ? 'number' : currentCriterion?.kind === 'color' ? 'color' : currentCriterion?.kind === 'icon' ? 'icon' : 'values');
  const [selected, setSelected] = useState<Set<string>>(() => new Set(currentCriterion?.kind === 'values' ? criterionValues(currentCriterion) : values));
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
  const [color, setColor] = useState(currentCriterion?.kind === 'color' ? currentCriterion.style?.background ?? currentCriterion.style?.textColor ?? colors[0]?.color ?? '' : colors[0]?.color ?? '');
  const [icon, setIcon] = useState(currentCriterion?.kind === 'icon' ? `${currentCriterion.iconSet}:${currentCriterion.iconId}` : icons[0] ? `${icons[0].iconSet}:${icons[0].iconId}` : '');

  const visibleValues = useMemo(
    () => values.filter((value) => value.toLocaleLowerCase().includes(search.toLocaleLowerCase())),
    [search, values],
  );
  const availableColors = useMemo(() => colors.filter((entry) => entry.target === colorTarget), [colorTarget, colors]);
  const allVisibleSelected = visibleValues.length > 0 && visibleValues.every((value) => selected.has(value));

  const apply = (): void => {
    if (mode !== 'values') {
      if (mode === 'number' && numberMode === 'top10') {
        onApply({ criterion: { kind: 'top10', top, percent, rank: Math.max(1, Number(rank) || 1) } });
        return;
      }
      if (mode === 'date' && dateMode === 'dynamic') {
        onApply({ criterion: { kind: 'dynamic', type: dynamicType } });
        return;
      }
      if (mode === 'color') {
        if (!color) return;
        onApply({ criterion: { kind: 'color', target: colorTarget, dxfId: -1, style: colorTarget === 'cell' ? { background: color } : { textColor: color } } });
        return;
      }
      if (mode === 'icon') {
        const [iconSet, iconId] = icon.split(':');
        if (!iconSet || !Number.isSafeInteger(Number(iconId))) return;
        onApply({ criterion: { kind: 'icon', iconSet, iconId: Number(iconId) } });
        return;
      }
      const conditions: [{ operator: CustomOperator; value: string }, { operator: CustomOperator; value: string }?] = [{ operator, value: operand }];
      if (useSecondCondition) conditions.push({ operator: secondOperator, value: secondOperand });
      onApply({ criterion: { kind: 'custom', join, conditions } });
      return;
    }
    const selectedValues = [...selected];
    const criterion = selectedValues.length === values.length && values.every((value) => selected.has(value))
      ? undefined
      : { kind: 'values' as const, values: selectedValues, includeBlank: selected.has('') };
    onApply({ criterion });
  };

  return (
    <Box
      className="fixed z-50 w-72 rounded-lg border border-slate-200 bg-white p-3 shadow-xl"
      style={{ left: x, top: y }}
      role="dialog"
      aria-label={`Filter ${sheet.columns[column] ?? ''}`}
    >
      <Stack gap="sm">
        <Text size="xs" weight="semibold">Filter column {sheet.columns[column]}</Text>
        <Inline gap="xs">
          <Button size="xs" variant="ghost" onClick={() => onSort(true)}>Sort A to Z</Button>
          <Button size="xs" variant="ghost" onClick={() => onSort(false)}>Sort Z to A</Button>
        </Inline>
        <Inline gap="xs">
          <Button size="xs" variant={mode === 'values' ? 'soft' : 'ghost'} onClick={() => setMode('values')}>Values</Button>
          <Button size="xs" variant={mode === 'text' ? 'soft' : 'ghost'} onClick={() => setMode('text')}>Text</Button>
          <Button size="xs" variant={mode === 'number' ? 'soft' : 'ghost'} onClick={() => setMode('number')}>Number</Button>
          <Button size="xs" variant={mode === 'date' ? 'soft' : 'ghost'} onClick={() => setMode('date')}>Date</Button>
          <Button size="xs" variant={mode === 'color' ? 'soft' : 'ghost'} onClick={() => setMode('color')}>Color</Button>
          <Button size="xs" variant={mode === 'icon' ? 'soft' : 'ghost'} onClick={() => setMode('icon')}>Icon</Button>
        </Inline>

        {mode === 'color' ? (
          <Stack gap="xs">
            <Select sizeVariant="sm" aria-label="Filter color target" value={colorTarget} onChange={(event) => setColorTarget(event.target.value as 'cell' | 'font')}>
              <option value="cell">Cell color</option>
              <option value="font">Font color</option>
            </Select>
            <Select sizeVariant="sm" aria-label="Filter color" value={color} onChange={(event) => setColor(event.target.value)}>
              {availableColors.map((entry) => <option key={`${entry.target}:${entry.color}`} value={entry.color}>{entry.color}</option>)}
            </Select>
          </Stack>
        ) : mode === 'icon' ? (
          <Select sizeVariant="sm" aria-label="Filter icon" value={icon} onChange={(event) => setIcon(event.target.value)}>
            {icons.map((entry) => <option key={`${entry.iconSet}:${entry.iconId}`} value={`${entry.iconSet}:${entry.iconId}`}>{entry.iconSet} {entry.iconId}</option>)}
          </Select>
        ) : mode !== 'values' ? (
          <Stack gap="xs">
            {mode === 'number' ? (
              <Select sizeVariant="sm" aria-label="Number filter mode" value={numberMode} onChange={(event) => setNumberMode(event.target.value as 'condition' | 'top10')}>
                <option value="condition">Number condition</option>
                <option value="top10">Top 10</option>
              </Select>
            ) : null}
            {mode === 'date' ? (
              <Select sizeVariant="sm" aria-label="Date filter mode" value={dateMode} onChange={(event) => setDateMode(event.target.value as DateMode)}>
                <option value="condition">Date condition</option>
                <option value="dynamic">Dynamic date</option>
              </Select>
            ) : null}
            {mode === 'number' && numberMode === 'top10' ? (
              <Inline gap="xs">
                <Select sizeVariant="sm" aria-label="Top or bottom" value={top ? 'top' : 'bottom'} onChange={(event) => setTop(event.target.value === 'top')}>
                  <option value="top">Top</option>
                  <option value="bottom">Bottom</option>
                </Select>
                <TextInput aria-label="Top 10 rank" value={rank} onChange={(event) => setRank(event.target.value)} />
                <CheckToggle checked={percent} label="%" onChange={(event) => setPercent(event.target.checked)} />
              </Inline>
            ) : mode === 'date' && dateMode === 'dynamic' ? (
              <Select sizeVariant="sm" aria-label="Dynamic date filter" value={dynamicType} onChange={(event) => setDynamicType(event.target.value as DynamicType)}>
                <option value="today">Today</option>
                <option value="yesterday">Yesterday</option>
                <option value="tomorrow">Tomorrow</option>
                <option value="thisWeek">This week</option>
                <option value="lastWeek">Last week</option>
                <option value="nextWeek">Next week</option>
                <option value="thisMonth">This month</option>
                <option value="lastMonth">Last month</option>
                <option value="nextMonth">Next month</option>
                <option value="thisYear">This year</option>
                <option value="lastYear">Last year</option>
                <option value="nextYear">Next year</option>
                <option value="yearToDate">Year to date</option>
              </Select>
            ) : (
              <>
            <Select sizeVariant="sm" aria-label="Custom filter operator" value={operator} onChange={(event) => setOperator(event.target.value as CustomOperator)}>
              <option value="equals">Equals</option>
              <option value="notEquals">Not equals</option>
              {mode !== 'text' ? <option value="lessThan">Less than</option> : null}
              {mode !== 'text' ? <option value="lessThanOrEqual">Less than or equal</option> : null}
              {mode !== 'text' ? <option value="greaterThan">Greater than</option> : null}
              {mode !== 'text' ? <option value="greaterThanOrEqual">Greater than or equal</option> : null}
              <option value="contains">Contains</option>
              <option value="notContains">Does not contain</option>
              <option value="beginsWith">Begins with</option>
              <option value="endsWith">Ends with</option>
            </Select>
            <TextInput aria-label="Custom filter value" placeholder={mode === 'date' ? 'YYYY-MM-DD' : 'Value'} value={operand} onChange={(event) => setOperand(event.target.value)} />
            <CheckToggle checked={useSecondCondition} label="Use second condition" onChange={(event) => setUseSecondCondition(event.target.checked)} />
            {useSecondCondition ? (
              <>
                <Select sizeVariant="sm" aria-label="Second custom filter operator" value={secondOperator} onChange={(event) => setSecondOperator(event.target.value as CustomOperator)}>
                  <option value="equals">Equals</option>
                  <option value="notEquals">Not equals</option>
                  {mode !== 'text' ? <option value="lessThan">Less than</option> : null}
                  {mode !== 'text' ? <option value="lessThanOrEqual">Less than or equal</option> : null}
                  {mode !== 'text' ? <option value="greaterThan">Greater than</option> : null}
                  {mode !== 'text' ? <option value="greaterThanOrEqual">Greater than or equal</option> : null}
                  <option value="contains">Contains</option>
                  <option value="notContains">Does not contain</option>
                  <option value="beginsWith">Begins with</option>
                  <option value="endsWith">Ends with</option>
                </Select>
                <TextInput aria-label="Second custom filter value" placeholder={mode === 'date' ? 'YYYY-MM-DD' : 'Value'} value={secondOperand} onChange={(event) => setSecondOperand(event.target.value)} />
                <Select sizeVariant="sm" aria-label="Custom filter join" value={join} onChange={(event) => setJoin(event.target.value as FilterJoin)}>
                  <option value="and">And</option>
                  <option value="or">Or</option>
                </Select>
              </>
            ) : null}
              </>
            )}
          </Stack>
        ) : (
          <>
            <TextInput
              aria-label="Filter search"
              placeholder="Search values"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
            />
            <CheckToggle
              checked={allVisibleSelected}
              label="Select All"
              onChange={(event) => {
                const next = new Set(selected);
                visibleValues.forEach((value) => event.target.checked ? next.add(value) : next.delete(value));
                setSelected(next);
              }}
            />
            <VirtualList
              className="rounded border border-slate-100 p-2"
              height={224}
              itemHeight={28}
              items={visibleValues}
              itemKey={(value) => value || '__blank__'}
              renderItem={(value) => (
                <CheckToggle
                  checked={selected.has(value)}
                  label={value === '' ? '(Blanks)' : value}
                  onChange={(event) => {
                    const next = new Set(selected);
                    if (event.target.checked) next.add(value);
                    else next.delete(value);
                    setSelected(next);
                  }}
                />
              )}
            />
          </>
        )}

        <Inline gap="sm" className="justify-end">
          <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
          <Button size="sm" variant="ghost" onClick={() => onApply({ criterion: undefined })}>Clear</Button>
          <Button size="sm" variant="primary" onClick={apply}>OK</Button>
        </Inline>
      </Stack>
    </Box>
  );
}
