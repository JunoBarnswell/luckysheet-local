import React, { useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Select, Stack, TextInput, Text } from '@react-sheets/ui-system';
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
  onClose: () => void;
}

type FilterMode = 'values' | 'custom';
type CustomOperator = 'equals' | 'contains' | 'beginsWith' | 'endsWith' | 'notEquals';

function criterionValues(criterion: FilterCriterion | undefined): string[] {
  return criterion?.kind === 'values' ? criterion.values.map((value) => String(value ?? '')) : [];
}

/** Excel-style filter task surface. UI draft state is local; only OK emits a command payload. */
export function FilterPopover({ column, x, y, sheet, onApply, onClose }: FilterPopoverProps): React.ReactElement {
  const values = useMemo(() => sheet.getFilterValueDomain(column), [column, sheet]);
  const currentCriterion = sheet.autoFilter?.columns[column]?.criterion;
  const [mode, setMode] = useState<FilterMode>(currentCriterion?.kind === 'custom' ? 'custom' : 'values');
  const [selected, setSelected] = useState<Set<string>>(() => new Set(criterionValues(currentCriterion).length > 0 ? criterionValues(currentCriterion) : values));
  const [search, setSearch] = useState('');
  const [operator, setOperator] = useState<CustomOperator>(currentCriterion?.kind === 'custom' ? currentCriterion.conditions[0]?.operator as CustomOperator ?? 'contains' : 'contains');
  const [operand, setOperand] = useState(currentCriterion?.kind === 'custom' ? String(currentCriterion.conditions[0]?.value ?? '') : '');

  const visibleValues = useMemo(
    () => values.filter((value) => value.toLocaleLowerCase().includes(search.toLocaleLowerCase())),
    [search, values],
  );
  const allVisibleSelected = visibleValues.length > 0 && visibleValues.every((value) => selected.has(value));

  const apply = (): void => {
    if (mode === 'custom') {
      onApply({ criterion: { kind: 'custom', join: 'and', conditions: [{ operator, value: operand }] } });
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
          <Button size="xs" variant={mode === 'values' ? 'soft' : 'ghost'} onClick={() => setMode('values')}>Values</Button>
          <Button size="xs" variant={mode === 'custom' ? 'soft' : 'ghost'} onClick={() => setMode('custom')}>Custom</Button>
        </Inline>

        {mode === 'custom' ? (
          <Stack gap="xs">
            <Select sizeVariant="sm" aria-label="Custom filter operator" value={operator} onChange={(event) => setOperator(event.target.value as CustomOperator)}>
              <option value="equals">Equals</option>
              <option value="notEquals">Not equals</option>
              <option value="contains">Contains</option>
              <option value="beginsWith">Begins with</option>
              <option value="endsWith">Ends with</option>
            </Select>
            <TextInput aria-label="Custom filter value" placeholder="Value" value={operand} onChange={(event) => setOperand(event.target.value)} />
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
            <Box className="max-h-56 overflow-y-auto rounded border border-slate-100 p-2">
              <Stack gap="xs">
                {visibleValues.map((value) => (
                  <CheckToggle
                    key={value || '__blank__'}
                    checked={selected.has(value)}
                    label={value === '' ? '(Blanks)' : value}
                    onChange={(event) => {
                      const next = new Set(selected);
                      if (event.target.checked) next.add(value);
                      else next.delete(value);
                      setSelected(next);
                    }}
                  />
                ))}
              </Stack>
            </Box>
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
