import React, { useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Stack, TextInput, Text } from '@react-sheets/ui-system';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';

export interface FilterPatch {
  selectedValues: string[] | null;
}

export interface FilterPopoverProps {
  column: number;
  sheet: CanvasSheetSnapshot;
  onApply: (patch: FilterPatch) => void;
  onClose: () => void;
}

/** 列头筛选弹层:值清单勾选 + 清除 */
export function FilterPopover({ column, sheet, onApply, onClose }: FilterPopoverProps): React.ReactElement {
  const values = useMemo(() => {
    const set = new Set<string>();
    const scanEnd = Math.min(sheet.rowCount, 2000);
    for (let row = 1; row < scanEnd; row++) {
      const cell = sheet.getCell(row, column);
      if (cell && cell.value !== '') set.add(cell.value);
      if (set.size > 200) break;
    }
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [column, sheet]);

  const [selected, setSelected] = useState<Set<string>>(() => new Set(values));
  const [search, setSearch] = useState('');

  const visibleValues = useMemo(
    () => values.filter((value) => value.toLowerCase().includes(search.toLowerCase())),
    [search, values],
  );

  return (
    <Box className="absolute left-1/2 top-16 z-40 w-64 -translate-x-1/2 rounded-lg border border-slate-200 bg-white p-3 shadow-xl">
      <Stack gap="sm">
        <Text size="xs" weight="semibold">Filter column {sheet.columns[column]}</Text>
        <TextInput
          aria-label="Filter search"
          placeholder="Search values"
          value={search}
          onChange={(event) => setSearch(event.target.value)}
        />
        <Box className="max-h-48 overflow-y-auto">
          <Stack gap="xs">
            {visibleValues.map((value) => (
              <CheckToggle
                key={value}
                checked={selected.has(value)}
                label={value}
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
        <Inline gap="sm" className="justify-end">
          <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
          <Button
            size="sm"
            variant="ghost"
            onClick={() => {
              onClose();
              onApply({ selectedValues: null });
            }}
          >
            Clear
          </Button>
          <Button
            size="sm"
            variant="primary"
            onClick={() => {
              onClose();
              onApply({ selectedValues: [...selected] });
            }}
          >
            Apply
          </Button>
        </Inline>
      </Stack>
    </Box>
  );
}
