import React, { useMemo, useState } from 'react';
import { Box, Button, Inline, Stack, TextInput, Text } from '@react-sheets/ui-system';
import type { SheetView } from '../state/workspace';

export interface FilterPatch {
  selectedValues: string[] | null;
}

export interface FilterPopoverProps {
  column: number;
  sheet: SheetView;
  onApply: (patch: FilterPatch) => void;
  onClose: () => void;
}

/** 列头筛选弹层:值清单勾选 + 清除 */
export function FilterPopover({ column, sheet, onApply, onClose }: FilterPopoverProps): React.ReactElement {
  const values = useMemo(() => {
    const set = new Set<string>();
    for (let row = 1; row < sheet.rows.length; row++) {
      const cell = sheet.rows[row]?.cells[column];
      if (cell && cell.value !== '') set.add(cell.value);
      if (set.size > 200) break;
    }
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [column, sheet.rows]);

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
              <label key={value} className="flex items-center gap-2 text-xs text-slate-700">
                <input
                  type="checkbox"
                  checked={selected.has(value)}
                  onChange={(event) => {
                    const next = new Set(selected);
                    if (event.target.checked) next.add(value);
                    else next.delete(value);
                    setSelected(next);
                  }}
                />
                {value}
              </label>
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
