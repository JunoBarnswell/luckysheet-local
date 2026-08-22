import React, { useState } from 'react';
import { Button, Dialog, Select, Stack, Text } from '@react-sheets/ui-system';

export interface SortDialogProps {
  open: boolean;
  columns: string[];
  onClose: () => void;
  onSort: (columnIdx: number, ascending: boolean, hasHeader: boolean) => void;
}

export function SortDialog({ open, columns, onClose, onSort }: SortDialogProps) {
  const [selectedCol, setSelectedCol] = useState(0);
  const [ascending, setAscending] = useState(true);
  const [hasHeader, setHasHeader] = useState(true);

  const handleApply = () => {
    onSort(selectedCol, ascending, hasHeader);
    onClose();
  };

  return (
    <Dialog
      open={open}
      title="Sort Range"
      description="Sort the selected cells by a specific column."
      onClose={onClose}
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="primary" onClick={handleApply}>
            Apply Sort
          </Button>
        </>
      }
    >
      <Stack gap="md">
        <div>
          <Text size="xs" weight="medium" className="mb-1 text-slate-700">
            Sort by Column
          </Text>
          <Select
            value={selectedCol}
            onChange={(e) => setSelectedCol(Number(e.target.value))}
            sizeVariant="sm"
          >
            {columns.map((col, idx) => (
              <option key={col} value={idx}>
                Column {col} (Index {idx + 1})
              </option>
            ))}
          </Select>
        </div>

        <div>
          <Text size="xs" weight="medium" className="mb-1 text-slate-700">
            Order
          </Text>
          <Select
            value={ascending ? 'asc' : 'desc'}
            onChange={(e) => setAscending(e.target.value === 'asc')}
            sizeVariant="sm"
          >
            <option value="asc">Ascending (A → Z, 0 → 9)</option>
            <option value="desc">Descending (Z → A, 9 → 0)</option>
          </Select>
        </div>

        <label className="flex items-center gap-2 cursor-pointer pt-1">
          <input
            type="checkbox"
            checked={hasHeader}
            onChange={(e) => setHasHeader(e.target.checked)}
            className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
          />
          <Text size="xs" className="text-slate-700">
            My data has headers (skip first row)
          </Text>
        </label>
      </Stack>
    </Dialog>
  );
}
