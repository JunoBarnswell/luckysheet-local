import React, { useState } from 'react';
import { Box, Button, Inline, Stack, Text } from '@react-sheets/ui-system';

export type ShiftCellsDirection = 'down' | 'up' | 'right' | 'left';

export interface ShiftCellsDialogProps {
  open: boolean;
  onClose: () => void;
  onShift: (direction: ShiftCellsDirection) => void;
}

const DIRECTIONS: Array<{ id: ShiftCellsDirection; label: string; hint: string }> = [
  { id: 'down', label: 'Shift cells down', hint: 'Insert blank cells and move selection down' },
  { id: 'up', label: 'Shift cells up', hint: 'Delete selected cells and move cells up' },
  { id: 'right', label: 'Shift cells right', hint: 'Insert blank cells and move selection right' },
  { id: 'left', label: 'Shift cells left', hint: 'Delete selected cells and move cells left' },
];

export function ShiftCellsDialog({ open, onClose, onShift }: ShiftCellsDialogProps): React.ReactElement | null {
  const [direction, setDirection] = useState<ShiftCellsDirection>('down');

  if (!open) return null;

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-24">
      <Box className="w-[24rem] rounded-xl border border-slate-200 bg-white p-4 shadow-2xl" data-testid="shift-cells-dialog">
        <Stack gap="md">
          <Text size="sm" weight="semibold">Insert / Delete Cells</Text>
          <Stack gap="xs">
            {DIRECTIONS.map((entry) => (
              <label key={entry.id} className="flex cursor-pointer items-start gap-2 rounded-md border border-slate-100 px-2 py-1.5 hover:bg-slate-50">
                <input
                  checked={direction === entry.id}
                  className="mt-0.5"
                  name="shift-direction"
                  type="radio"
                  value={entry.id}
                  onChange={() => setDirection(entry.id)}
                />
                <Stack gap="none">
                  <Text size="xs" weight="medium">{entry.label}</Text>
                  <Text size="xs" tone="subtle">{entry.hint}</Text>
                </Stack>
              </label>
            ))}
          </Stack>
          <Inline gap="sm" className="justify-end">
            <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
            <Button
              size="sm"
              variant="primary"
              data-testid="shift-cells-apply"
              onClick={() => {
                onShift(direction);
                onClose();
              }}
            >
              OK
            </Button>
          </Inline>
        </Stack>
      </Box>
    </Box>
  );
}
