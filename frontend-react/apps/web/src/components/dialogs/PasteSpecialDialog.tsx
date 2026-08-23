import React from 'react';
import type { PasteMode } from '@react-sheets/sheet-features';
import { Box, Button, Inline, Stack, Text } from '@react-sheets/ui-system';

const PASTE_OPTIONS: Array<{ mode: PasteMode; label: string }> = [
  { mode: 'all', label: 'All' },
  { mode: 'values', label: 'Values' },
  { mode: 'formats', label: 'Formats' },
  { mode: 'formulas', label: 'Formulas' },
  { mode: 'transpose', label: 'Transpose' },
];

export interface PasteSpecialDialogProps {
  open: boolean;
  onClose: () => void;
  onPaste: (mode: PasteMode) => void;
}

export function PasteSpecialDialog({ open, onClose, onPaste }: PasteSpecialDialogProps): React.ReactElement | null {
  if (!open) return null;

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-24">
      <Box className="w-80 rounded-xl border border-slate-200 bg-white p-4 shadow-2xl" data-testid="paste-special-dialog">
        <Stack gap="md">
          <Text size="sm" weight="semibold">Paste Special</Text>
          <Stack gap="xs">
            {PASTE_OPTIONS.map((option) => (
              <Button
                key={option.mode}
                size="sm"
                variant="ghost"
                className="justify-start"
                data-testid={`paste-special-${option.mode}`}
                onClick={() => {
                  onPaste(option.mode);
                  onClose();
                }}
              >
                {option.label}
              </Button>
            ))}
          </Stack>
          <Inline gap="sm" className="justify-end">
            <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
          </Inline>
        </Stack>
      </Box>
    </Box>
  );
}
