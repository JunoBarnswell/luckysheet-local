import React, { useState } from 'react';
import type { GoToSpecialKind } from '@react-sheets/sheet-features';
import { Box, Button, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';

const SPECIAL_KINDS: Array<{ id: GoToSpecialKind; label: string }> = [
  { id: 'blanks', label: 'Blanks' },
  { id: 'constants', label: 'Constants' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'comments', label: 'Comments' },
  { id: 'errors', label: 'Errors' },
  { id: 'visible', label: 'Visible cells only' },
];

export interface GoToDialogProps {
  open: boolean;
  onClose: () => void;
  onGoTo: (reference: string) => void;
  onGoToSpecial: (kind: GoToSpecialKind) => void;
}

export function GoToDialog({ open, onClose, onGoTo, onGoToSpecial }: GoToDialogProps): React.ReactElement | null {
  const [reference, setReference] = useState('A1');

  if (!open) return null;

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-24">
      <Box className="w-[28rem] rounded-xl border border-slate-200 bg-white p-4 shadow-2xl" data-testid="goto-dialog">
        <Stack gap="md">
          <Text size="sm" weight="semibold">Go To</Text>
          <Stack gap="xs">
            <Text size="xs" tone="subtle">Reference</Text>
            <TextInput
              aria-label="Reference"
              data-testid="goto-reference"
              placeholder="A1, A1:C10, or defined name"
              value={reference}
              onChange={(event) => setReference(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && reference.trim()) {
                  event.preventDefault();
                  onGoTo(reference.trim());
                  onClose();
                }
              }}
            />
          </Stack>
          <Stack gap="xs">
            <Text size="xs" tone="subtle">Go to special</Text>
            <Box className="grid grid-cols-2 gap-1">
              {SPECIAL_KINDS.map((kind) => (
                <Button
                  key={kind.id}
                  size="sm"
                  variant="ghost"
                  className="justify-start"
                  data-testid={`goto-special-${kind.id}`}
                  onClick={() => {
                    onGoToSpecial(kind.id);
                    onClose();
                  }}
                >
                  {kind.label}
                </Button>
              ))}
            </Box>
          </Stack>
          <Inline gap="sm" className="justify-end">
            <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
            <Button
              size="sm"
              variant="primary"
              data-testid="goto-apply"
              disabled={!reference.trim()}
              onClick={() => {
                onGoTo(reference.trim());
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
