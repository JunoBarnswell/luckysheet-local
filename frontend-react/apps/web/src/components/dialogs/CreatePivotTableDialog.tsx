import React, { useEffect, useState } from 'react';
import { Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { RangeRef } from '@react-sheets/core-model';

export interface CreatePivotTableRequest {
  destination: 'new-sheet' | 'existing-sheet';
  targetSheetId?: string;
  targetReference?: string;
}

export interface CreatePivotTableDialogProps {
  open: boolean;
  sourceRange: RangeRef;
  activeSheetName: string;
  onClose: () => void;
  onCreate: (request: CreatePivotTableRequest) => void;
}

/** Excel-style source/destination decision; validation and mutation remain in WorkbookSession. */
export function CreatePivotTableDialog({
  open,
  sourceRange,
  activeSheetName,
  onClose,
  onCreate,
}: CreatePivotTableDialogProps): React.ReactElement | null {
  const [destination, setDestination] = useState<CreatePivotTableRequest['destination']>('new-sheet');
  const [targetReference, setTargetReference] = useState('A1');

  useEffect(() => {
    if (open) setDestination('new-sheet');
  }, [open]);

  const sourceLabel = `${activeSheetName}!R${sourceRange.startRow + 1}C${sourceRange.startColumn + 1}:R${sourceRange.endRow + 1}C${sourceRange.endColumn + 1}`;
  return (
    <Dialog
      open={open}
      title="Create PivotTable"
      description="Choose where the blank PivotTable report will be placed."
      onClose={onClose}
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
          <Button
            size="sm"
            variant="primary"
            disabled={destination === 'existing-sheet' && !targetReference.trim()}
            data-testid="create-pivot-confirm"
            onClick={() => onCreate({
              destination,
              ...(destination === 'existing-sheet' ? { targetReference: targetReference.trim() } : {}),
            })}
          >
            Create
          </Button>
        </>
      )}
    >
      <Stack gap="md" data-testid="create-pivot-dialog">
        <Stack gap="xs">
          <Text size="xs" tone="subtle">Table or range</Text>
          <Text size="sm" weight="semibold">{sourceLabel}</Text>
        </Stack>
        <Stack gap="xs">
          <Text size="xs" tone="subtle">Choose where you want the PivotTable report to be placed</Text>
          <Inline gap="sm" className="flex-wrap">
            <Button
              size="sm"
              variant={destination === 'new-sheet' ? 'primary' : 'outline'}
              onClick={() => setDestination('new-sheet')}
            >
              New Worksheet
            </Button>
            <Button
              size="sm"
              variant={destination === 'existing-sheet' ? 'primary' : 'outline'}
              onClick={() => setDestination('existing-sheet')}
            >
              Existing Worksheet
            </Button>
          </Inline>
        </Stack>
        {destination === 'existing-sheet' ? (
          <Stack gap="xs">
            <Text size="xs" tone="subtle">Location on {activeSheetName}</Text>
            <TextInput
              aria-label="PivotTable location"
              data-testid="create-pivot-location"
              placeholder="A1"
              value={targetReference}
              onChange={(event) => setTargetReference(event.target.value)}
            />
          </Stack>
        ) : (
          <Text size="xs" tone="muted">A new worksheet will contain a blank PivotTable and its Field List.</Text>
        )}
      </Stack>
    </Dialog>
  );
}
