import React, { useEffect, useState } from 'react';
import { Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { RangeRef } from '@react-sheets/core-model';

export interface CreatePivotTableRequest {
  sourceId: string;
  destination: 'new-sheet' | 'existing-sheet';
  targetSheetId?: string;
  targetReference?: string;
}

export interface CreatePivotTableDialogProps {
  open: boolean;
  sourceRegion: RangeRef;
  sourceOptions: readonly { id: string; label: string }[];
  activeSheetName: string;
  onClose: () => void;
  onCreate: (request: CreatePivotTableRequest) => void;
}

/** Excel-style source/destination decision; validation and mutation remain in WorkbookSession. */
export function CreatePivotTableDialog({
  open,
  sourceRegion,
  sourceOptions,
  activeSheetName,
  onClose,
  onCreate,
}: CreatePivotTableDialogProps): React.ReactElement | null {
  const [destination, setDestination] = useState<CreatePivotTableRequest['destination']>('new-sheet');
  const [sourceId, setSourceId] = useState('current-region');
  const [targetReference, setTargetReference] = useState('A1');

  useEffect(() => {
    if (!open) return;
    setDestination('new-sheet');
    setSourceId(sourceOptions[0]?.id ?? 'current-region');
  }, [open, sourceOptions]);

  const sourceLabel = `${activeSheetName}!R${sourceRegion.startRow + 1}C${sourceRegion.startColumn + 1}:R${sourceRegion.endRow + 1}C${sourceRegion.endColumn + 1}`;
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
              sourceId,
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
          <Select aria-label="PivotTable source" sizeVariant="sm" value={sourceId} onChange={(event) => setSourceId(event.target.value)}>
            {sourceOptions.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
          </Select>
          <Text size="xs" tone="subtle">Current region · {sourceLabel}</Text>
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
