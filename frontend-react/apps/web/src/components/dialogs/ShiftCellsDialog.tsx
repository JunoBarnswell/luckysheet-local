import React, { useState } from 'react';
import { Button, Dialog, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';

export type ShiftCellsAxis = 'row' | 'column';
export type ShiftCellsOperation = 'insert' | 'delete';

export interface ShiftCellsDialogProps {
  open: boolean;
  locale?: Locale;
  operation: ShiftCellsOperation;
  onClose: () => void;
  onShift: (axis: ShiftCellsAxis) => void;
}

const DIRECTIONS: Record<ShiftCellsOperation, Array<{ id: ShiftCellsAxis; labelKey: HomeUiTextKey; hintKey: HomeUiTextKey }>> = {
  insert: [
    { id: 'row', labelKey: 'shiftDown', hintKey: 'shiftDownHint' },
    { id: 'column', labelKey: 'shiftRight', hintKey: 'shiftRightHint' },
  ],
  delete: [
    { id: 'row', labelKey: 'shiftUp', hintKey: 'shiftUpHint' },
    { id: 'column', labelKey: 'shiftLeft', hintKey: 'shiftLeftHint' },
  ],
};

export function ShiftCellsDialog({ open, locale, operation, onClose, onShift }: ShiftCellsDialogProps): React.ReactElement | null {
  const [axis, setAxis] = useState<ShiftCellsAxis>('row');
  const activeLocale = resolveHomeLocale(locale);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={operation === 'insert' ? homeText(activeLocale, 'insertCells') : homeText(activeLocale, 'deleteCells')}
      closeLabel={homeText(activeLocale, 'close')}
      testId={`${operation}-cells-dialog`}
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button
            size="sm"
            variant="primary"
            data-testid={`${operation}-cells-apply`}
            onClick={() => {
              onShift(axis);
              onClose();
            }}
          >
            {homeText(activeLocale, 'ok')}
          </Button>
        </>
      )}
    >
      <Stack gap="xs">
        {DIRECTIONS[operation].map((entry) => (
          <Button
            key={entry.id}
            size="sm"
            variant={axis === entry.id ? 'secondary' : 'ghost'}
            className="h-auto justify-start px-3 py-2 text-left"
            aria-pressed={axis === entry.id}
            onClick={() => setAxis(entry.id)}
          >
            <Stack gap="none" className="items-start">
              <Text size="sm" weight="medium">{homeText(activeLocale, entry.labelKey)}</Text>
              <Text size="xs" tone="subtle">{homeText(activeLocale, entry.hintKey)}</Text>
            </Stack>
          </Button>
        ))}
      </Stack>
    </Dialog>
  );
}
