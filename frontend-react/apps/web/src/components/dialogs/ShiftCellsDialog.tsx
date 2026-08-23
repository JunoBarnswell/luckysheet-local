import React, { useState } from 'react';
import { Button, Dialog, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';

export type ShiftCellsDirection = 'down' | 'up' | 'right' | 'left';

export interface ShiftCellsDialogProps {
  open: boolean;
  locale?: Locale;
  onClose: () => void;
  onShift: (direction: ShiftCellsDirection) => void;
}

const DIRECTIONS: Array<{ id: ShiftCellsDirection; labelKey: HomeUiTextKey; hintKey: HomeUiTextKey }> = [
  { id: 'down', labelKey: 'shiftDown', hintKey: 'shiftDownHint' },
  { id: 'up', labelKey: 'shiftUp', hintKey: 'shiftUpHint' },
  { id: 'right', labelKey: 'shiftRight', hintKey: 'shiftRightHint' },
  { id: 'left', labelKey: 'shiftLeft', hintKey: 'shiftLeftHint' },
];

export function ShiftCellsDialog({ open, locale, onClose, onShift }: ShiftCellsDialogProps): React.ReactElement | null {
  const [direction, setDirection] = useState<ShiftCellsDirection>('down');
  const activeLocale = resolveHomeLocale(locale);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'insertDeleteCells')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="shift-cells-dialog"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button
            size="sm"
            variant="primary"
            data-testid="shift-cells-apply"
            onClick={() => {
              onShift(direction);
              onClose();
            }}
          >
            {homeText(activeLocale, 'ok')}
          </Button>
        </>
      )}
    >
      <Stack gap="xs">
        {DIRECTIONS.map((entry) => (
          <Button
            key={entry.id}
            size="sm"
            variant={direction === entry.id ? 'secondary' : 'ghost'}
            className="h-auto justify-start px-3 py-2 text-left"
            aria-pressed={direction === entry.id}
            onClick={() => setDirection(entry.id)}
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
