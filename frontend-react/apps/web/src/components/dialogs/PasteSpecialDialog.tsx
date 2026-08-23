import React from 'react';
import type { PasteMode } from '@react-sheets/sheet-features';
import { Button, Dialog, Stack } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';

const PASTE_OPTIONS: Array<{ mode: PasteMode; labelKey: HomeUiTextKey }> = [
  { mode: 'all', labelKey: 'pasteAll' },
  { mode: 'values', labelKey: 'pasteValues' },
  { mode: 'formats', labelKey: 'pasteFormats' },
  { mode: 'formulas', labelKey: 'pasteFormulas' },
  { mode: 'transpose', labelKey: 'pasteTranspose' },
];

export interface PasteSpecialDialogProps {
  open: boolean;
  locale?: Locale;
  onClose: () => void;
  onPaste: (mode: PasteMode) => void;
}

export function PasteSpecialDialog({ open, locale, onClose, onPaste }: PasteSpecialDialogProps): React.ReactElement | null {
  const activeLocale = resolveHomeLocale(locale);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'pasteSpecial')}
      description={homeText(activeLocale, 'pasteSpecialDescription')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="paste-special-dialog"
      footer={<Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>}
    >
      <Stack gap="xs">
        {PASTE_OPTIONS.map((option) => (
          <Button
            key={option.mode}
            size="sm"
            variant={option.mode === 'all' ? 'secondary' : 'ghost'}
            className="justify-start"
            data-testid={`paste-special-${option.mode}`}
            onClick={() => {
              onPaste(option.mode);
              onClose();
            }}
          >
            {homeText(activeLocale, option.labelKey)}
          </Button>
        ))}
      </Stack>
    </Dialog>
  );
}
