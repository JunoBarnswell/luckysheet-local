import React, { useState } from 'react';
import type { GoToSpecialKind } from '@react-sheets/sheet-features';
import { Box, Button, Dialog, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale } from '../home/home-localization';

const SPECIAL_KINDS: Array<{ id: GoToSpecialKind; labelKey: 'blanks' | 'constants' | 'formulas' | 'comments' | 'errors' | 'visibleCells' }> = [
  { id: 'blanks', labelKey: 'blanks' },
  { id: 'constants', labelKey: 'constants' },
  { id: 'formulas', labelKey: 'formulas' },
  { id: 'comments', labelKey: 'comments' },
  { id: 'errors', labelKey: 'errors' },
  { id: 'visible', labelKey: 'visibleCells' },
];

export interface GoToDialogProps {
  open: boolean;
  locale?: Locale;
  onClose: () => void;
  onGoTo: (reference: string) => void;
  onGoToSpecial: (kind: GoToSpecialKind) => void;
}

export function GoToDialog({ open, locale, onClose, onGoTo, onGoToSpecial }: GoToDialogProps): React.ReactElement | null {
  const [reference, setReference] = useState('A1');
  const activeLocale = resolveHomeLocale(locale);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'goto')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="goto-dialog"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
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
            {homeText(activeLocale, 'ok')}
          </Button>
        </>
      )}
    >
      <Stack gap="md">
        <Stack gap="xs">
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'reference')}</Text>
          <TextInput
            aria-label={homeText(activeLocale, 'reference')}
            data-testid="goto-reference"
            placeholder={homeText(activeLocale, 'referencePlaceholder')}
            value={reference}
            autoFocus
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
          <Text size="xs" tone="subtle">{homeText(activeLocale, 'gotoSpecial')}</Text>
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
                {homeText(activeLocale, kind.labelKey)}
              </Button>
            ))}
          </Box>
        </Stack>
      </Stack>
    </Dialog>
  );
}
