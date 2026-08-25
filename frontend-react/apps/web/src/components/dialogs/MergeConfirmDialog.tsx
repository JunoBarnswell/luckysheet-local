import React from 'react';
import { Button, Dialog, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { translateRibbonText } from '../../i18n';
import { homeTemplate, homeText, resolveHomeLocale } from '../home/home-localization';

export interface MergeConfirmDialogProps {
  open: boolean;
  discardedCellCount: number;
  operation: 'center' | 'cells' | 'across' | 'unmerge';
  locale?: Locale;
  onCancel: () => void;
  onConfirm: () => void;
}

/** Explicit destructive-action gate for Merge & Center. */
export function MergeConfirmDialog({
  open,
  discardedCellCount,
  operation,
  locale,
  onCancel,
  onConfirm,
}: MergeConfirmDialogProps): React.ReactElement {
  const activeLocale = resolveHomeLocale(locale);
  return (
    <Dialog
      open={open}
      onClose={onCancel}
      title={`${translateRibbonText(activeLocale, operation === 'center' ? 'commands.mergeCenter' : operation === 'across' ? 'commands.mergeAcross' : 'commands.mergeCells')}?`}
      closeLabel={homeText(activeLocale, 'close')}
      testId="merge-confirm-dialog"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onCancel}>{homeText(activeLocale, 'cancel')}</Button>
          <Button size="sm" variant="primary" data-testid="merge-confirm" onClick={onConfirm}>
            {homeText(activeLocale, 'mergeConfirm')}
          </Button>
        </>
      )}
    >
      <Stack gap="sm">
        <Text size="sm">{homeTemplate(activeLocale, 'mergeDataLossDescription', { count: discardedCellCount })}</Text>
      </Stack>
    </Dialog>
  );
}
