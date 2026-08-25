import React from 'react';
import { createPasteSpecialSpec, isPasteSpecialSpecSupported, type PasteSpecialSpec } from '@react-sheets/sheet-features';
import { Button, Dialog, Stack } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';

const PASTE_OPTIONS: Array<{ id: string; labelKey: HomeUiTextKey; spec: PasteSpecialSpec }> = [
  { id: 'all', labelKey: 'pasteAll', spec: createPasteSpecialSpec() },
  { id: 'values', labelKey: 'pasteValues', spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'formats', labelKey: 'pasteFormats', spec: createPasteSpecialSpec({ content: 'none', formatting: 'source-formatting', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'formulas', labelKey: 'pasteFormulas', spec: createPasteSpecialSpec({ content: 'formulas', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'values-number-format', labelKey: 'pasteValuesNumberFormat', spec: createPasteSpecialSpec({ content: 'values', formatting: 'number-format', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'formulas-number-format', labelKey: 'pasteFormulasNumberFormat', spec: createPasteSpecialSpec({ content: 'formulas', formatting: 'number-format', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'values-source-formatting', labelKey: 'pasteValuesSourceFormatting', spec: createPasteSpecialSpec({ content: 'values', formatting: 'source-formatting', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'all-except-borders', labelKey: 'pasteAllExceptBorders', spec: createPasteSpecialSpec({ formatting: 'all-except-borders' }) },
  { id: 'source-theme', labelKey: 'pasteSourceTheme', spec: createPasteSpecialSpec({ formatting: 'source-theme' }) },
  { id: 'comments-notes', labelKey: 'pasteCommentsNotes', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', metadata: { commentsNotes: true, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'validation', labelKey: 'pasteValidation', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', metadata: { commentsNotes: false, validation: true, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'column-widths', labelKey: 'pasteColumnWidths', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: true, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'conditional-formats', labelKey: 'pasteConditionalFormats', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: true, hyperlinks: false } }) },
  { id: 'add', labelKey: 'pasteAdd', spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', operation: 'add', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'subtract', labelKey: 'pasteSubtract', spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', operation: 'subtract', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'multiply', labelKey: 'pasteMultiply', spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', operation: 'multiply', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'divide', labelKey: 'pasteDivide', spec: createPasteSpecialSpec({ content: 'values', formatting: 'none', operation: 'divide', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
  { id: 'skip-blanks', labelKey: 'pasteSkipBlanks', spec: createPasteSpecialSpec({ skipBlanks: true }) },
  { id: 'transpose', labelKey: 'pasteTranspose', spec: createPasteSpecialSpec({ transpose: true }) },
  { id: 'link', labelKey: 'pasteLink', spec: createPasteSpecialSpec({ content: 'none', formatting: 'none', link: true, metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }) },
];

export interface PasteSpecialDialogProps {
  open: boolean;
  locale?: Locale;
  onClose: () => void;
  onPaste: (spec: PasteSpecialSpec) => Promise<unknown>;
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
            key={option.id}
            size="sm"
            variant={option.id === 'all' ? 'secondary' : 'ghost'}
            className="justify-start"
            disabled={!isPasteSpecialSpecSupported(option.spec)}
            data-testid={`paste-special-${option.id}`}
            onClick={async () => {
              await onPaste(option.spec);
            }}
          >
            {homeText(activeLocale, option.labelKey)}
          </Button>
        ))}
      </Stack>
    </Dialog>
  );
}
