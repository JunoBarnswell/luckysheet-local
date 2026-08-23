import React, { useEffect, useState } from 'react';
import { Button, CheckToggle, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeTemplate, homeText, resolveHomeLocale } from '../home/home-localization';

export interface FindReplaceDialogProps {
  open: boolean;
  initialFind?: string;
  locale?: Locale;
  onClose: () => void;
  onReplaceAll: (params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }) => number;
}

/** UI-only Find & Replace form; the host owns command execution and result count. */
export function FindReplaceDialog({ initialFind = '', locale, open, onClose, onReplaceAll }: FindReplaceDialogProps): React.ReactElement | null {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [entireCell, setEntireCell] = useState(false);
  const [scope, setScope] = useState<'sheet' | 'workbook'>('sheet');
  const [result, setResult] = useState<number | null>(null);
  const activeLocale = resolveHomeLocale(locale);

  useEffect(() => {
    if (!open) return;
    setFind(initialFind);
    setResult(null);
  }, [initialFind, open]);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'findReplace')}
      description={homeText(activeLocale, 'findReplaceDescription')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="find-replace-dialog"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button
            size="sm"
            variant="primary"
            disabled={!find}
            data-testid="find-replace-all"
            onClick={() => {
              const count = onReplaceAll({ find, replace, matchCase, entireCell, scope });
              setResult(count);
            }}
          >
            {homeText(activeLocale, 'replaceAll')}
          </Button>
        </>
      )}
    >
      <Stack gap="sm">
        <TextInput
          aria-label={homeText(activeLocale, 'find')}
          placeholder={homeText(activeLocale, 'find')}
          value={find}
          autoFocus
          onChange={(event) => setFind(event.target.value)}
        />
        <TextInput
          aria-label={homeText(activeLocale, 'replaceWith')}
          placeholder={homeText(activeLocale, 'replaceWith')}
          value={replace}
          onChange={(event) => setReplace(event.target.value)}
        />
        <Inline gap="md" className="flex-wrap items-center">
          <CheckToggle checked={matchCase} label={homeText(activeLocale, 'matchCase')} onChange={(event) => setMatchCase(event.target.checked)} />
          <CheckToggle checked={entireCell} label={homeText(activeLocale, 'entireCell')} onChange={(event) => setEntireCell(event.target.checked)} />
          <Select
            aria-label={homeText(activeLocale, 'scope')}
            sizeVariant="sm"
            className="w-36"
            value={scope}
            onChange={(event) => setScope(event.target.value as 'sheet' | 'workbook')}
          >
            <option value="sheet">{homeText(activeLocale, 'thisSheet')}</option>
            <option value="workbook">{homeText(activeLocale, 'allSheets')}</option>
          </Select>
        </Inline>
        {result !== null ? <Text size="sm" tone="subtle">{homeTemplate(activeLocale, 'replacedCount', { count: result })}</Text> : null}
      </Stack>
    </Dialog>
  );
}
