import React, { useEffect, useState } from 'react';
import { Button, CheckToggle, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { FindDialogMode } from '@react-sheets/spreadsheet-app';
import type { FindSearchOrder, FindSearchTarget, FindScope } from '@react-sheets/sheet-features';
import type { Locale } from '../../i18n';
import { homeTemplate, homeText, resolveHomeLocale } from '../home/home-localization';

export type FindSearchChoice = 'values' | 'formulas' | 'values-formulas' | 'notes' | 'comments' | 'all';

export interface FindDialogActionParams {
  query: string;
  replace?: string;
  searchOrder: FindSearchOrder;
  matchCase: boolean;
  entireCell: boolean;
  wildcard: boolean;
  scope: FindScope;
  targets: readonly FindSearchTarget[];
}

export interface FindReplaceDialogProps {
  open: boolean;
  initialFind?: string;
  mode?: FindDialogMode;
  locale?: Locale;
  onClose: () => void;
  onFindNext: (params: FindDialogActionParams) => number | Promise<number>;
  onFindPrevious: (params: FindDialogActionParams) => number | Promise<number>;
  onFindAll: (params: FindDialogActionParams) => number | Promise<number>;
  onReplace: (params: FindDialogActionParams) => number | Promise<number>;
  onReplaceAll: (params: FindDialogActionParams) => number | Promise<number>;
}

const TARGETS: Record<FindSearchChoice, readonly FindSearchTarget[]> = {
  values: ['values'],
  formulas: ['formulas'],
  'values-formulas': ['values', 'formulas'],
  notes: ['notes'],
  comments: ['comments'],
  all: ['values', 'formulas', 'notes', 'comments'],
};

/** Presentation-only form. Search planning, cursor ownership, and writes stay in WorkbookSession. */
export function FindReplaceDialog({
  initialFind = '',
  locale,
  mode = 'replace',
  open,
  onClose,
  onFindNext,
  onFindPrevious,
  onFindAll,
  onReplace,
  onReplaceAll,
}: FindReplaceDialogProps): React.ReactElement | null {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [entireCell, setEntireCell] = useState(false);
  const [wildcard, setWildcard] = useState(false);
  const [scope, setScope] = useState<FindScope>('sheet');
  const [searchOrder, setSearchOrder] = useState<FindSearchOrder>('rows');
  const [searchIn, setSearchIn] = useState<FindSearchChoice>('values-formulas');
  const [result, setResult] = useState<{ count: number; kind: 'found' | 'replaced' } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const activeLocale = resolveHomeLocale(locale);
  const actionParams = (): FindDialogActionParams => ({
    query: find,
    ...(mode === 'replace' ? { replace } : {}),
    searchOrder,
    matchCase,
    entireCell,
    wildcard,
    scope,
    targets: TARGETS[searchIn],
  });
  const runFind = (action: (params: FindDialogActionParams) => number | Promise<number>) => {
    void Promise.resolve(action(actionParams())).then((count) => { setError(null); setResult({ count, kind: 'found' }); }).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)));
  };
  const runReplace = (action: (params: FindDialogActionParams) => number | Promise<number>) => {
    void Promise.resolve(action(actionParams())).then((count) => { setError(null); setResult({ count, kind: 'replaced' }); }).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)));
  };

  useEffect(() => {
    if (!open) return;
    setFind(initialFind);
    setReplace('');
    setSearchOrder('rows');
    setResult(null);
    setError(null);
  }, [initialFind, open, mode]);

  const findDisabled = find.length === 0;
  const replaceDisabled = mode !== 'replace' || find.length === 0 || replace.length === 0;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, mode === 'find' ? 'findTitle' : 'replaceTitle')}
      description={homeText(activeLocale, mode === 'find' ? 'findDescription' : 'replaceDescription')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="find-replace-dialog"
      footer={(
        <Inline gap="sm" className="flex-wrap justify-end">
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button size="sm" variant="secondary" disabled={findDisabled} data-testid="find-previous" onClick={() => runFind(onFindPrevious)}>{homeText(activeLocale, 'findPrevious')}</Button>
          <Button size="sm" variant="secondary" disabled={findDisabled} data-testid="find-next" onClick={() => runFind(onFindNext)}>{homeText(activeLocale, 'findNext')}</Button>
          <Button size="sm" variant="secondary" disabled={findDisabled} data-testid="find-all" onClick={() => runFind(onFindAll)}>{homeText(activeLocale, 'findAll')}</Button>
          {mode === 'replace' ? <>
            <Button size="sm" variant="secondary" disabled={replaceDisabled} data-testid="replace-one" onClick={() => runReplace(onReplace)}>{homeText(activeLocale, 'replaceOne')}</Button>
            <Button size="sm" variant="primary" disabled={replaceDisabled} data-testid="find-replace-all" onClick={() => runReplace(onReplaceAll)}>{homeText(activeLocale, 'replaceAll')}</Button>
          </> : null}
        </Inline>
      )}
    >
      <Stack gap="sm">
        <TextInput
          aria-label={homeText(activeLocale, 'find')}
          placeholder={homeText(activeLocale, 'find')}
          value={find}
          autoFocus
          data-testid="find-input"
          onChange={(event) => { setFind(event.target.value); setResult(null); }}
        />
        {mode === 'replace' ? <TextInput
          aria-label={homeText(activeLocale, 'replaceWith')}
          placeholder={homeText(activeLocale, 'replaceWith')}
          value={replace}
          data-testid="replace-input"
          onChange={(event) => { setReplace(event.target.value); setResult(null); }}
        /> : null}
        <Inline gap="md" className="flex-wrap items-center">
          <CheckToggle checked={matchCase} label={homeText(activeLocale, 'matchCase')} onChange={(event) => { setMatchCase(event.target.checked); setResult(null); }} />
          <CheckToggle checked={entireCell} label={homeText(activeLocale, 'entireCell')} onChange={(event) => { setEntireCell(event.target.checked); setResult(null); }} />
          <CheckToggle checked={wildcard} label={homeText(activeLocale, 'wildcard')} onChange={(event) => { setWildcard(event.target.checked); setResult(null); }} />
        </Inline>
        <Inline gap="md" className="flex-wrap items-center">
          <Select aria-label={homeText(activeLocale, 'scope')} sizeVariant="sm" className="w-36" value={scope} onChange={(event) => setScope(event.target.value as FindScope)}>
            <option value="selection">{homeText(activeLocale, 'selection')}</option>
            <option value="sheet">{homeText(activeLocale, 'thisSheet')}</option>
            <option value="workbook">{homeText(activeLocale, 'allSheets')}</option>
          </Select>
          <Select aria-label={homeText(activeLocale, 'searchIn')} sizeVariant="sm" className="w-44" value={searchIn} onChange={(event) => setSearchIn(event.target.value as FindSearchChoice)}>
            <option value="values">{homeText(activeLocale, 'values')}</option>
            <option value="formulas">{homeText(activeLocale, 'formulas')}</option>
            <option value="values-formulas">{homeText(activeLocale, 'valuesAndFormulas')}</option>
            <option value="notes">{homeText(activeLocale, 'notes')}</option>
            <option value="comments">{homeText(activeLocale, 'comments')}</option>
            <option value="all">{homeText(activeLocale, 'allContent')}</option>
          </Select>
          <Select aria-label={homeText(activeLocale, 'searchOrder')} sizeVariant="sm" className="w-36" value={searchOrder} onChange={(event) => setSearchOrder(event.target.value as FindSearchOrder)}>
            <option value="rows">{homeText(activeLocale, 'byRows')}</option>
            <option value="columns">{homeText(activeLocale, 'byColumns')}</option>
          </Select>
        </Inline>
        {result ? <Text size="sm" tone="subtle">{result.kind === 'found' ? homeTemplate(activeLocale, 'foundCount', { count: result.count }) : homeTemplate(activeLocale, 'replacedCount', { count: result.count })}</Text> : null}
        {error ? <Text size="sm" tone="danger" role="alert">{error}</Text> : null}
      </Stack>
    </Dialog>
  );
}
