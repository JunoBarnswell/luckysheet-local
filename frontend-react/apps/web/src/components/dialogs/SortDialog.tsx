import React, { useEffect, useState } from 'react';
import { Button, CheckToggle, Dialog, Inline, Select, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeTemplate, homeText, resolveHomeLocale } from '../home/home-localization';

export interface SortCriterionInput {
  colIdx: number;
  ascending: boolean;
}

export interface SortDialogProps {
  open: boolean;
  columns: string[];
  locale?: Locale;
  onClose: () => void;
  onSort: (criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader: boolean) => void;
}

/** Sort UI only. The caller owns resolving the current selection into a command. */
export function SortDialog({ open, columns, locale, onClose, onSort }: SortDialogProps): React.ReactElement | null {
  const [criteria, setCriteria] = useState<Array<{ colIdx: number; ascending: boolean }>>([{ colIdx: 0, ascending: true }]);
  const [hasHeader, setHasHeader] = useState(false);
  const activeLocale = resolveHomeLocale(locale);

  useEffect(() => {
    if (!open) return;
    setCriteria([{ colIdx: 0, ascending: true }]);
    setHasHeader(false);
  }, [open]);

  const canSort = columns.length > 0 && criteria.length > 0;

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'sortRange')}
      description={homeText(activeLocale, 'sortDescription')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="sort-dialog"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button
            size="sm"
            variant="primary"
            disabled={!canSort}
            data-testid="sort-apply"
            onClick={() => {
              onClose();
              onSort(criteria, hasHeader);
            }}
          >
            {homeText(activeLocale, 'sort')}
          </Button>
        </>
      )}
    >
      <Stack gap="sm">
        {columns.length === 0 ? <Text size="sm" tone="subtle">{homeText(activeLocale, 'noColumns')}</Text> : null}
        {criteria.map((criterion, index) => (
          <Inline key={`${criterion.colIdx}-${index}`} gap="sm" className="items-center">
            <Select
              aria-label={homeTemplate(activeLocale, 'sortColumn', { index: index + 1 })}
              sizeVariant="sm"
              className="flex-1"
              disabled={columns.length === 0}
              value={String(criterion.colIdx)}
              onChange={(event) => {
                const next = [...criteria];
                next[index] = { ...criterion, colIdx: Number(event.target.value) };
                setCriteria(next);
              }}
            >
              {columns.map((label, columnIndex) => (
                <option key={`${label}-${columnIndex}`} value={columnIndex}>{label}</option>
              ))}
            </Select>
            <Select
              aria-label={homeTemplate(activeLocale, 'sortOrder', { index: index + 1 })}
              sizeVariant="sm"
              className="w-32"
              disabled={columns.length === 0}
              value={criterion.ascending ? 'asc' : 'desc'}
              onChange={(event) => {
                const next = [...criteria];
                next[index] = { ...criterion, ascending: event.target.value === 'asc' };
                setCriteria(next);
              }}
            >
              <option value="asc">{homeText(activeLocale, 'ascending')}</option>
              <option value="desc">{homeText(activeLocale, 'descending')}</option>
            </Select>
            {criteria.length > 1 ? (
              <Button
                size="sm"
                variant="ghost"
                icon="x"
                iconOnly
                aria-label={homeText(activeLocale, 'close')}
                onClick={() => setCriteria(criteria.filter((_, criterionIndex) => criterionIndex !== index))}
              />
            ) : null}
          </Inline>
        ))}
        <Button
          size="sm"
          variant="ghost"
          icon="plus"
          className="self-start"
          disabled={columns.length === 0}
          onClick={() => setCriteria([...criteria, { colIdx: 0, ascending: true }])}
        >
          {homeText(activeLocale, 'addSortColumn')}
        </Button>
        <CheckToggle checked={hasHeader} label={homeText(activeLocale, 'dataHasHeader')} onChange={(event) => setHasHeader(event.target.checked)} />
      </Stack>
    </Dialog>
  );
}
