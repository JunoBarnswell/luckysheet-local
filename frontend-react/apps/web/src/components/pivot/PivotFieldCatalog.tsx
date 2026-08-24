import { Box, Button, CheckToggle, Inline, ScrollArea, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState, type DragEvent } from 'react';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';
import type { PivotFieldArea } from './pivot-contract';

export interface PivotFieldCatalogProps {
  fields: readonly PivotFieldDefinition[];
  selectedFieldIds: ReadonlySet<string>;
  locale: Locale;
  disabled?: boolean;
  onToggle: (fieldId: string, checked: boolean) => void;
  onDragField: (event: DragEvent<HTMLElement>, field: PivotFieldDefinition) => void;
  onKeyboardAssign: (fieldId: string, area: PivotFieldArea) => void;
}

export function PivotFieldCatalog({ disabled = false, fields, locale, onDragField, onKeyboardAssign, onToggle, selectedFieldIds }: PivotFieldCatalogProps) {
  const [query, setQuery] = useState('');
  const [expanded, setExpanded] = useState(false);
  const visibleFields = useMemo(() => fields.filter((field) => field.fieldId && field.name.toLowerCase().includes(query.toLowerCase())), [fields, query]);
  const allSelected = visibleFields.length > 0 && visibleFields.every((field) => selectedFieldIds.has(field.fieldId));
  return (
    <Stack gap="xs" className="min-h-0">
      <TextInput aria-label={pivotText(locale, 'search')} disabled={disabled} leadingIcon="search" placeholder={pivotText(locale, 'search')} value={query} onChange={(event) => setQuery(event.target.value)} />
      <Box className="h-[178px] overflow-hidden border border-[#d7d7d7] bg-white">
        <Inline gap="xs" className="h-8 border-b border-[#e2e2e2] px-2">
          <CheckToggle label={pivotText(locale, 'total')} checked={allSelected} disabled={disabled || visibleFields.length === 0} onChange={(event) => visibleFields.forEach((field) => onToggle(field.fieldId, event.target.checked))} />
          <Button aria-label={expanded ? 'Collapse fields' : 'Expand fields'} icon={expanded ? 'chevron-up' : 'chevron-down'} iconOnly size="xs" variant="ghost" className="ml-auto" onClick={() => setExpanded((value) => !value)} />
        </Inline>
        {expanded ? <ScrollArea className="h-[145px] px-2 py-1"><Stack gap="none">{visibleFields.length === 0 ? <Text size="xs" tone="subtle" className="py-3 text-center">{pivotText(locale, 'noMatches')}</Text> : visibleFields.map((field) => <Box key={field.fieldId} draggable={!disabled} className="cursor-grab py-0.5" onDragStart={(event) => onDragField(event, field)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onKeyboardAssign(field.fieldId, 'rows'); } }}><CheckToggle label={field.name} checked={selectedFieldIds.has(field.fieldId)} disabled={disabled} onChange={(event) => onToggle(field.fieldId, event.target.checked)} /></Box>)}</Stack></ScrollArea> : null}
      </Box>
    </Stack>
  );
}
