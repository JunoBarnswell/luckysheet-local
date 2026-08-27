import { Box, Button, CheckToggle, DropdownMenu, Icon, Inline, ScrollArea, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState, type DragEvent } from 'react';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';
import { PIVOT_FIELD_AREAS, type PivotFieldArea } from './pivot-contract';

export interface PivotFieldCatalogProps {
  fields: readonly PivotFieldDefinition[];
  selectedFieldIds: ReadonlySet<string>;
  locale: Locale;
  disabled?: boolean;
  onToggle: (fieldId: string, checked: boolean) => void;
  onToggleVisible: (fieldIds: readonly string[], checked: boolean) => void;
  onDragField: (event: DragEvent<HTMLElement>, field: PivotFieldDefinition) => void;
  onKeyboardAssign: (field: PivotFieldDefinition) => void;
  onAssignField: (field: PivotFieldDefinition, area: PivotFieldArea) => void;
  className?: string;
}

export function PivotFieldCatalog({ className, disabled = false, fields, locale, onAssignField, onDragField, onKeyboardAssign, onToggle, onToggleVisible, selectedFieldIds }: PivotFieldCatalogProps) {
  const [query, setQuery] = useState('');
  const [expanded, setExpanded] = useState(true);
  const visibleFields = useMemo(() => fields.filter((field) => field.fieldId && field.name.toLowerCase().includes(query.toLowerCase())), [fields, query]);
  const allSelected = visibleFields.length > 0 && visibleFields.every((field) => selectedFieldIds.has(field.fieldId));
  return (
    <Stack gap="xs" className={`min-h-0 min-w-0 overflow-hidden ${className ?? ''}`}>
      <TextInput aria-label={pivotText(locale, 'search')} disabled={disabled} leadingIcon="search" placeholder={pivotText(locale, 'search')} value={query} onChange={(event) => setQuery(event.target.value)} />
      <Box className="flex min-h-0 flex-1 flex-col overflow-hidden border border-[#d7d7d7] bg-white">
        <Inline gap="xs" className="h-8 border-b border-[#e2e2e2] px-2">
          <CheckToggle label={pivotText(locale, 'total')} checked={allSelected} disabled={disabled || visibleFields.length === 0} onChange={(event) => onToggleVisible(visibleFields.map((field) => field.fieldId), event.target.checked)} />
          <Button aria-label={pivotText(locale, expanded ? 'collapseFields' : 'expandFields')} icon={expanded ? 'chevron-up' : 'chevron-down'} iconOnly size="xs" variant="ghost" className="ml-auto" onClick={() => setExpanded((value) => !value)} />
        </Inline>
        {expanded ? <ScrollArea className="min-h-0 flex-1 p-1"><Stack gap="none">{visibleFields.length === 0 ? <Text size="xs" tone="subtle" className="py-3 text-center">{pivotText(locale, 'noMatches')}</Text> : visibleFields.map((field) => <Inline key={field.fieldId} draggable={!disabled} gap="xs" className="group min-h-8 cursor-grab px-1 hover:bg-[#f3f6fb] active:cursor-grabbing" onDragStart={(event) => onDragField(event, field)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onKeyboardAssign(field); } }}><Icon name="menu" size="xs" className="shrink-0 text-slate-300 group-hover:text-accent" /><CheckToggle className="min-w-0 flex-1" label={field.name} checked={selectedFieldIds.has(field.fieldId)} disabled={disabled} onChange={(event) => onToggle(field.fieldId, event.target.checked)} /><DropdownMenu align="right" trigger={<Button aria-label={`${pivotText(locale, 'fieldMenu')}: ${field.name}`} icon="chevron-down" iconOnly size="xs" variant="ghost" disabled={disabled} />}>
          {({ close }) => <Stack gap="none" className="min-w-40 p-1"><Text size="xs" weight="semibold" className="px-2 py-1">{pivotText(locale, 'assignTo')}</Text>{PIVOT_FIELD_AREAS.map((area) => <Button key={area} size="xs" variant="ghost" className="w-full justify-start" onClick={() => { onAssignField(field, area); close(); }}>{pivotText(locale, area)}</Button>)}</Stack>}
        </DropdownMenu></Inline>)}</Stack></ScrollArea> : null}
      </Box>
    </Stack>
  );
}
