import { Button, Icon, ScrollArea, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useMemo, useState, type DragEvent } from 'react';
import type { PivotFieldArea, PivotFieldDefinition } from './types';

export interface PivotFieldCatalogProps {
  fields: readonly PivotFieldDefinition[];
  selectedFieldIds: ReadonlySet<string>;
  disabled?: boolean;
  onToggle: (fieldId: string, checked: boolean) => void;
  onDragField: (event: DragEvent<HTMLElement>, field: PivotFieldDefinition) => void;
  onKeyboardAssign: (fieldId: string, area: PivotFieldArea) => void;
}

export function PivotFieldCatalog({ disabled = false, fields, onDragField, onKeyboardAssign, onToggle, selectedFieldIds }: PivotFieldCatalogProps) {
  const [query, setQuery] = useState('');
  const visibleFields = useMemo(() => fields.filter((field) => `${field.label} ${field.type}`.toLowerCase().includes(query.toLowerCase())), [fields, query]);
  return (
    <Stack gap="sm" className="min-h-0">
      <TextInput aria-label="Search pivot fields" disabled={disabled} leadingIcon="search" placeholder="Search fields" value={query} onChange={(event) => setQuery(event.target.value)} />
      <ScrollArea className="max-h-64 pr-1">
        <Stack gap="xs">
          {visibleFields.length === 0 ? <Text size="xs" tone="subtle" className="px-1 py-3">No matching fields</Text> : null}
          {visibleFields.map((field) => (
            <Button
              key={field.id}
              aria-pressed={selectedFieldIds.has(field.id)}
              disabled={disabled}
              icon={selectedFieldIds.has(field.id) ? 'check' : 'plus'}
              size="xs"
              variant={selectedFieldIds.has(field.id) ? 'soft' : 'ghost'}
              className="w-full justify-start text-left"
              onClick={() => onToggle(field.id, !selectedFieldIds.has(field.id))}
              onKeyDown={(event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                  event.preventDefault();
                  onKeyboardAssign(field.id, 'rows');
                }
              }}
              onDragStart={(event) => onDragField(event, field)}
              draggable={!disabled}
            >
              {field.label} · {field.type}
            </Button>
          ))}
        </Stack>
      </ScrollArea>
      <Text size="xs" tone="subtle" className="flex items-center gap-1"><Icon name="keyboard" size="xs" /> Enter assigns selected fields to ROWS; use each chip menu to move/remove.</Text>
    </Stack>
  );
}
