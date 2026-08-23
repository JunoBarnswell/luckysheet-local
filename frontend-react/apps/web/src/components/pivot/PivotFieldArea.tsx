import { Box, Button, CheckToggle, DropdownMenu, FieldDropZone, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { DragEvent } from 'react';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { PivotFieldArea as Area, PivotSortDirection } from './pivot-contract';

interface AreaItem extends PivotFieldDefinition {
  index: number;
}

export interface PivotFieldAreaProps {
  area: Area;
  fields: readonly PivotFieldDefinition[];
  fieldIds: readonly string[];
  disabled?: boolean;
  onDrop: (event: DragEvent<HTMLElement>, index?: number) => void;
  onRemove: (fieldId: string, index: number) => void;
  onMoveByKeyboard: (fieldId: string, index: number, direction: -1 | 1) => void;
  filterSelections?: Readonly<Record<string, string[]>>;
  onFilter?: (fieldId: string, selectedValues: string[]) => void;
  onSort?: (fieldId: string, direction: PivotSortDirection) => void;
  onGroup?: (fieldId: string, grouped: boolean) => void;
}

const labels: Record<Area, string> = { filters: 'FILTERS', columns: 'COLUMNS', rows: 'ROWS', values: 'VALUES' };
const icons: Record<Area, 'filter' | 'columns' | 'rows' | 'calculator'> = { filters: 'filter', columns: 'columns', rows: 'rows', values: 'calculator' };

export function PivotFieldArea({ area, disabled = false, fieldIds, fields, filterSelections = {}, onDrop, onFilter, onGroup, onMoveByKeyboard, onRemove, onSort }: PivotFieldAreaProps) {
  const items: AreaItem[] = fieldIds.map((id, index) => ({ ...(fields.find((field) => field.id === id) ?? { id, name: id, dataType: 'text' as const, ordinal: index }), index }));
  return (
    <Box as="section" aria-label={`${labels[area]} field area`} className="min-w-0">
      <Inline gap="xs" className="mb-1.5">
        <Icon name={icons[area]} size="xs" className="text-accent" />
        <Text size="xs" weight="semibold" tone="muted">{labels[area]}</Text>
        <Text size="xs" tone="subtle">{items.length}</Text>
      </Inline>
      <FieldDropZone
        disabled={disabled}
        emptyLabel={`Drop fields here for ${labels[area].toLowerCase()}`}
        items={items}
        onDropItem={onDrop}
        renderItem={(field) => (
          <Inline
            draggable={!disabled}
            gap="xs"
            className="group min-h-8 cursor-grab rounded-md border border-blue-100 bg-white px-2 py-1 shadow-sm active:cursor-grabbing"
            onDragStart={(event) => event.dataTransfer.setData('application/x-pivot-field', field.id)}
          >
            <Icon name="menu" size="xs" className="text-slate-300" />
            <Text size="xs" weight="medium" className="min-w-0 flex-1 truncate">{field.name}</Text>
            <DropdownMenu
              align="right"
              trigger={<Button aria-label={`Keyboard menu for ${field.name}`} icon="more-horizontal" iconOnly size="xs" variant="ghost" />}
            >
              {({ close }) => (
                <Inline gap="xs" className="p-1">
                   <Stack gap="xs" className="min-w-48 p-1">
                     <Inline gap="xs"><Button disabled={field.index === 0} icon="arrow-up" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.id, field.index, -1); close(); }} /><Button disabled={field.index === items.length - 1} icon="arrow-down" iconOnly size="xs" variant="ghost" onClick={() => { onMoveByKeyboard(field.id, field.index, 1); close(); }} /><Button icon="trash" iconOnly size="xs" variant="danger" onClick={() => { onRemove(field.id, field.index); close(); }} /></Inline>
                     {onSort ? <Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => { onSort(field.id, 'ascending'); close(); }}>Sort A–Z</Button><Button size="xs" variant="ghost" onClick={() => { onSort(field.id, 'descending'); close(); }}>Sort Z–A</Button></Inline> : null}
                     {onGroup ? <Button size="xs" variant="ghost" onClick={() => { onGroup(field.id, true); close(); }}>Group field</Button> : null}
                     {onFilter && field.values?.length ? <Stack gap="xs" className="border-t border-slate-100 pt-1"><Text size="xs" weight="semibold">Filter values</Text>{field.values.map((value) => { const textValue = String(value); const selected = filterSelections[field.id] ?? field.values?.map(String) ?? []; return <CheckToggle key={textValue} label={textValue} checked={selected.includes(textValue)} onChange={(event) => onFilter(field.id, event.target.checked ? [...selected, textValue] : selected.filter((item) => item !== textValue))} />; })}</Stack> : null}
                   </Stack>
                </Inline>
              )}
            </DropdownMenu>
          </Inline>
        )}
        onDragOverItem={(event) => { event.dataTransfer.dropEffect = 'move'; }}
      />
    </Box>
  );
}
