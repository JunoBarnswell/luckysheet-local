import { Box, Button, Icon, Inline, Tab, Text } from '@react-sheets/ui-system';
import type { SheetView } from '../state/workspace';

export interface SheetTabsProps {
  activeSheetId: string;
  disabled: boolean;
  onAdd: () => void;
  onSelect: (sheetId: string) => void;
  sheets: SheetView[];
}

export function SheetTabs({ activeSheetId, disabled, onAdd, onSelect, sheets }: SheetTabsProps) {
  return (
    <Box as="nav" aria-label="Worksheets" className="flex h-11 items-center justify-between gap-4 px-4">
      <Inline gap="xs" className="min-w-0 overflow-x-auto">
        <Button aria-label="Add worksheet" disabled={disabled} icon="plus" iconOnly onClick={onAdd} size="sm" variant="soft" />
        <Box className="mx-1 h-5 w-px shrink-0 bg-line" />
        {sheets.map((sheet) => (
          <Tab key={sheet.id} active={sheet.id === activeSheetId} disabled={disabled} onClick={() => onSelect(sheet.id)}>
            <Inline gap="xs">
              <Icon name={sheet.isEmpty ? 'file-plus' : 'grid'} size="xs" />
              <Text size="xs" weight={sheet.id === activeSheetId ? 'semibold' : 'medium'}>{sheet.name}</Text>
            </Inline>
          </Tab>
        ))}
        <Button aria-label="More worksheet actions" disabled={disabled} icon="more-horizontal" iconOnly size="sm" variant="ghost" />
      </Inline>
      <Inline gap="sm" className="hidden shrink-0 md:flex">
        <Text size="xs" tone="subtle">{sheets.length} sheets</Text>
        <Box className="h-5 w-px bg-line" />
        <Text size="xs" tone="muted">All changes synced</Text>
      </Inline>
    </Box>
  );
}
