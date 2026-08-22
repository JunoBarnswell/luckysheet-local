import React, { useState } from 'react';
import { Box, Button, ContextMenu, Icon, Inline, Tab, Text, type ContextMenuItem } from '@react-sheets/ui-system';
import type { SheetView } from '../state/workspace';

export interface SheetTabsProps {
  activeSheetId: string;
  disabled: boolean;
  sheets: SheetView[];
  onAdd: () => void;
  onSelect: (sheetId: string) => void;
  onRenameSheet?: (sheetId: string, name: string) => void;
  onDeleteSheet?: (sheetId: string) => void;
}

export function SheetTabs({
  activeSheetId,
  disabled,
  sheets,
  onAdd,
  onSelect,
  onRenameSheet,
  onDeleteSheet,
}: SheetTabsProps) {
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    open: boolean;
    targetSheetId?: string;
  }>({ x: 0, y: 0, open: false });

  const handleContextMenu = (e: React.MouseEvent, sheetId: string) => {
    e.preventDefault();
    if (disabled) return;
    setContextMenu({
      x: e.clientX,
      y: e.clientY,
      open: true,
      targetSheetId: sheetId,
    });
  };

  const menuItems: ContextMenuItem[] = [
    {
      id: 'rename',
      label: 'Rename Sheet',
      icon: 'file-text',
      onSelect: () => {
        if (!contextMenu.targetSheetId) return;
        const currentName = sheets.find((s) => s.id === contextMenu.targetSheetId)?.name ?? '';
        const newName = prompt('Enter new sheet name:', currentName);
        if (newName && newName.trim()) {
          onRenameSheet?.(contextMenu.targetSheetId, newName.trim());
        }
      },
    },
    {
      id: 'delete',
      label: 'Delete Sheet',
      icon: 'trash',
      danger: true,
      disabled: sheets.length <= 1,
      onSelect: () => {
        if (!contextMenu.targetSheetId || sheets.length <= 1) return;
        if (confirm('Are you sure you want to delete this sheet?')) {
          onDeleteSheet?.(contextMenu.targetSheetId);
        }
      },
    },
  ];

  return (
    <Box as="nav" aria-label="Worksheets" className="flex h-10 items-center justify-between gap-4 border-t border-slate-200 bg-white px-3">
      <Inline gap="xs" className="min-w-0 overflow-x-auto">
        <Button
          aria-label="Add worksheet"
          disabled={disabled}
          icon="plus"
          iconOnly
          onClick={onAdd}
          size="sm"
          variant="soft"
          className="h-7 w-7"
        />
        <Box className="mx-1 h-4 w-px shrink-0 bg-slate-200" />
        {sheets.map((sheet) => (
          <div
            key={sheet.id}
            onContextMenu={(e) => handleContextMenu(e, sheet.id)}
            className="inline-flex"
          >
            <Tab
              active={sheet.id === activeSheetId}
              disabled={disabled}
              onClick={() => onSelect(sheet.id)}
              className="h-7 px-3 py-0 text-xs font-semibold"
            >
              <Inline gap="xs">
                <Icon name={sheet.isEmpty ? 'file-plus' : 'grid'} size="xs" />
                <span>{sheet.name}</span>
              </Inline>
            </Tab>
          </div>
        ))}
      </Inline>

      <Inline gap="sm" className="hidden shrink-0 md:flex">
        <Text size="xs" tone="subtle">{sheets.length} worksheets</Text>
        <Box className="h-4 w-px bg-slate-200" />
        <Text size="xs" tone="muted">Auto-saved to WAL SQLite</Text>
      </Inline>

      <ContextMenu
        x={contextMenu.x}
        y={contextMenu.y}
        open={contextMenu.open}
        items={menuItems}
        onClose={() => setContextMenu((prev) => ({ ...prev, open: false }))}
      />
    </Box>
  );
}
