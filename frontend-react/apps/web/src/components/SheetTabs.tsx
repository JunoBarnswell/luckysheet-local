import React, { useRef, useState } from 'react';
import { Box, Button, ContextMenu, Dialog, DropdownMenu, Inline, Stack, Tab, Text, TextInput, type ContextMenuItem } from '@react-sheets/ui-system';
import type { CanvasSheetSnapshot, SheetDialogState } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';

export interface SheetTabsProps {
  activeSheetId: string;
  locale: Locale;
  disabled: boolean;
  sheets: CanvasSheetSnapshot[];
  onAdd: () => void;
  onSelect: (sheetId: string) => void;
  onRenameSheet?: (sheetId: string, name: string) => void;
  onDeleteSheet?: (sheetId: string) => void;
  onDuplicateSheet?: (sheetId: string) => void;
  onHideSheet?: (sheetId: string) => void;
  onSetTabColor?: (sheetId: string, color?: string) => void;
  onMoveSheet?: (sheetId: string, toIndex: number) => void;
  dialog: SheetDialogState | null;
  onOpenDialog: (dialog: SheetDialogState) => void;
  onUpdateDialog: (value: string) => void;
  onCloseDialog: () => void;
}

export function SheetTabs({
  activeSheetId,
  locale,
  disabled,
  sheets,
  onAdd,
  onSelect,
  onRenameSheet,
  onDeleteSheet,
  onDuplicateSheet,
  onHideSheet,
  onSetTabColor,
  onMoveSheet,
  dialog,
  onOpenDialog,
  onUpdateDialog,
  onCloseDialog,
}: SheetTabsProps) {
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    open: boolean;
    targetSheetId?: string;
  }>({ x: 0, y: 0, open: false });
  const tabsViewportRef = useRef<HTMLDivElement>(null);

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
      label: locale === 'zh-CN' ? '重命名' : 'Rename Sheet',
      icon: 'file-text',
      onSelect: () => {
        if (!contextMenu.targetSheetId) return;
        const currentName = sheets.find((s) => s.id === contextMenu.targetSheetId)?.name ?? '';
        onOpenDialog({ kind: 'rename', sheetId: contextMenu.targetSheetId, value: currentName });
      },
    },
    {
      id: 'duplicate',
      label: locale === 'zh-CN' ? '复制工作表' : 'Duplicate Sheet',
      icon: 'copy',
      onSelect: () => {
        if (contextMenu.targetSheetId) onDuplicateSheet?.(contextMenu.targetSheetId);
      },
    },
    {
      id: 'move-left',
      label: locale === 'zh-CN' ? '左移' : 'Move Left',
      icon: 'arrow-left',
      disabled: !contextMenu.targetSheetId || sheets.findIndex((s) => s.id === contextMenu.targetSheetId) <= 0,
      onSelect: () => {
        if (!contextMenu.targetSheetId) return;
        const index = sheets.findIndex((s) => s.id === contextMenu.targetSheetId);
        if (index > 0) onMoveSheet?.(contextMenu.targetSheetId, index - 1);
      },
    },
    {
      id: 'move-right',
      label: locale === 'zh-CN' ? '右移' : 'Move Right',
      icon: 'arrow-right',
      disabled: !contextMenu.targetSheetId || sheets.findIndex((s) => s.id === contextMenu.targetSheetId) >= sheets.length - 1,
      onSelect: () => {
        if (!contextMenu.targetSheetId) return;
        const index = sheets.findIndex((s) => s.id === contextMenu.targetSheetId);
        if (index >= 0 && index < sheets.length - 1) onMoveSheet?.(contextMenu.targetSheetId, index + 1);
      },
    },
    {
      id: 'tab-color',
      label: locale === 'zh-CN' ? '标签颜色' : 'Tab Color',
      icon: 'palette',
      onSelect: () => {
        if (!contextMenu.targetSheetId) return;
        onOpenDialog({ kind: 'tab-color', sheetId: contextMenu.targetSheetId, value: '#217345' });
      },
    },
    {
      id: 'hide',
      label: locale === 'zh-CN' ? '隐藏工作表' : 'Hide Sheet',
      icon: 'eye',
      disabled: sheets.filter((sheet) => !sheet.hidden).length <= 1,
      onSelect: () => {
        if (contextMenu.targetSheetId) onHideSheet?.(contextMenu.targetSheetId);
      },
    },
    {
      id: 'delete',
      label: locale === 'zh-CN' ? '删除工作表' : 'Delete Sheet',
      icon: 'trash',
      danger: true,
      disabled: sheets.length <= 1,
      onSelect: () => {
        if (!contextMenu.targetSheetId || sheets.length <= 1) return;
        onOpenDialog({ kind: 'delete', sheetId: contextMenu.targetSheetId, value: '' });
      },
    },
  ];

  return (
    <Box as="nav" aria-label={locale === 'zh-CN' ? '工作表' : 'Worksheets'} className="flex h-[29px] items-center gap-0 border-t border-[#d9d9d9] bg-white px-0">
      <Inline gap="none" className="h-full min-w-0 flex-1 items-center pl-2">
        <Button aria-label={locale === 'zh-CN' ? '向左滚动工作表' : 'Scroll worksheets left'} disabled={disabled} icon="chevron-left" iconOnly onClick={() => tabsViewportRef.current?.scrollBy({ left: -180, behavior: 'smooth' })} size="xs" variant="ghost" className="!h-7 !min-h-0 !w-7 rounded-none px-0 text-[#68736e]" />
        <Button aria-label={locale === 'zh-CN' ? '向右滚动工作表' : 'Scroll worksheets right'} disabled={disabled} icon="chevron-right" iconOnly onClick={() => tabsViewportRef.current?.scrollBy({ left: 180, behavior: 'smooth' })} size="xs" variant="ghost" className="!h-7 !min-h-0 !w-7 rounded-none px-0 text-[#68736e]" />
        <Button aria-label={locale === 'zh-CN' ? '添加工作表' : 'Add worksheet'} disabled={disabled} icon="plus" iconOnly onClick={onAdd} size="xs" variant="ghost" className="!h-7 !min-h-0 !w-7 rounded-none px-0 text-[#217345]" />
        <Box className="mx-3 h-5 w-px shrink-0 bg-[#d9d9d9]" />
        <Box ref={tabsViewportRef} className="h-full w-[580px] shrink-0 overflow-x-auto overflow-y-hidden [scrollbar-width:none] [&::-webkit-scrollbar]:hidden">
          <Inline gap="none" className="h-full min-w-max items-center">
            {sheets.filter((sheet) => !sheet.hidden).map((sheet) => (
              <Box key={sheet.id} onContextMenu={(e) => handleContextMenu(e, sheet.id)} className="inline-flex h-full">
                <Tab
                  active={sheet.id === activeSheetId}
                  disabled={disabled}
                  onClick={() => onSelect(sheet.id)}
                  className="!h-7 !min-h-0 !rounded-none border-b-2 border-transparent px-2 py-0 text-xs font-semibold aria-selected:!bg-[#e7feee] aria-selected:!text-[#217345]"
                  style={sheet.tabColor ? { borderBottomColor: sheet.tabColor, borderBottomWidth: 2 } : undefined}
                >
                  <Text as="span">{sheet.name}</Text>
                </Tab>
              </Box>
            ))}
          </Inline>
        </Box>
        <DropdownMenu
          align="right"
          disabled={disabled}
          trigger={<Button aria-label={locale === 'zh-CN' ? '更多工作表操作' : 'More worksheet actions'} icon="more-vertical" iconOnly size="xs" variant="ghost" className="!h-7 !min-h-0 !w-7 rounded-none px-0 text-[#68736e]" />}
        >
          {({ close }) => (
            <Stack gap="none" className="min-w-[13rem] p-1">
              <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onAdd(); }}>
                {locale === 'zh-CN' ? '新增工作表' : 'Add worksheet'}
              </Button>
              {sheets.filter((sheet) => !sheet.hidden).map((sheet) => (
                <Button key={sheet.id} size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onSelect(sheet.id); }}>
                  {sheet.name}
                </Button>
              ))}
            </Stack>
          )}
        </DropdownMenu>
        <Box aria-hidden="true" className="relative mx-3 h-2 min-w-0 flex-1">
          <Box className="absolute inset-x-0 top-1/2 h-px -translate-y-1/2 bg-[#c7c7c7]" />
          <Box className="absolute left-[38%] top-1/2 h-2 w-1 -translate-y-1/2 rounded-sm bg-[#8a8a8a]" />
        </Box>
      </Inline>

      <ContextMenu
        x={contextMenu.x}
        y={contextMenu.y}
        open={contextMenu.open}
        items={menuItems}
        onClose={() => setContextMenu((prev) => ({ ...prev, open: false }))}
      />
      <Dialog
        open={dialog !== null}
        title={dialog?.kind === 'rename' ? (locale === 'zh-CN' ? '重命名工作表' : 'Rename Sheet') : dialog?.kind === 'tab-color' ? (locale === 'zh-CN' ? '标签颜色' : 'Tab Color') : (locale === 'zh-CN' ? '删除工作表' : 'Delete Sheet')}
        onClose={onCloseDialog}
        maxWidth="sm"
      >
        {dialog?.kind === 'delete' ? (
          <Text size="sm">{locale === 'zh-CN' ? '确定删除当前工作表？此操作支持撤销。' : 'Delete this worksheet? The action can be undone.'}</Text>
        ) : (
          <Stack gap="xs">
            <Text size="xs" tone="muted">{dialog?.kind === 'rename' ? (locale === 'zh-CN' ? '输入新的工作表名称。' : 'Enter a new worksheet name.') : (locale === 'zh-CN' ? '输入 #RRGGBB 颜色，留空清除。' : 'Enter a #RRGGBB color, or leave empty to clear.')}</Text>
            <TextInput
              aria-label={dialog?.kind === 'rename' ? '工作表名称' : '标签颜色'}
              autoFocus
              value={dialog?.value ?? ''}
              onChange={(event) => onUpdateDialog(event.target.value)}
            />
          </Stack>
        )}
        <Inline gap="sm" className="mt-4 justify-end">
          <Button size="sm" variant="ghost" onClick={onCloseDialog}>{locale === 'zh-CN' ? '取消' : 'Cancel'}</Button>
          <Button
            size="sm"
            variant={dialog?.kind === 'delete' ? 'danger' : 'primary'}
            onClick={() => {
              if (!dialog) return;
              if (dialog.kind === 'rename' && dialog.value.trim()) onRenameSheet?.(dialog.sheetId, dialog.value.trim());
              if (dialog.kind === 'tab-color') onSetTabColor?.(dialog.sheetId, dialog.value.trim() || undefined);
              if (dialog.kind === 'delete') onDeleteSheet?.(dialog.sheetId);
              onCloseDialog();
            }}
          >
            {dialog?.kind === 'delete' ? (locale === 'zh-CN' ? '删除' : 'Delete') : (locale === 'zh-CN' ? '确定' : 'Apply')}
          </Button>
        </Inline>
      </Dialog>
    </Box>
  );
}
