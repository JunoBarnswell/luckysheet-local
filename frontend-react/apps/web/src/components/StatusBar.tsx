import { Box, Button, Icon, Inline, Kbd, Text } from '@react-sheets/ui-system';
import type { AppPhase, SaveState } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';

export interface StatusBarProps {
  activeCell: string;
  locale: Locale;
  onOpenShortcuts: () => void;
  onZoomChange: (zoom: number) => void;
  phase: AppPhase;
  saveState: SaveState;
  sheetCount: number;
  zoom: number;
  collabStatus?: 'connecting' | 'open' | 'closed';
  pendingChangeSetCount?: number;
  collabRevision?: number;
  hasPendingOperations?: boolean;
}

export function StatusBar({
  activeCell,
  locale,
  onOpenShortcuts,
  onZoomChange,
  phase,
  saveState,
  sheetCount,
  zoom,
  collabStatus = 'closed',
  pendingChangeSetCount = 0,
  collabRevision = 0,
  hasPendingOperations = false,
}: StatusBarProps) {
  const disabled = phase !== 'ready';
  const status = (() => {
    switch (saveState) {
      case 'saved':
        return hasPendingOperations
          ? { icon: 'loader' as const, label: locale === 'zh-CN' ? '离线待同步' : 'Offline pending sync', tone: 'text-amber-300 animate-spin' }
          : { icon: 'cloud-check' as const, label: locale === 'zh-CN' ? '本地已保存' : 'Saved locally', tone: 'text-emerald-400' };
      case 'saving':
        return { icon: 'loader' as const, label: locale === 'zh-CN' ? '本地保存中' : 'Saving locally', tone: 'text-amber-300 animate-spin' };
      case 'offline':
        return { icon: 'cloud-check' as const, label: locale === 'zh-CN' ? '离线待同步' : 'Offline pending sync', tone: 'text-amber-300' };
      case 'syncing':
        return { icon: 'loader' as const, label: locale === 'zh-CN' ? '同步中' : 'Syncing', tone: 'text-amber-300 animate-spin' };
      case 'conflict':
        return { icon: 'alert-circle' as const, label: locale === 'zh-CN' ? '冲突' : 'Conflict', tone: 'text-rose-400' };
      case 'calculating':
        return { icon: 'loader' as const, label: locale === 'zh-CN' ? '计算中' : 'Calculating', tone: 'text-sky-300 animate-spin' };
      case 'error':
        return { icon: 'alert-circle' as const, label: locale === 'zh-CN' ? '错误' : 'Error', tone: 'text-rose-400' };
    }
  })();
  return (
    <Box aria-label="Workbook status bar" className="flex h-10 items-center justify-between gap-4 px-4">
      <Inline gap="md" className="min-w-0">
        <Inline gap="xs" className="shrink-0">
          <Icon name={status.icon} size="xs" className={status.tone} />
          <Text size="xs" tone="inverse">{status.label}</Text>
        </Inline>
        <Box className="hidden h-3 w-px bg-slate-700 sm:block" />
        <Text size="xs" tone="subtle" className="hidden sm:inline">{locale === 'zh-CN' ? '单元格' : 'Cell'} {activeCell}</Text>
        <Text size="xs" tone="subtle" className="hidden lg:inline">{sheetCount} {locale === 'zh-CN' ? '个工作表' : 'worksheets'}</Text>
        <Box className="hidden h-3 w-px bg-slate-700 lg:block" />
        <Text size="xs" tone="subtle" className="hidden lg:inline">
          {locale === 'zh-CN'
            ? `协同 ${collabStatus === 'open' ? '在线' : collabStatus === 'connecting' ? '连接中' : '离线'} · rev ${collabRevision}${pendingChangeSetCount > 0 ? ` · 待同步 ${pendingChangeSetCount}` : ''}`
            : `Collab ${collabStatus}${pendingChangeSetCount > 0 ? ` · ${pendingChangeSetCount} pending` : ''} · rev ${collabRevision}`}
        </Text>
      </Inline>
      <Inline gap="sm" className="shrink-0">
        <Button aria-label="Open keyboard shortcuts" disabled={disabled} icon="keyboard" onClick={onOpenShortcuts} size="xs" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white">
          <Text as="span" size="xs" className="hidden sm:inline">{locale === 'zh-CN' ? '快捷键' : 'Shortcuts'}</Text>
        </Button>
        <Kbd>⌘ /</Kbd>
        <Box className="hidden h-3 w-px bg-slate-700 sm:block" />
        <Button aria-label="Zoom out" disabled={disabled || zoom <= 75} icon="zoom-out" iconOnly onClick={() => onZoomChange(zoom - 5)} size="xs" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white" />
        <Text size="xs" tone="inverse" className="w-9 text-center tabular-nums">{zoom}%</Text>
        <Button aria-label="Zoom in" disabled={disabled || zoom >= 125} icon="zoom-in" iconOnly onClick={() => onZoomChange(zoom + 5)} size="xs" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white" />
      </Inline>
    </Box>
  );
}
