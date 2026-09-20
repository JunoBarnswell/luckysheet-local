import { Box, Button, Inline, Text } from '@react-sheets/ui-system';
import { useCellEdit, type AppPhase, type CellEditController, type SaveState } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';

export interface StatusBarProps {
  activeCell: string;
  cellEdit: CellEditController;
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
  fixedDecimalPlaces?: number | null;
}

export function StatusBar({
  activeCell,
  cellEdit,
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
  fixedDecimalPlaces = null,
}: StatusBarProps) {
  const edit = useCellEdit(cellEdit);
  const disabled = phase !== 'ready';
  const saveLabels: Record<SaveState, string> = locale === 'zh-CN'
    ? { saved: '已保存', saving: '正在保存', offline: '已断线 · 编辑暂停', syncing: '正在核对服务器版本', conflict: '存在冲突 · 草稿已保留', calculating: '计算中', error: '保存失败' }
    : { saved: 'Saved', saving: 'Saving', offline: 'Offline · Editing paused', syncing: 'Reconciling', conflict: 'Conflict · Draft retained', calculating: 'Calculating', error: 'Save failed' };
  const persistenceText = phase === 'error' ? saveLabels.error
    : collabStatus !== 'open' ? saveLabels.offline
    : hasPendingOperations || pendingChangeSetCount > 0 ? saveLabels[saveState === 'saved' ? 'saving' : saveState]
    : saveLabels[saveState];
  const labels = locale === 'zh-CN'
    ? { ready: '就绪', enter: '输入', edit: '编辑', point: '点选', overtype: '覆盖' }
    : { ready: 'Ready', enter: 'Enter', edit: 'Edit', point: 'Point', overtype: 'Overtype' };
  const statusText = `${labels[edit.status]}${edit.session?.overtype ? ` / ${labels.overtype}` : ''}`;
  return (
    <Box aria-label="Workbook status bar" className="relative flex h-[22px] items-center justify-between px-2">
      <Inline gap="sm">
        <Text size="xs" tone="inverse" className="text-[11px] leading-none">{statusText}</Text>
        <Text role="status" aria-live="polite" size="xs" tone="inverse" className="border-l border-white/30 pl-3 text-[11px] leading-none">{persistenceText}</Text>
        <Text size="xs" tone="inverse" className="text-[10px] opacity-75">{activeCell} · r{collabRevision}</Text>
        {fixedDecimalPlaces !== null ? <Text size="xs" tone="inverse" className="text-[10px] leading-none">Fixed Decimal: {fixedDecimalPlaces}</Text> : null}
      </Inline>
      <Button aria-label="Open keyboard shortcuts" disabled={disabled} className="sr-only" onClick={onOpenShortcuts}>快捷键</Button>
      <Inline gap="none" className="h-full shrink-0 items-center">
        <Button aria-label="Zoom out" disabled={disabled || zoom <= 75} onClick={() => onZoomChange(zoom - 5)} size="xs" variant="ghost" className="!h-5 !min-h-0 !w-6 rounded-none px-0 text-white hover:bg-emerald-700 hover:text-white">−</Button>
        <Box aria-hidden="true" className="relative mx-2 h-px w-[92px] bg-white/60">
          <Box className="absolute left-1/2 top-1/2 h-2 w-2 -translate-x-1/2 -translate-y-1/2 rounded-full bg-white" />
        </Box>
        <Button aria-label="Zoom in" disabled={disabled || zoom >= 125} onClick={() => onZoomChange(zoom + 5)} size="xs" variant="ghost" className="!h-5 !min-h-0 !w-6 rounded-none px-0 text-white hover:bg-emerald-700 hover:text-white">+</Button>
        <Text size="xs" tone="inverse" className="ml-1 w-9 text-center text-[10px] tabular-nums">{zoom}%</Text>
      </Inline>
    </Box>
  );
}
