import { Box, Button, Inline, Text } from '@react-sheets/ui-system';
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
  void activeCell;
  void locale;
  void saveState;
  void sheetCount;
  void collabStatus;
  void pendingChangeSetCount;
  void collabRevision;
  void hasPendingOperations;
  return (
    <Box aria-label="Workbook status bar" className="relative flex h-[22px] items-center justify-between px-2">
      <Text size="xs" tone="inverse" className="!text-[#6ba78b] text-[10px] leading-none">就绪</Text>
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
