import type { ReactNode } from 'react';
import { Button } from './Button';
import { Box, Inline, Text } from './layout';
import { Icon } from './Icon';
import type { SaveState } from './shell-types';

export interface DocumentBarProps {
  workbookName: string;
  saveState: SaveState;
  onSave?: () => void;
  onSearch?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onShare?: () => void;
  onComments?: () => void;
  children?: ReactNode;
}

function saveLabel(saveState: SaveState): string {
  if (saveState === 'saved') return '已保存';
  if (saveState === 'saving' || saveState === 'syncing' || saveState === 'calculating') return '正在保存';
  if (saveState === 'offline') return '离线';
  if (saveState === 'conflict') return '存在冲突';
  return '保存失败';
}

/** Fixed 36px editor document bar. It owns chrome composition, never workbook state. */
export function DocumentBar({ workbookName, saveState, onSave, onSearch, onUndo, onRedo, onShare, onComments, children }: DocumentBarProps) {
  const status = saveLabel(saveState);
  const statusTone = saveState === 'saved' ? 'text-[#107C41]' : saveState === 'error' || saveState === 'conflict' ? 'text-rose-600' : 'text-[#6B6B6B]';
  return (
    <Box as="header" className="flex h-[36px] shrink-0 items-center justify-between border-b border-[#D1D1D1] bg-white px-3 text-[13px]" data-testid="document-bar">
      <Inline gap="sm" className="min-w-0">
        <Box aria-label="Cloud Sheets" className="flex size-6 shrink-0 items-center justify-center rounded-[3px] bg-[#107C41] text-white" role="img">
          <svg aria-hidden="true" className="size-4" fill="none" viewBox="0 0 20 20"><rect x="2.5" y="2.5" width="15" height="15" rx="1.5" stroke="currentColor" strokeWidth="1.5" /><path d="M2.5 7h15M7 2.5v15M12 7v10.5" stroke="currentColor" strokeWidth="1.2" /><path d="M12 10h5.5" stroke="currentColor" strokeWidth="1.2" /></svg>
        </Box>
        <Text className="truncate font-semibold text-[#242424]" title={workbookName}>{workbookName}</Text>
        <Button aria-label={`保存状态：${status}`} icon="cloud-check" iconOnly onClick={onSave} size="xs" variant="ghost" className={`!h-7 !w-7 ${statusTone}`} title={status} />
        <Text size="xs" tone="muted" className="hidden whitespace-nowrap sm:inline">{status}</Text>
      </Inline>
      <Inline gap="xs" className="shrink-0">
        <Button aria-label="撤销" icon="undo" iconOnly onClick={onUndo} size="xs" variant="ghost" className="!h-7 !w-7" />
        <Button aria-label="重做" icon="redo" iconOnly onClick={onRedo} size="xs" variant="ghost" className="!h-7 !w-7" />
        <Button aria-label="搜索" icon="search" iconOnly onClick={onSearch} size="xs" variant="ghost" className="!h-7 !w-7" />
        <Button aria-label="协作" icon="users" iconOnly size="xs" variant="ghost" className="!h-7 !w-7" />
        <Button aria-label="评论" icon="comment" iconOnly onClick={onComments} size="xs" variant="ghost" className="!h-7 !w-7" />
        {children}
        <Button aria-label="分享" icon="share" onClick={onShare} size="xs" variant="brand" className="!h-7 rounded-[3px] px-3">分享</Button>
      </Inline>
    </Box>
  );
}
