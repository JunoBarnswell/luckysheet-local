import { Button, DropdownMenu, Icon, Stack, Text } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem, WorkbookRole } from './types';

export interface WorkbookRowMenuProps {
  item: WorkbookCatalogItem;
  onOpen: (unitId: string) => void;
  onOpenInNewWindow: (unitId: string) => void;
  onExport: (unitId: string) => void;
  onSync: (unitId: string) => void;
  onRename: (unitId: string) => void;
  onCopy: (unitId: string) => void;
  onMove: (unitId: string) => void;
  onTrash: (unitId: string) => void;
  onRestore: (unitId: string) => void;
  onPurge: (unitId: string) => void;
  onFavorite: (unitId: string, favorite: boolean) => void;
  onShare: (unitId: string) => void;
}

const canEdit = (role: WorkbookRole) => role === 'owner' || role === 'editor';

export function WorkbookRowMenu({ item, onOpen, onOpenInNewWindow, onExport, onSync, onRename, onCopy, onMove, onTrash, onRestore, onPurge, onFavorite, onShare }: WorkbookRowMenuProps) {
  const isTrashed = item.lifecycle === 'trashed';
  const canManage = item.role === 'owner';
  const canRename = canEdit(item.role) && !isTrashed;
  const canMove = canEdit(item.role) && !isTrashed;
  const canTrash = canManage && !isTrashed;
  const canRestore = canManage && isTrashed;
  const canSync = item.storageLocation !== 'remote' && !isTrashed;

  return (
    <DropdownMenu
      align="right"
      trigger={<Button aria-label={`打开 ${item.name} 的更多操作`} icon="more-horizontal" iconOnly size="sm" variant="ghost" className="h-8 w-8 text-slate-500 hover:bg-brand-soft hover:text-brand-dark" />}
    >
      {({ close }) => (
        <Stack gap="none" className="min-w-[190px]">
          <Text className="px-2.5 py-1.5 text-[10px] font-semibold uppercase tracking-[0.12em] text-slate-400">文件操作</Text>
          {!isTrashed ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="folder-open" onClick={() => { onOpen(item.unitId); close(); }} size="sm" variant="ghost">打开</Button> : null}
          {!isTrashed ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="external-link" onClick={() => { onOpenInNewWindow(item.unitId); close(); }} size="sm" variant="ghost">在新窗口打开</Button> : null}
          {!isTrashed ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="download" onClick={() => { onExport(item.unitId); close(); }} size="sm" variant="ghost">导出副本</Button> : null}
          {canSync ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="cloud-check" onClick={() => { onSync(item.unitId); close(); }} size="sm" variant="ghost">同步到服务端</Button> : null}
          <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="star" onClick={() => { onFavorite(item.unitId, !item.favorite); close(); }} size="sm" variant="ghost">{item.favorite ? '取消星标' : '添加星标'}</Button>
          {!isTrashed ? <Stack gap="none" className="my-1 border-t border-slate-100 pt-1">
            {canRename ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="pencil" onClick={() => { onRename(item.unitId); close(); }} size="sm" variant="ghost">重命名</Button> : null}
            {canMove ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="folder" onClick={() => { onMove(item.unitId); close(); }} size="sm" variant="ghost">移动到</Button> : null}
            <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="copy" onClick={() => { onCopy(item.unitId); close(); }} size="sm" variant="ghost">创建副本</Button>
            {canManage ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="share" onClick={() => { onShare(item.unitId); close(); }} size="sm" variant="ghost">共享</Button> : null}
            {canTrash ? <Button className="w-full justify-start rounded-md px-2.5 text-xs text-rose-600 hover:bg-rose-50" icon="trash" onClick={() => { onTrash(item.unitId); close(); }} size="sm" variant="ghost">移到回收站</Button> : null}
          </Stack> : <Stack gap="none" className="my-1 border-t border-slate-100 pt-1">
            {canRestore ? <Button className="w-full justify-start rounded-md px-2.5 text-xs" icon="refresh" onClick={() => { onRestore(item.unitId); close(); }} size="sm" variant="ghost">恢复</Button> : null}
            {canManage ? <Button className="w-full justify-start rounded-md px-2.5 text-xs text-rose-600 hover:bg-rose-50" icon="trash" onClick={() => { onPurge(item.unitId); close(); }} size="sm" variant="ghost">永久删除</Button> : null}
          </Stack>}
          {item.role === 'viewer' ? <Text className="mt-1 flex items-center gap-1 border-t border-slate-100 px-2.5 pt-2 text-[10px] text-slate-400"><Icon name="lock" size="xs" /> 只读文件</Text> : null}
        </Stack>
      )}
    </DropdownMenu>
  );
}
