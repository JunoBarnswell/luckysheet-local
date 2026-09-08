import { StatusBadge, type StatusBadgeKind } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem } from './types';

function resolveKind(item: WorkbookCatalogItem): StatusBadgeKind {
  if (item.lifecycle === 'trashed') return 'trashed';
  if (item.syncStatus === 'error') return 'error';
  if (item.syncStatus === 'conflict') return 'conflict';
  if (item.syncStatus === 'syncing') return 'syncing';
  if (item.role !== 'owner') return 'shared';
  return 'synced';
}

export function WorkbookStatusBadge({ item }: { item: WorkbookCatalogItem }) {
  const kind = resolveKind(item);
  const label = kind === 'synced'
    ? item.role === 'owner' ? '已保存 · 所有者' : '已保存 · 共享'
    : undefined;
  return <StatusBadge kind={kind} label={label} />;
}
