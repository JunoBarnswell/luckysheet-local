import { StatusBadge, type StatusBadgeKind } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem } from './types';

function resolveKind(item: WorkbookCatalogItem): StatusBadgeKind {
  if (item.lifecycle === 'trashed') return 'trashed';
  if (item.syncStatus === 'error') return 'error';
  if (item.syncStatus === 'conflict') return 'conflict';
  if (item.syncStatus === 'offline') return 'offline';
  if (item.syncStatus === 'syncing') return 'syncing';
  if (item.syncStatus === 'pending') return 'pending';
  if (item.storageLocation === 'local') return 'local';
  if (item.role !== 'owner') return 'shared';
  return 'synced';
}

export function WorkbookStatusBadge({ item }: { item: WorkbookCatalogItem }) {
  return <StatusBadge kind={resolveKind(item)} />;
}
