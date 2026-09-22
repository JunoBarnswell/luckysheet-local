import type { WorkbookCatalogEntry, WorkbookCatalogQuery, WorkbookSyncState } from './types';

/** Lower values sort first: actionable failures must remain visible. */
export const WORKBOOK_SYNC_STATE_PRIORITY: Readonly<Record<WorkbookSyncState, number>> = {
  error: 0,
  conflict: 1,
  syncing: 2,
  pending: 3,
  offline: 4,
  synced: 5,
};

export function compareWorkbookSyncState(left: WorkbookSyncState, right: WorkbookSyncState): number {
  return WORKBOOK_SYNC_STATE_PRIORITY[left] - WORKBOOK_SYNC_STATE_PRIORITY[right];
}

export function sortWorkbookCatalog(entries: readonly WorkbookCatalogEntry[], view: WorkbookCatalogQuery['view'] = 'all'): WorkbookCatalogEntry[] {
  return [...entries].sort((left, right) => {
    if (view === 'recent') {
      const rightOpened = Date.parse(right.lastOpenedAt ?? right.updatedAt);
      const leftOpened = Date.parse(left.lastOpenedAt ?? left.updatedAt);
      if (Number.isFinite(rightOpened) && Number.isFinite(leftOpened) && rightOpened !== leftOpened) return rightOpened - leftOpened;
    }
    if (left.favorite !== right.favorite) return left.favorite ? -1 : 1;
    const stateOrder = compareWorkbookSyncState(left.syncState, right.syncState);
    if (stateOrder !== 0) return stateOrder;
    const updatedOrder = Date.parse(right.updatedAt) - Date.parse(left.updatedAt);
    if (Number.isFinite(updatedOrder) && updatedOrder !== 0) return updatedOrder;
    return left.name.localeCompare(right.name, 'zh-Hans-CN');
  });
}

export function filterWorkbookCatalog(
  entries: readonly WorkbookCatalogEntry[],
  query: WorkbookCatalogQuery = {},
): WorkbookCatalogEntry[] {
  const text = query.query?.trim().toLocaleLowerCase();
  const filtered = entries.filter((entry) => {
    if (query.view === 'trash' && entry.lifecycle !== 'trashed') return false;
    if (query.view !== 'trash' && entry.lifecycle === 'trashed') return false;
    if (query.view === 'local' && entry.storage === 'remote') return false;
    if (query.view === 'owned' && entry.role !== 'owner') return false;
    if (query.view === 'shared' && !['commenter', 'editor', 'viewer'].includes(entry.role)) return false;
    if (query.spaceId && entry.spaceId !== query.spaceId) return false;
    if (query.folderId && entry.folderId !== query.folderId) return false;
    if (!text) return true;
    return [entry.name, entry.ownerName, entry.spaceName, ...entry.locationPath]
      .filter((value): value is string => Boolean(value))
      .some((value) => value.toLocaleLowerCase().includes(text));
  });
  return sortWorkbookCatalog(filtered, query.view);
}

export function resolveWorkbookSyncState(input: {
  syncState?: WorkbookSyncState;
  storage: 'local' | 'remote' | 'mirrored';
  pendingOperationCount?: number;
  syncMode?: 'remote' | 'local-only';
  remoteAvailable?: boolean;
}): WorkbookSyncState {
  if (input.syncState) return input.syncState;
  if (input.pendingOperationCount && input.pendingOperationCount > 0) return 'pending';
  if (input.syncMode === 'local-only' && input.storage !== 'local') return 'offline';
  if (input.storage === 'local') return input.remoteAvailable === false ? 'offline' : 'synced';
  return input.remoteAvailable === false ? 'offline' : 'synced';
}
