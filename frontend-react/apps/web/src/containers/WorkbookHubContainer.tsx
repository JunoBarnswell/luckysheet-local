import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, CheckToggle, Dialog, LocationPicker, Select, Stack, Text, TextInput, type LocationOption } from '@react-sheets/ui-system';
import {
  CreateWorkbookDialog,
  DeleteWorkbookDialog,
  ImportWorkbookDialog,
  RenameWorkbookDialog,
  ShareWorkbookDialog,
  WorkbookHubPage,
  type WorkbookCatalogItem,
  type WorkbookCategoryTab,
  type WorkbookHubSection,
  type WorkbookTemplateKind,
} from '../workbooks';
import { useApplicationServices } from '../ApplicationServicesProvider';
import { useAuthSession, useAuthSnapshot } from '../auth/AuthProvider';
import {
  createTemplateSnapshot,
  createWorkbookUnitId,
  type WorkbookCatalogEntry,
  type WorkbookTemplateId,
} from '@react-sheets/spreadsheet-app';
import type { SpaceMember, WorkspaceFolder, WorkspaceSpace } from '@react-sheets/protocol';
import type { UserPreferences } from '@react-sheets/protocol';

type ActiveDialog = 'create' | 'help' | 'import' | 'move' | 'options' | 'purge' | 'rename' | 'share' | 'trash' | null;

interface FolderLocation {
  folder: WorkspaceFolder;
  label: string;
}

interface WorkbookHubContainerProps {
  onOpenWorkbook: (unitId: string) => void;
}

function itemFromEntry(entry: WorkbookCatalogEntry): WorkbookCatalogItem {
  const locationLabel = entry.locationPath.length > 0
    ? entry.locationPath.join(' › ')
    : entry.storage === 'local'
      ? '本地设备'
      : entry.role === 'owner'
        ? '服务器 / 云端'
        : `共享给我${entry.ownerName ? ` · ${entry.ownerName}` : ''}`;
  return {
    unitId: entry.unitId,
    name: entry.name,
    updatedAt: entry.updatedAt,
    locationLabel,
    storageLocation: entry.storage,
    syncStatus: entry.syncState,
    lifecycle: entry.lifecycle,
    role: entry.role,
    sourceKind: entry.source,
    ownerName: entry.ownerName,
    ownerSubject: entry.ownerId,
    folderId: entry.folderId,
    folderPath: entry.locationPath,
    favorite: entry.favorite,
    revision: entry.revision,
    localRevision: entry.localRecord?.localRevision,
    serverRevision: entry.localRecord?.serverRevision,
    pendingOperationCount: entry.pendingOperationCount,
    sourceFileName: entry.sourceFileName,
  };
}

function destinationFromLocation(locationId: string): { destination: 'local' | 'remote'; folderId?: string; spaceId?: string } {
  if (locationId === 'local') return { destination: 'local' };
  if (locationId.startsWith('folder:')) {
    const [, spaceId, folderId] = locationId.split(':');
    if (!spaceId || !folderId) throw new Error('工作簿文件夹位置无效');
    return { destination: 'remote', spaceId, folderId };
  }
  if (locationId.startsWith('space:')) return { destination: 'remote', spaceId: locationId.slice('space:'.length) };
  throw new Error('请选择有效的工作簿保存位置');
}

function folderLocations(folders: readonly WorkspaceFolder[], spaces: readonly WorkspaceSpace[]): readonly FolderLocation[] {
  const byId = new Map(folders.map((folder) => [folder.folderId, folder]));
  const spaceName = new Map(spaces.map((space) => [space.spaceId, space.name]));
  const resolveLabel = (folder: WorkspaceFolder, visited = new Set<string>()): string => {
    if (visited.has(folder.folderId)) return folder.name;
    visited.add(folder.folderId);
    const parent = folder.parentFolderId ? byId.get(folder.parentFolderId) : undefined;
    return parent ? `${resolveLabel(parent, visited)} › ${folder.name}` : `${spaceName.get(folder.spaceId) ?? '云端'} › ${folder.name}`;
  };
  return folders.map((folder) => ({ folder, label: resolveLabel(folder) }));
}

function templateIdFromUi(kind: WorkbookTemplateKind): WorkbookTemplateId | null {
  if (kind === 'import') return null;
  if (kind === 'project') return 'project-plan';
  if (kind === 'budget' || kind === 'pivot' || kind === 'blank' || kind === 'template') return kind;
  return null;
}

function downloadXlsx(buffer: ArrayBuffer, fileName: string): void {
  const href = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const anchor = document.createElement('a');
  anchor.href = href;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(href);
}

function isAbortError(cause: unknown): boolean {
  return Boolean(cause && typeof cause === 'object' && 'name' in cause && cause.name === 'AbortError');
}

export function WorkbookHubContainer({ onOpenWorkbook }: WorkbookHubContainerProps) {
  const { catalog } = useApplicationServices();
  const auth = useAuthSession();
  const authSnapshot = useAuthSnapshot();
  const [activeSection, setActiveSection] = useState<WorkbookHubSection>('start');
  const [activeTab, setActiveTab] = useState<WorkbookCategoryTab>('recent');
  const [entries, setEntries] = useState<readonly WorkbookCatalogEntry[]>([]);
  const [spaces, setSpaces] = useState<readonly WorkspaceSpace[]>([]);
  const [folders, setFolders] = useState<readonly WorkspaceFolder[]>([]);
  const [spaceMembers, setSpaceMembers] = useState<readonly SpaceMember[]>([]);
  const [preferences, setPreferences] = useState<UserPreferences>();
  const [selectedSpaceId, setSelectedSpaceId] = useState('');
  const [newFolderName, setNewFolderName] = useState('');
  const [newSpaceName, setNewSpaceName] = useState('');
  const [memberSubject, setMemberSubject] = useState('');
  const [memberRole, setMemberRole] = useState<'editor' | 'commenter' | 'viewer'>('viewer');
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string>();
  const [activeDialog, setActiveDialog] = useState<ActiveDialog>(() => {
    if (typeof window === 'undefined') return null;
    return new URLSearchParams(window.location.search).get('dialog') === 'import' ? 'import' : null;
  });
  const [pendingTemplate, setPendingTemplate] = useState<WorkbookTemplateId>('blank');
  const [targetId, setTargetId] = useState<string>();
  const [moveLocationId, setMoveLocationId] = useState('local');
  const loadGeneration = useRef(0);
  const loadAbortController = useRef<AbortController | null>(null);

  const load = useCallback(async () => {
    loadAbortController.current?.abort();
    const controller = new AbortController();
    loadAbortController.current = controller;
    const generation = ++loadGeneration.current;
    const requestOptions = { signal: controller.signal };
    setLoading(true);
    setError(undefined);
    try {
      const [listed, remoteSpaces, remotePreferences] = await Promise.all([
        catalog.list({ view: activeSection === 'trash' ? 'trash' : 'all' }, requestOptions),
        authSnapshot.phase === 'authenticated' ? catalog.listSpaces(requestOptions) : Promise.resolve([]),
        authSnapshot.phase === 'authenticated' ? catalog.getUserPreferences(requestOptions) : Promise.resolve(undefined),
      ]);
      const remoteFolders = remoteSpaces.length > 0
        ? (await Promise.all(remoteSpaces.map((space) => catalog.listFolders(space.spaceId, requestOptions)))).flat()
        : [];
      if (generation !== loadGeneration.current || controller.signal.aborted) return;
      setEntries(listed);
      setSpaces(remoteSpaces);
      setFolders(remoteFolders);
      setPreferences(remotePreferences);
      setSelectedSpaceId((current) => current || remoteSpaces[0]?.spaceId || '');
    } catch (cause) {
      if (generation !== loadGeneration.current || controller.signal.aborted || isAbortError(cause)) return;
      setError(cause instanceof Error ? cause.message : '无法加载工作簿目录');
    } finally {
      if (generation === loadGeneration.current) setLoading(false);
    }
  }, [activeSection, authSnapshot.phase, catalog]);

  useEffect(() => { void load(); }, [load]);

  useEffect(() => () => {
    loadAbortController.current?.abort();
    loadGeneration.current += 1;
  }, []);

  useEffect(() => {
    if (authSnapshot.phase !== 'authenticated' || !selectedSpaceId) {
      setSpaceMembers([]);
      return;
    }
    void catalog.listSpaceMembers(selectedSpaceId)
      .then(setSpaceMembers)
      .catch((cause) => {
        setSpaceMembers([]);
        setError(cause instanceof Error ? cause.message : '无法加载空间成员');
      });
  }, [authSnapshot.phase, catalog, selectedSpaceId]);

  const folderLocationOptions = useMemo(() => folderLocations(folders, spaces), [folders, spaces]);
  const locationOptions = useMemo<readonly LocationOption[]>(() => [
    { id: 'local', label: '此浏览器 · 本地文件' },
    ...spaces.map((space) => ({ id: `space:${space.spaceId}`, label: `${space.kind === 'team' ? '团队空间' : '我的云端'} · ${space.name}` })),
    ...folderLocationOptions.map(({ folder, label }) => ({ id: `folder:${folder.spaceId}:${folder.folderId}`, label })),
  ], [folderLocationOptions, spaces]);
  const defaultLocationId = useMemo(() => {
    if (authSnapshot.phase !== 'authenticated') return 'local';
    if (preferences?.defaultFolderId) {
      const folder = folders.find((candidate) => candidate.folderId === preferences.defaultFolderId);
      if (folder) return `folder:${folder.spaceId}:${folder.folderId}`;
    }
    if (preferences?.defaultSpaceId && spaces.some((space) => space.spaceId === preferences.defaultSpaceId)) {
      return `space:${preferences.defaultSpaceId}`;
    }
    return spaces[0] ? `space:${spaces[0].spaceId}` : 'local';
  }, [authSnapshot.phase, folders, preferences?.defaultFolderId, preferences?.defaultSpaceId, spaces]);
  const items = useMemo(() => entries.map(itemFromEntry), [entries]);
  const target = useMemo(() => targetId ? items.find((item) => item.unitId === targetId) : undefined, [items, targetId]);

  const execute = useCallback(async (operation: () => Promise<void>) => {
    setSubmitting(true);
    setError(undefined);
    try {
      await operation();
      setActiveDialog(null);
      setTargetId(undefined);
      await load();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : '工作簿操作失败');
    } finally {
      setSubmitting(false);
    }
  }, [load]);

  const executeSettings = useCallback(async (operation: () => Promise<void>) => {
    setSubmitting(true);
    setError(undefined);
    try {
      await operation();
      await load();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : '设置操作失败');
    } finally {
      setSubmitting(false);
    }
  }, [load]);

  const requireCloudSignIn = useCallback(async (): Promise<boolean> => {
    if (authSnapshot.phase === 'authenticated') return true;
    if (authSnapshot.phase === 'unconfigured') {
      setError('云端工作簿需要配置 VITE_OIDC_ISSUER 与 VITE_OIDC_CLIENT_ID。');
      return false;
    }
    await auth.signIn('/workbooks');
    return false;
  }, [auth, authSnapshot.phase]);

  const requestTemplate = useCallback((kind: WorkbookTemplateKind) => {
    const templateId = templateIdFromUi(kind);
    if (!templateId) {
      setActiveDialog('import');
      return;
    }
    setPendingTemplate(templateId);
    setActiveDialog('create');
  }, []);

  const createWorkbook = useCallback((value: { name: string; locationId: string }) => {
    void execute(async () => {
      const targetLocation = destinationFromLocation(value.locationId);
      if (targetLocation.destination === 'remote' && !await requireCloudSignIn()) return;
      const unitId = createWorkbookUnitId();
      const snapshot = createTemplateSnapshot(pendingTemplate, unitId, value.name);
      const entry = await catalog.create({
        snapshot,
        destination: targetLocation.destination,
        metadata: { spaceId: targetLocation.spaceId, folderId: targetLocation.folderId },
        source: 'native',
      });
      onOpenWorkbook(entry.unitId);
    });
  }, [catalog, execute, onOpenWorkbook, pendingTemplate, requireCloudSignIn]);

  const importWorkbook = useCallback((value: { file: File; locationId: string }) => {
    void execute(async () => {
      const targetLocation = destinationFromLocation(value.locationId);
      if (targetLocation.destination === 'remote' && !await requireCloudSignIn()) return;
      const result = await catalog.importXlsx({
        fileName: value.file.name,
        buffer: await value.file.arrayBuffer(),
        destination: targetLocation.destination,
        spaceId: targetLocation.spaceId,
        folderId: targetLocation.folderId,
        options: { compatibilityTarget: preferences?.importCompatibility ?? 'B' },
      });
      onOpenWorkbook(result.entry.unitId);
    });
  }, [catalog, execute, onOpenWorkbook, preferences?.importCompatibility, requireCloudSignIn]);

  const openWorkbook = useCallback((unitId: string) => {
    void execute(async () => {
      await catalog.open(unitId);
      onOpenWorkbook(unitId);
    });
  }, [catalog, execute, onOpenWorkbook]);

  const exportWorkbook = useCallback((unitId: string) => {
    void execute(async () => {
      const exported = await catalog.exportXlsx(unitId);
      downloadXlsx(exported.buffer, exported.fileName);
    });
  }, [catalog, execute]);

  const syncWorkbook = useCallback((unitId: string) => {
    void execute(async () => {
      if (!await requireCloudSignIn()) return;
      await catalog.syncToServer(unitId);
    });
  }, [catalog, execute, requireCloudSignIn]);

  const navigateSection = useCallback((section: WorkbookHubSection) => {
    if (section === 'new') {
      requestTemplate('blank');
      return;
    }
    if (section === 'open' || section === 'import') {
      setActiveDialog('import');
      return;
    }
    if (section === 'info') {
      setActiveDialog('help');
      return;
    }
    if (section === 'options') {
      setActiveDialog('options');
      return;
    }
    if (section === 'save' || section === 'export' || section === 'close') return;
    setActiveSection(section);
    if (section === 'recent') setActiveTab('recent');
    if (section === 'shared') setActiveTab('shared');
  }, [requestTemplate]);

  return (
    <>
      <WorkbookHubPage
        activeSection={activeSection}
        activeTab={activeTab}
        error={error}
        hasActiveWorkbook={false}
        items={items}
        loading={loading}
        onCopyWorkbook={(unitId) => void execute(async () => { await catalog.copy(unitId); })}
        onCreateTemplate={requestTemplate}
        onExportWorkbook={exportWorkbook}
        onFavoriteWorkbook={(unitId, favorite) => void execute(async () => { await catalog.setFavorite(unitId, favorite); })}
        onImportWorkbook={() => setActiveDialog('import')}
        onMoveWorkbook={(unitId) => { setTargetId(unitId); setMoveLocationId(defaultLocationId); setActiveDialog('move'); }}
        onNavigate={navigateSection}
        onOpenInNewWindow={(unitId) => { window.open(`/workbooks/${encodeURIComponent(unitId)}`, '_blank', 'noopener,noreferrer'); }}
        onOpenWorkbook={openWorkbook}
        onPurgeWorkbook={(unitId) => { setTargetId(unitId); setActiveDialog('purge'); }}
        onRenameWorkbook={(unitId) => { setTargetId(unitId); setActiveDialog('rename'); }}
        onRestoreWorkbook={(unitId) => void execute(async () => { await catalog.restore(unitId); })}
        onRetry={() => void load()}
        onSelectTab={setActiveTab}
        onShowHelp={() => setActiveDialog('help')}
        onShowSettings={() => setActiveDialog('options')}
        onShareWorkbook={(unitId) => { setTargetId(unitId); setActiveDialog('share'); }}
        onSyncWorkbook={syncWorkbook}
        onTrashWorkbook={(unitId) => { setTargetId(unitId); setActiveDialog('trash'); }}
        userName={authSnapshot.displayName ?? undefined}
      />

      <CreateWorkbookDialog
        defaultLocationId={defaultLocationId}
        defaultName={pendingTemplate === 'blank' ? '未命名工作簿' : undefined}
        locationOptions={locationOptions}
        onClose={() => setActiveDialog(null)}
        onSubmit={createWorkbook}
        open={activeDialog === 'create'}
        submitting={submitting}
      />
      <ImportWorkbookDialog
        defaultLocationId={defaultLocationId}
        locationOptions={locationOptions}
        onClose={() => setActiveDialog(null)}
        onSubmit={importWorkbook}
        open={activeDialog === 'import'}
        submitting={submitting}
      />
      <RenameWorkbookDialog
        currentName={target?.name ?? ''}
        onClose={() => setActiveDialog(null)}
        onSubmit={(name) => target && void execute(async () => { await catalog.rename(target.unitId, name); })}
        open={activeDialog === 'rename'}
        submitting={submitting}
      />
      <DeleteWorkbookDialog
        onClose={() => setActiveDialog(null)}
        onConfirm={() => target && void execute(async () => {
          if (activeDialog === 'purge') await catalog.purge(target.unitId);
          else await catalog.moveToTrash(target.unitId);
        })}
        open={activeDialog === 'trash' || activeDialog === 'purge'}
        permanent={activeDialog === 'purge'}
        submitting={submitting}
        workbookName={target?.name ?? ''}
      />
      <ShareWorkbookDialog
        onClose={() => setActiveDialog(null)}
        onSubmit={(value) => target && void execute(async () => { await catalog.grantAccess(target.unitId, value.subject, value.role); })}
        open={activeDialog === 'share'}
        submitting={submitting}
        workbookName={target?.name ?? ''}
      />
      <Dialog closeLabel="关闭移动工作簿" onClose={() => setActiveDialog(null)} open={activeDialog === 'move'} title="移动工作簿">
        <Stack gap="md">
          <Text size="sm">选择目标空间。跨空间移动只允许所有者，服务端会执行最终授权校验。</Text>
          <LocationPicker aria-label="目标空间" id="move-workbook-location" onChange={(event) => setMoveLocationId(event.target.value)} options={locationOptions} value={moveLocationId} />
          <Button disabled={!target} loading={submitting} onClick={() => {
            if (!target) return;
            void execute(async () => {
              const destination = destinationFromLocation(moveLocationId);
              if (destination.destination === 'remote' && !await requireCloudSignIn()) return;
              await catalog.move(target.unitId, { spaceId: destination.spaceId, folderId: destination.folderId ?? null });
            });
          }} size="sm" variant="brand">移动</Button>
        </Stack>
      </Dialog>
      <Dialog closeLabel="关闭帮助" onClose={() => setActiveDialog(null)} open={activeDialog === 'help'} title="工作簿存储说明">
        <Stack gap="sm">
          <Text size="sm">云端工作簿由服务器保存，本地文件和离线待同步变更保留在此浏览器的 IndexedDB 中。</Text>
          <Text size="sm">导入 Excel 会创建新工作簿；导出会基于最新快照和原始 XLSX 包生成副本。</Text>
          {authSnapshot.phase !== 'authenticated' ? <Button onClick={() => void auth.signIn('/workbooks')} size="sm" variant="brand">登录以使用云端文件</Button> : null}
        </Stack>
      </Dialog>
      <Dialog closeLabel="关闭选项" onClose={() => setActiveDialog(null)} open={activeDialog === 'options'} title="选项">
        <Stack gap="lg">
          <Stack gap="xs">
            <Text size="sm" weight="semibold">默认位置与云端会话</Text>
            <Text size="sm">当前默认新建位置：{defaultLocationId === 'local' ? '本地文件' : '云端空间'}。自动保存、自动同步、离线缓存和导入兼容级别保存到当前用户的全局偏好。</Text>
            <Select aria-label="XLSX 导入兼容级别" disabled={!preferences || authSnapshot.phase !== 'authenticated'} value={preferences?.importCompatibility ?? 'B'} options={[{ value: 'A', label: '严格：无法安全保留则拒绝' }, { value: 'B', label: '平衡：编辑已支持部分并保留原包' }, { value: 'C', label: '尽可能转换：允许明确近似' }]} onChange={(event) => void executeSettings(async () => { setPreferences(await catalog.putUserPreferences({ importCompatibility: event.target.value as 'A' | 'B' | 'C' })); })} />
            <CheckToggle checked={preferences?.autoSave ?? true} disabled={!preferences || authSnapshot.phase !== 'authenticated'} label="自动保存" onChange={(event) => void executeSettings(async () => { setPreferences(await catalog.putUserPreferences({ autoSave: event.target.checked })); })} />
            <CheckToggle checked={preferences?.autoSync ?? true} disabled={!preferences || authSnapshot.phase !== 'authenticated'} label="自动同步" onChange={(event) => void executeSettings(async () => { setPreferences(await catalog.putUserPreferences({ autoSync: event.target.checked })); })} />
            <CheckToggle checked={preferences?.offlineCache ?? true} disabled={!preferences || authSnapshot.phase !== 'authenticated'} label="离线缓存" onChange={(event) => void executeSettings(async () => { setPreferences(await catalog.putUserPreferences({ offlineCache: event.target.checked })); })} />
            <Button disabled={authSnapshot.phase !== 'authenticated'} onClick={() => void auth.signOut()} size="sm" variant="outline">退出云端会话</Button>
          </Stack>
          <Stack gap="xs" className="border-t border-slate-100 pt-4">
            <Text size="sm" weight="semibold">团队空间</Text>
            <TextInput aria-label="团队空间名称" onChange={(event) => setNewSpaceName(event.target.value)} placeholder="例如：项目团队" value={newSpaceName} />
            <Button disabled={!newSpaceName.trim() || authSnapshot.phase !== 'authenticated'} loading={submitting} onClick={() => void executeSettings(async () => {
              const created = await catalog.createSpace({ kind: 'team', name: newSpaceName.trim() });
              setNewSpaceName('');
              setSelectedSpaceId(created.spaceId);
            })} size="sm" variant="brand">创建团队空间</Button>
          </Stack>
          <Stack gap="xs" className="border-t border-slate-100 pt-4">
            <Text size="sm" weight="semibold">文件夹与成员</Text>
            <LocationPicker aria-label="管理空间" onChange={(event) => setSelectedSpaceId(event.target.value)} options={spaces.map((space) => ({ id: space.spaceId, label: `${space.kind === 'team' ? '团队' : '个人'} · ${space.name}` }))} value={selectedSpaceId} />
            <TextInput aria-label="新文件夹名称" onChange={(event) => setNewFolderName(event.target.value)} placeholder="新文件夹名称" value={newFolderName} />
            <Button disabled={!selectedSpaceId || !newFolderName.trim() || authSnapshot.phase !== 'authenticated'} loading={submitting} onClick={() => void executeSettings(async () => {
              await catalog.createFolder(selectedSpaceId, { name: newFolderName.trim() });
              setNewFolderName('');
            })} size="sm" variant="secondary">创建文件夹</Button>
            <TextInput aria-label="空间成员" onChange={(event) => setMemberSubject(event.target.value)} placeholder="成员账号或邮箱" value={memberSubject} />
            <Select aria-label="空间成员角色" onChange={(event) => setMemberRole(event.target.value as 'editor' | 'commenter' | 'viewer')} options={[{ value: 'editor', label: '可编辑' }, { value: 'commenter', label: '可评论' }, { value: 'viewer', label: '只读' }]} value={memberRole} />
            <Button disabled={!selectedSpaceId || !memberSubject.trim() || authSnapshot.phase !== 'authenticated'} loading={submitting} onClick={() => void executeSettings(async () => {
              await catalog.putSpaceMember(selectedSpaceId, memberSubject.trim(), memberRole);
              setMemberSubject('');
              setSpaceMembers(await catalog.listSpaceMembers(selectedSpaceId));
            })} size="sm" variant="secondary">添加或更新成员</Button>
            {spaceMembers.map((member) => <Text key={`${member.spaceId}:${member.subject}`} size="xs" tone="muted">{member.subject} · {member.role}</Text>)}
          </Stack>
        </Stack>
      </Dialog>
    </>
  );
}
