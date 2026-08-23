import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { Box, Button, Dialog, Heading, Inline, Panel, PanelBody, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import type { WorkbookCategoryTab, WorkbookHubPageProps, WorkbookHubSection } from './types';
import { CreateTemplateGrid } from './CreateTemplateGrid';
import { StorageInfoBanner } from './StorageInfoBanner';
import { WorkbookActionBar } from './WorkbookActionBar';
import { WorkbookCategoryTabs } from './WorkbookCategoryTabs';
import { WorkbookGrid } from './WorkbookGrid';
import { WorkbookFilterDialog, type WorkbookFilterValues } from './WorkbookFilterDialog';
import { WorkbookSearch } from './WorkbookSearch';
import { WorkbookSidebar } from './WorkbookSidebar';
import { WorkbookTable } from './WorkbookTable';
import { WorkbookTopBar } from './WorkbookTopBar';

const infoSections = new Set<WorkbookHubSection>(['info', 'save', 'import', 'export', 'close', 'options']);

function categoryItems(section: WorkbookHubSection, tab: WorkbookCategoryTab, items: WorkbookHubPageProps['items']) {
  return items.filter((item) => {
    if (section === 'trash') return item.lifecycle === 'trashed';
    if (item.lifecycle === 'trashed') return false;
    if (section === 'shared' || tab === 'shared') return item.role !== 'owner';
    if (tab === 'local') return item.storageLocation === 'local';
    if (tab === 'cloud') return item.storageLocation !== 'local';
    return true;
  });
}

function SectionTitle({ section }: { section: WorkbookHubSection }) {
  if (section === 'trash') return <Heading as="h1" size="xl" className="text-[28px] tracking-[-0.04em]">回收站</Heading>;
  if (section === 'new') return <Heading as="h1" size="xl" className="text-[28px] tracking-[-0.04em]">新建工作簿</Heading>;
  if (section === 'open') return <Heading as="h1" size="xl" className="text-[28px] tracking-[-0.04em]">打开工作簿</Heading>;
  if (section === 'shared') return <Heading as="h1" size="xl" className="text-[28px] tracking-[-0.04em]">与我共享</Heading>;
  return <Heading as="h1" size="xl" className="text-[28px] tracking-[-0.04em]">早上好</Heading>;
}

function SectionFallback({ section, hasSelection, onBack, onImport, onExport }: { section: WorkbookHubSection; hasSelection: boolean; onBack: () => void; onImport: () => void; onExport: () => void }) {
  const labels: Record<string, string> = {
    info: '信息',
    save: '保存',
    import: '导入',
    export: '导出',
    close: '关闭',
    options: '选项',
  };
  if (section === 'import') {
    return <Panel tone="subtle"><PanelBody><Stack gap="sm"><PanelTitle as="h2" size="sm">导入 Excel 文件</PanelTitle><Text size="sm" tone="muted">导入会创建新的工作簿，不会覆盖当前文件。</Text><Button icon="upload" onClick={onImport} size="sm" variant="brand">选择 Excel 文件</Button></Stack></PanelBody></Panel>;
  }
  if (section === 'export') {
    return hasSelection
      ? <Panel tone="subtle"><PanelBody><Stack gap="sm"><PanelTitle as="h2" size="sm">导出工作簿</PanelTitle><Text size="sm" tone="muted">已选择文件，可以导出为 Excel 副本。</Text><Button icon="download" onClick={onExport} size="sm" variant="brand">导出已选文件</Button></Stack></PanelBody></Panel>
      : <StatePanel actionLabel="返回文件列表" description="先在文件列表中选择要导出的工作簿。" kind="empty" onAction={onBack} title="尚未选择工作簿" />;
  }
  if (section === 'options') {
    return <Panel tone="subtle"><PanelBody><Stack gap="sm"><PanelTitle as="h2" size="sm">文件中心选项</PanelTitle><Text size="sm" tone="muted">默认新建位置、自动同步、离线缓存和导入兼容级别由宿主容器提供并持久化。</Text><Button onClick={onBack} size="sm" variant="outline">返回文件列表</Button></Stack></PanelBody></Panel>;
  }
  return <StatePanel actionLabel="返回工作簿中心" description={`${labels[section] ?? section}需要从活动工作簿上下文打开。`} kind="empty" onAction={onBack} title={`请选择一个工作簿后使用${labels[section] ?? section}`} />;
}

export function WorkbookHubPage({
  items,
  activeSection = 'start',
  activeTab: controlledTab,
  loading = false,
  error,
  userName,
  onRetry,
  onSelectTab,
  onNavigate,
  onCreateTemplate,
  onOpenWorkbook,
  onOpenInNewWindow,
  onImportWorkbook,
  onExportWorkbook,
  onSyncWorkbook,
  onRenameWorkbook,
  onCopyWorkbook,
  onMoveWorkbook,
  onTrashWorkbook,
  onRestoreWorkbook,
  onPurgeWorkbook,
  onFavoriteWorkbook,
  onShareWorkbook,
  onShowHelp,
  onShowSettings,
  sectionContent,
  hasActiveWorkbook = false,
}: WorkbookHubPageProps) {
  const [tab, setTab] = useState<WorkbookCategoryTab>(controlledTab ?? 'recent');
  const [query, setQuery] = useState('');
  const [selectedKeys, setSelectedKeys] = useState<readonly string[]>([]);
  const [viewMode, setViewMode] = useState<'list' | 'grid'>('list');
  const [filterOpen, setFilterOpen] = useState(false);
  const [filters, setFilters] = useState<WorkbookFilterValues>({ favoritesOnly: false, localOnly: false, sharedOnly: false, needsSync: false });
  const [draftFilters, setDraftFilters] = useState<WorkbookFilterValues>(filters);
  const [mobileNavigationOpen, setMobileNavigationOpen] = useState(false);

  useEffect(() => { if (controlledTab) setTab(controlledTab); }, [controlledTab]);

  const visibleItems = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase();
    return categoryItems(activeSection, tab, items).filter((item) => {
      if (filters.favoritesOnly && !item.favorite) return false;
      if (filters.localOnly && item.storageLocation !== 'local') return false;
      if (filters.sharedOnly && item.role === 'owner') return false;
      if (filters.needsSync && !['pending', 'syncing', 'offline', 'conflict', 'error'].includes(item.syncStatus)) return false;
      if (!normalized) return true;
      return [item.name, item.locationLabel, item.ownerName, item.ownerSubject, item.folderPath?.join(' / ')].filter(Boolean).some((value) => value?.toLocaleLowerCase().includes(normalized));
    });
  }, [activeSection, filters, items, query, tab]);

  const selectedItems = useMemo(() => items.filter((item) => selectedKeys.includes(item.unitId)), [items, selectedKeys]);
  const showTemplates = activeSection === 'start' || activeSection === 'new';
  const showCatalog = !infoSections.has(activeSection);
  const handleTabChange = (next: WorkbookCategoryTab) => { setTab(next); onSelectTab?.(next); };
  const openCreate = () => onCreateTemplate('blank');
  const exportSelected = () => { selectedItems.forEach((item) => onExportWorkbook(item.unitId)); };
  const syncSelected = () => { selectedItems.filter((item) => item.storageLocation !== 'remote').forEach((item) => onSyncWorkbook(item.unitId)); };
  const navigate = (section: WorkbookHubSection) => { setMobileNavigationOpen(false); onNavigate(section); };
  const openFilters = () => { setDraftFilters(filters); setFilterOpen(true); };
  const applyFilters = () => { setFilters(draftFilters); setFilterOpen(false); };

  const rowMenuProps = {
    onExport: onExportWorkbook,
    onOpenInNewWindow,
    onSync: onSyncWorkbook,
    onRename: onRenameWorkbook,
    onCopy: onCopyWorkbook,
    onMove: onMoveWorkbook,
    onTrash: onTrashWorkbook,
    onRestore: onRestoreWorkbook,
    onPurge: onPurgeWorkbook,
    onFavorite: onFavoriteWorkbook,
    onShare: onShareWorkbook,
  };

  const emptyState: ReactNode = (
    <StatePanel
      actionLabel={activeSection === 'trash' ? undefined : '创建空白工作簿'}
      description={activeSection === 'trash' ? '移到回收站的工作簿会显示在这里。' : query ? '没有匹配的工作簿，请尝试其他关键词。' : '创建或导入一个工作簿后，它会显示在这里。'}
      kind="empty"
      onAction={query || activeSection === 'trash' ? undefined : openCreate}
      title={query ? '没有找到工作簿' : activeSection === 'trash' ? '回收站为空' : '还没有工作簿'}
    />
  );

  return (
    <Box as="main" className="flex min-h-screen min-w-0 flex-col overflow-hidden bg-white text-ink" data-surface="workbook-hub" data-testid="workbook-hub">
      <WorkbookTopBar onBrandClick={() => navigate('start')} onHelp={onShowHelp} onSettings={onShowSettings} />
      <Box className="flex min-h-0 flex-1">
        <Box className="hidden min-[860px]:flex"><WorkbookSidebar activeSection={activeSection} hasActiveWorkbook={hasActiveWorkbook} onNavigate={navigate} /></Box>
        <Box className="min-w-0 flex-1 overflow-y-auto">
          <Box className="w-full px-6 py-6 min-[1100px]:px-12 min-[1100px]:py-7 min-[1440px]:max-w-none min-[1440px]:pl-[50px] min-[1440px]:pr-[36px] min-[1440px]:pt-[30px]">
            <Inline gap="sm" className="mb-5 min-[860px]:hidden">
              <Button aria-label="打开工作簿导航" icon="menu" iconOnly onClick={() => setMobileNavigationOpen(true)} size="sm" variant="outline" />
              <Text size="sm" weight="semibold">工作簿中心</Text>
            </Inline>
            <Stack gap="lg">
              <SectionTitle section={activeSection} />
              {showTemplates ? <CreateTemplateGrid onMoreTemplates={() => onCreateTemplate('template')} onSelect={onCreateTemplate} /> : null}
              {activeSection === 'start' ? <StorageInfoBanner onLearnMore={() => navigate('info')} /> : null}
              {showCatalog ? (
                <Stack gap="md">
                  <WorkbookActionBar canExport={selectedItems.length > 0} canSync={selectedItems.some((item) => item.storageLocation !== 'remote')} onCreate={openCreate} onExportSelected={exportSelected} onImport={onImportWorkbook} onMoveSelected={() => selectedItems[0] && onMoveWorkbook(selectedItems[0].unitId)} onSyncSelected={syncSelected} selectedCount={selectedKeys.length} />
                  <Inline gap="lg" className="items-end justify-between">
                    <WorkbookCategoryTabs activeTab={activeSection === 'shared' ? 'shared' : tab} onChange={handleTabChange} />
                    <Box className="hidden min-w-0 flex-1 justify-end min-[860px]:flex"><WorkbookSearch onChange={setQuery} onFilter={openFilters} onToggleView={() => setViewMode((current) => current === 'list' ? 'grid' : 'list')} value={query} viewMode={viewMode} /></Box>
                  </Inline>
                  <Box className="min-[860px]:hidden"><WorkbookSearch onChange={setQuery} onFilter={openFilters} onToggleView={() => setViewMode((current) => current === 'list' ? 'grid' : 'list')} value={query} viewMode={viewMode} /></Box>
                  {loading ? <StatePanel kind="loading" description="正在加载工作簿目录。" title="加载文件" /> : null}
                  {error ? <StatePanel actionLabel="重试" description={error} kind="error" onAction={onRetry} title="工作簿目录加载失败" /> : null}
                  {!loading && !error && viewMode === 'list' ? <WorkbookTable empty={emptyState} items={visibleItems} onOpen={onOpenWorkbook} onSelectionChange={setSelectedKeys} selectedKeys={selectedKeys} {...rowMenuProps} /> : null}
                  {!loading && !error && viewMode === 'grid' && visibleItems.length > 0 ? <WorkbookGrid items={visibleItems} onOpen={onOpenWorkbook} onSelectionChange={setSelectedKeys} selectedKeys={selectedKeys} {...rowMenuProps} /> : null}
                  {!loading && !error && viewMode === 'grid' && visibleItems.length === 0 ? emptyState : null}
                  {!loading && !error ? <Inline gap="none" className="justify-center border-t border-slate-100 pt-3"><Text size="xs" tone="subtle">共 {visibleItems.length} 项</Text></Inline> : null}
                </Stack>
              ) : sectionContent?.[activeSection] ?? <SectionFallback hasSelection={selectedItems.length > 0} onBack={() => navigate('start')} onExport={exportSelected} onImport={onImportWorkbook} section={activeSection} />}
            </Stack>
          </Box>
        </Box>
      </Box>
      <Dialog closeLabel="关闭工作簿导航" onClose={() => setMobileNavigationOpen(false)} open={mobileNavigationOpen} title="工作簿中心" bodyClassName="p-0" maxWidth="sm" testId="workbook-mobile-navigation">
        <WorkbookSidebar activeSection={activeSection} hasActiveWorkbook={hasActiveWorkbook} onNavigate={navigate} />
      </Dialog>
      <WorkbookFilterDialog onApply={applyFilters} onChange={setDraftFilters} onClose={() => setFilterOpen(false)} open={filterOpen} value={draftFilters} />
      {userName ? <Text className="sr-only">当前用户：{userName}</Text> : null}
    </Box>
  );
}
