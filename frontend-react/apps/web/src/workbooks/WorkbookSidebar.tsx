import { Button, Divider, Icon, Stack, Text } from '@react-sheets/ui-system';
import type { IconName } from '@react-sheets/ui-system';
import type { WorkbookHubSection } from './types';

export interface WorkbookSidebarProps {
  activeSection: WorkbookHubSection;
  onNavigate: (section: WorkbookHubSection) => void;
  hasActiveWorkbook?: boolean;
}

interface SidebarItem {
  id: WorkbookHubSection;
  label: string;
  icon: IconName;
}

const primaryItems: readonly SidebarItem[] = [
  { id: 'start', label: '开始', icon: 'home' },
  { id: 'new', label: '新建', icon: 'file-plus' },
  { id: 'open', label: '打开', icon: 'folder-open' },
  { id: 'recent', label: '最近', icon: 'clock' },
  { id: 'shared', label: '共享', icon: 'share' },
];

const fileItems: readonly SidebarItem[] = [
  { id: 'info', label: '信息', icon: 'info' },
  { id: 'save', label: '保存', icon: 'save' },
  { id: 'import', label: '导入', icon: 'upload' },
  { id: 'export', label: '导出', icon: 'download' },
];

const utilityItems: readonly SidebarItem[] = [
  { id: 'trash', label: '回收站', icon: 'trash' },
  { id: 'close', label: '关闭', icon: 'x' },
];

function SidebarItemView({ item, active, disabled = false, onSelect }: { item: SidebarItem; active: boolean; disabled?: boolean; onSelect: () => void }) {
  return (
    <Button
      aria-current={active ? 'page' : undefined}
      aria-disabled={disabled || undefined}
      aria-label={item.label}
      className={active
        ? 'relative h-12 w-full justify-start rounded-lg bg-brand-soft px-5 text-[15px] font-semibold text-brand-dark shadow-none before:absolute before:inset-y-0 before:left-0 before:w-1 before:rounded-r-full before:bg-brand'
        : 'h-12 w-full justify-start rounded-lg px-5 text-[15px] font-normal text-slate-700 shadow-none hover:bg-slate-100 hover:text-brand-dark'}
      icon={item.icon}
      onClick={onSelect}
      disabled={disabled}
      size="md"
      variant="ghost"
    >
      {item.label}
    </Button>
  );
}

function SidebarGroup({ items, activeSection, hasActiveWorkbook, onNavigate }: { items: readonly SidebarItem[]; activeSection: WorkbookHubSection; hasActiveWorkbook: boolean; onNavigate: (section: WorkbookHubSection) => void }) {
  return (
    <Stack gap="none" className="w-full">
      {items.map((item) => {
        const needsActiveWorkbook = item.id === 'info' || item.id === 'save' || item.id === 'export' || item.id === 'close';
        return <SidebarItemView key={item.id} active={activeSection === item.id} disabled={needsActiveWorkbook && !hasActiveWorkbook} item={item} onSelect={() => onNavigate(item.id)} />;
      })}
    </Stack>
  );
}

export function WorkbookSidebar({ activeSection, hasActiveWorkbook = false, onNavigate }: WorkbookSidebarProps) {
  return (
    <Stack as="nav" aria-label="工作簿导航" gap="none" className="w-[168px] shrink-0 border-r border-slate-200 bg-chrome px-3 py-5">
      <SidebarGroup activeSection={activeSection} hasActiveWorkbook={hasActiveWorkbook} items={primaryItems} onNavigate={onNavigate} />
      <Divider className="my-5 bg-slate-200" />
      <SidebarGroup activeSection={activeSection} hasActiveWorkbook={hasActiveWorkbook} items={fileItems} onNavigate={onNavigate} />
      <Divider className="my-5 bg-slate-200" />
      <SidebarGroup activeSection={activeSection} hasActiveWorkbook={hasActiveWorkbook} items={utilityItems} onNavigate={onNavigate} />
      <Stack gap="none" className="mt-auto">
        <Divider className="mb-5 bg-slate-200" />
        <SidebarItemView active={activeSection === 'options'} item={{ id: 'options', label: '选项', icon: 'settings' }} onSelect={() => onNavigate('options')} />
      </Stack>
      <Text className="sr-only">工作簿导航</Text>
    </Stack>
  );
}
