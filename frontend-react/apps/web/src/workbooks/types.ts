import type { ReactNode } from 'react';
import type { IconName, TemplatePreviewKind } from '@react-sheets/ui-system';

export type WorkbookRole = 'owner' | 'editor' | 'commenter' | 'viewer';
export type WorkbookStorageLocation = 'local' | 'remote' | 'mirrored';
export type WorkbookSyncStatus = 'synced' | 'syncing' | 'pending' | 'offline' | 'conflict' | 'error';
export type WorkbookLifecycle = 'active' | 'trashed';
export type WorkbookSourceKind = 'native' | 'document-import';

export interface WorkbookCatalogItem {
  unitId: string;
  name: string;
  updatedAt: string;
  locationLabel: string;
  storageLocation: WorkbookStorageLocation;
  syncStatus: WorkbookSyncStatus;
  lifecycle: WorkbookLifecycle;
  role: WorkbookRole;
  sourceKind: WorkbookSourceKind;
  ownerName?: string;
  ownerSubject?: string;
  folderId?: string;
  folderPath?: readonly string[];
  favorite: boolean;
  revision?: number;
  localRevision?: number;
  serverRevision?: number;
  pendingOperationCount?: number;
  sourceFileName?: string;
  sizeBytes?: number;
}

export type WorkbookHubSection = 'start' | 'new' | 'open' | 'recent' | 'shared' | 'info' | 'save' | 'import' | 'export' | 'trash' | 'close' | 'options';
export type WorkbookCategoryTab = 'recent' | 'cloud' | 'local' | 'shared';
export type WorkbookViewMode = 'list' | 'grid';
export type WorkbookTemplateKind = TemplatePreviewKind;

export interface WorkbookOpenOptions {
  initialCell?: string;
}

export interface WorkbookTemplateDefinition {
  kind: WorkbookTemplateKind;
  title: string;
  description: string;
  icon?: IconName;
}

export const workbookTemplates: readonly WorkbookTemplateDefinition[] = [
  { kind: 'blank', title: '空白工作簿', description: '从空白表格开始' },
  { kind: 'template', title: '从模板创建', description: '浏览更多模板', icon: 'sparkles' },
  { kind: 'import', title: '打开 / 导入原生文档', description: '从原生文档协议开始', icon: 'upload' },
  { kind: 'pivot', title: '数据透视表模板', description: '快速分析业务数据', icon: 'table-pivot' },
  { kind: 'project', title: '项目计划模板', description: '规划项目进度与任务', icon: 'chart' },
  { kind: 'budget', title: '预算模板', description: '管理收支与预算', icon: 'calculator' },
  { kind: 'designer-demo', title: 'Designer Demo', description: 'SpreadJS Designer 视觉验收', icon: 'grid' },
];

export interface WorkbookHubController {
  onNavigate: (section: WorkbookHubSection) => void;
  onCreateTemplate: (kind: WorkbookTemplateKind) => void;
  onOpenWorkbook: (unitId: string, options?: WorkbookOpenOptions) => void;
  onOpenInNewWindow: (unitId: string) => void;
  onImportWorkbook: () => void;
  onExportWorkbook: (unitId: string) => void;
  onSyncWorkbook: (unitId: string) => void;
  onRenameWorkbook: (unitId: string) => void;
  onCopyWorkbook: (unitId: string) => void;
  onMoveWorkbook: (unitId: string) => void;
  onTrashWorkbook: (unitId: string) => void;
  onRestoreWorkbook: (unitId: string) => void;
  onPurgeWorkbook: (unitId: string) => void;
  onFavoriteWorkbook: (unitId: string, favorite: boolean) => void;
  onShareWorkbook: (unitId: string) => void;
  onShowHelp: () => void;
  onShowSettings: () => void;
}

export interface WorkbookHubPageProps extends WorkbookHubController {
  items: readonly WorkbookCatalogItem[];
  activeSection?: WorkbookHubSection;
  activeTab?: WorkbookCategoryTab;
  loading?: boolean;
  error?: string;
  userName?: string;
  onRetry?: () => void;
  onSelectTab?: (tab: WorkbookCategoryTab) => void;
  hasActiveWorkbook?: boolean;
  /** Host-provided panels for contextual File/Backstage sections. */
  sectionContent?: Partial<Record<WorkbookHubSection, ReactNode>>;
}

export interface WorkbookBackstageAction {
  id: string;
  label: string;
  description?: string;
  icon: IconName;
  disabled?: boolean;
  onSelect: () => void;
}

export interface WorkbookBackstageShellProps {
  workbookName: string;
  syncStatus: WorkbookSyncStatus;
  readOnly?: boolean;
  onBack: () => void;
  onHelp: () => void;
  onSettings: () => void;
  actions: readonly WorkbookBackstageAction[];
  activeActionId?: string;
  children?: ReactNode;
}
