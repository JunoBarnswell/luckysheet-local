import { Box, Button, Icon, Inline } from '@react-sheets/ui-system';

export interface WorkbookActionBarProps {
  selectedCount: number;
  onCreate: () => void;
  onImport: () => void;
  onExportSelected: () => void;
  onSyncSelected: () => void;
  onMoveSelected?: () => void;
  canExport?: boolean;
  canSync?: boolean;
  canMove?: boolean;
}

export function WorkbookActionBar({ selectedCount, onCreate, onImport, onExportSelected, onSyncSelected, onMoveSelected, canExport = true, canSync = true, canMove = true }: WorkbookActionBarProps) {
  const hasSelection = selectedCount > 0;
  return (
    <Inline gap="sm" className="min-h-10 flex-wrap justify-between">
      <Inline gap="sm">
        <Button icon="plus" onClick={onCreate} size="sm" variant="brand">新建工作簿</Button>
        <Button icon="upload" onClick={onImport} size="sm" variant="secondary">上传导入</Button>
        <Button disabled={!hasSelection || !canExport} icon="download" onClick={onExportSelected} size="sm" variant="secondary">导出当前工作簿</Button>
        <Button disabled={!hasSelection || !canSync} icon="cloud-check" onClick={onSyncSelected} size="sm" variant="secondary">同步到服务端</Button>
        {onMoveSelected ? <Button disabled={!hasSelection || !canMove} icon="folder" onClick={onMoveSelected} size="sm" variant="secondary">移动</Button> : null}
      </Inline>
      {hasSelection ? <Box className="rounded-full bg-brand-soft px-3 py-1 text-[11px] font-medium text-brand-dark">已选择 {selectedCount} 项</Box> : null}
    </Inline>
  );
}
