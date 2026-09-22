import { Button, Dialog, FileIcon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem } from '../types';

export interface WorkbookInfoDialogProps {
  open: boolean;
  item?: WorkbookCatalogItem;
  onClose: () => void;
  onOpen?: (unitId: string) => void;
  onExport?: (unitId: string) => void;
}

function format(value: string | number | undefined) {
  if (value === undefined) return '—';
  if (typeof value === 'number') return value.toLocaleString('zh-CN');
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString('zh-CN');
}

export function WorkbookInfoDialog({ open, item, onClose, onOpen, onExport }: WorkbookInfoDialogProps) {
  return (
    <Dialog closeLabel="关闭工作簿信息" footer={<Button onClick={onClose} size="sm" variant="brand">完成</Button>} maxWidth="md" onClose={onClose} open={open} title="工作簿信息" testId="workbook-info-dialog">
      {item ? <Stack gap="md">
        <Inline gap="sm"><FileIcon size="md" /><Stack gap="none"><Text size="sm" weight="semibold">{item.name}</Text><Text size="xs" tone="muted">{item.unitId}</Text></Stack></Inline>
        <Stack gap="sm" className="rounded-lg border border-slate-100 bg-slate-50 p-4">
          <Inline gap="md" className="justify-between"><Text size="xs" tone="muted">位置</Text><Text size="sm">{item.locationLabel}</Text></Inline>
          <Inline gap="md" className="justify-between"><Text size="xs" tone="muted">最近修改</Text><Text size="sm">{format(item.updatedAt)}</Text></Inline>
          <Inline gap="md" className="justify-between"><Text size="xs" tone="muted">权限</Text><Text size="sm">{item.role}</Text></Inline>
          <Inline gap="md" className="justify-between"><Text size="xs" tone="muted">本地 / 服务端版本</Text><Text size="sm">{format(item.localRevision)} / {format(item.serverRevision)}</Text></Inline>
          <Inline gap="md" className="justify-between"><Text size="xs" tone="muted">待同步操作</Text><Text size="sm">{format(item.pendingOperationCount ?? 0)}</Text></Inline>
        </Stack>
        <Inline gap="sm">{onOpen ? <Button icon="folder-open" onClick={() => onOpen(item.unitId)} size="sm" variant="brand">打开工作簿</Button> : null}{onExport ? <Button icon="download" onClick={() => onExport(item.unitId)} size="sm" variant="outline">导出副本</Button> : null}</Inline>
      </Stack> : <Text size="sm" tone="muted">请先选择一个工作簿。</Text>}
    </Dialog>
  );
}
