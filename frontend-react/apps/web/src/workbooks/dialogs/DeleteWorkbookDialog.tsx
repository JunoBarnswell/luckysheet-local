import { Button, Dialog, Stack, Text } from '@react-sheets/ui-system';

export interface DeleteWorkbookDialogProps {
  open: boolean;
  workbookName: string;
  permanent?: boolean;
  onClose: () => void;
  onConfirm: () => void;
  submitting?: boolean;
}

export function DeleteWorkbookDialog({ open, workbookName, permanent = false, onClose, onConfirm, submitting = false }: DeleteWorkbookDialogProps) {
  return (
    <Dialog
      closeLabel="关闭删除确认"
      description={permanent ? '此操作不可撤销' : '工作簿会先移动到回收站'}
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button loading={submitting} onClick={onConfirm} size="sm" variant="danger">{permanent ? '永久删除' : '移到回收站'}</Button></>}
      maxWidth="sm"
      onClose={onClose}
      open={open}
      title={permanent ? '永久删除工作簿？' : '移到回收站？'}
      testId="delete-workbook-dialog"
    >
      <Stack gap="sm">
        <Text className="text-[14px] text-slate-700">确定要{permanent ? '永久删除' : '将'} <Text as="strong" weight="semibold">{workbookName}</Text>{permanent ? '吗？' : '移到回收站吗？'}</Text>
        <Text size="xs" tone="muted">{permanent ? '工作簿的快照、历史记录、共享设置和本地镜像都会被清理。' : 'Owner 可以在回收站中恢复工作簿。'}</Text>
      </Stack>
    </Dialog>
  );
}
