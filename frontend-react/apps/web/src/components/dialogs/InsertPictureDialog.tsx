import { Dialog, FileButton, Stack, Text } from '@react-sheets/ui-system';

export interface InsertPictureDialogProps {
  open: boolean;
  onClose: () => void;
  onInsert: (file: File, placement: 'cell' | 'floating') => Promise<void>;
}

export function InsertPictureDialog({ open, onClose, onInsert }: InsertPictureDialogProps) {
  return (
    <Dialog open={open} title="插入图片" maxWidth="sm" onClose={onClose}>
      <Stack gap="md">
        <Text size="sm" tone="muted">选择图片后将作为可移动、可缩放的工作表对象插入。</Text>
        <FileButton accept="image/*" icon="picture" size="sm" variant="primary" onFile={(file) => { void onInsert(file, 'cell').then(onClose); }}>嵌入当前单元格</FileButton>
        <FileButton accept="image/*" icon="picture" size="sm" variant="secondary" onFile={(file) => { void onInsert(file, 'floating').then(onClose); }}>浮于单元格上方</FileButton>
      </Stack>
    </Dialog>
  );
}
