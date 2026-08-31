import { useEffect, useState } from 'react';
import { Dialog, FileButton, Stack, Text } from '@react-sheets/ui-system';

export interface InsertPictureDialogProps {
  open: boolean;
  onClose: () => void;
  onInsert: (file: File, placement: 'cell' | 'floating') => Promise<void>;
}

export function InsertPictureDialog({ open, onClose, onInsert }: InsertPictureDialogProps) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => { if (!open) { setBusy(false); setError(null); } }, [open]);
  const insert = (file: File, placement: 'cell' | 'floating'): void => {
    setBusy(true);
    setError(null);
    void onInsert(file, placement)
      .then(onClose)
      .catch((cause) => setError(cause instanceof Error ? cause.message : 'IMAGE_INSERT_FAILED: image insertion failed'))
      .finally(() => setBusy(false));
  };
  return (
    <Dialog open={open} title="插入图片" maxWidth="sm" onClose={onClose}>
      <Stack gap="md">
        <Text size="sm" tone="muted">选择图片后将作为可移动、可缩放的工作表对象插入。</Text>
        {error ? <Text size="xs" tone="danger">{error}</Text> : null}
        <FileButton accept="image/*" disabled={busy} loading={busy} icon="picture" size="sm" variant="primary" onFile={(file) => insert(file, 'cell')}>嵌入当前单元格</FileButton>
        <FileButton accept="image/*" disabled={busy} icon="picture" size="sm" variant="secondary" onFile={(file) => insert(file, 'floating')}>浮于单元格上方</FileButton>
      </Stack>
    </Dialog>
  );
}
