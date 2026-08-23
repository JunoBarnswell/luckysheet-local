import { useEffect, useState } from 'react';
import { Button, Dialog, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface RenameWorkbookDialogProps {
  open: boolean;
  currentName: string;
  onClose: () => void;
  onSubmit: (name: string) => void;
  submitting?: boolean;
}

export function RenameWorkbookDialog({ open, currentName, onClose, onSubmit, submitting = false }: RenameWorkbookDialogProps) {
  const [name, setName] = useState(currentName);

  useEffect(() => { if (open) setName(currentName); }, [currentName, open]);
  const submit = () => { const next = name.trim(); if (next && !submitting) onSubmit(next); };

  return (
    <Dialog
      closeLabel="关闭重命名工作簿"
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button disabled={!name.trim()} loading={submitting} onClick={submit} size="sm" variant="brand">保存</Button></>}
      maxWidth="sm"
      onClose={onClose}
      open={open}
      title="重命名工作簿"
      testId="rename-workbook-dialog"
    >
      <Stack gap="xs">
        <Text as="label" htmlFor="rename-workbook-name" size="sm" weight="medium">名称</Text>
        <TextInput autoFocus id="rename-workbook-name" onChange={(event) => setName(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') submit(); }} value={name} />
      </Stack>
    </Dialog>
  );
}
