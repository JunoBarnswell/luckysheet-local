import { useEffect, useState } from 'react';
import { Button, Dialog, LocationPicker, Stack, StatePanel, Text, TextInput, type LocationOption } from '@react-sheets/ui-system';

export interface CreateWorkbookDialogValue {
  name: string;
  locationId: string;
}

export interface CreateWorkbookDialogProps {
  open: boolean;
  defaultName?: string;
  locationOptions: readonly LocationOption[];
  defaultLocationId?: string;
  error?: string;
  onClose: () => void;
  onSubmit: (value: CreateWorkbookDialogValue) => void;
  submitting?: boolean;
}

export function CreateWorkbookDialog({ open, defaultName = '未命名工作簿', locationOptions, defaultLocationId, error, onClose, onSubmit, submitting = false }: CreateWorkbookDialogProps) {
  const [name, setName] = useState(defaultName);
  const [locationId, setLocationId] = useState(defaultLocationId ?? locationOptions[0]?.id ?? '');

  useEffect(() => {
    if (!open) return;
    setName(defaultName);
    setLocationId(defaultLocationId ?? locationOptions[0]?.id ?? '');
  }, [defaultLocationId, defaultName, locationOptions, open]);

  const submit = () => {
    const trimmedName = name.trim();
    if (!trimmedName || !locationId || submitting) return;
    onSubmit({ name: trimmedName, locationId });
  };

  return (
    <Dialog
      closeLabel="关闭新建工作簿"
      description="选择工作簿名称和保存位置"
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button disabled={!name.trim() || !locationId} loading={submitting} onClick={submit} size="sm" variant="brand">创建工作簿</Button></>}
      maxWidth="sm"
      onClose={onClose}
      open={open}
      title="新建工作簿"
      testId="create-workbook-dialog"
    >
      <Stack gap="md">
        {error ? <StatePanel kind="error" title="无法创建工作簿" description={error} /> : null}
        <Stack gap="xs">
          <Text as="label" htmlFor="workbook-name" size="sm" weight="medium">工作簿名称</Text>
          <TextInput autoFocus id="workbook-name" onChange={(event) => setName(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') submit(); }} value={name} />
        </Stack>
        <Stack gap="xs">
          <Text as="label" htmlFor="workbook-location" size="sm" weight="medium">保存位置</Text>
          <LocationPicker aria-label="保存位置" id="workbook-location" onChange={(event) => setLocationId(event.target.value)} options={locationOptions} value={locationId} />
          <Text size="xs" tone="subtle">创建后仍可在文件中心移动工作簿。</Text>
        </Stack>
      </Stack>
    </Dialog>
  );
}
