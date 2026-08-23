import { useEffect, useState } from 'react';
import { Button, Dialog, FileButton, LocationPicker, Stack, Text, type LocationOption } from '@react-sheets/ui-system';

export interface ImportWorkbookDialogValue {
  file: File;
  locationId: string;
}

export interface ImportWorkbookDialogProps {
  open: boolean;
  locationOptions: readonly LocationOption[];
  defaultLocationId?: string;
  onClose: () => void;
  onSubmit: (value: ImportWorkbookDialogValue) => void;
  submitting?: boolean;
}

export function ImportWorkbookDialog({ open, locationOptions, defaultLocationId, onClose, onSubmit, submitting = false }: ImportWorkbookDialogProps) {
  const [file, setFile] = useState<File>();
  const [locationId, setLocationId] = useState(defaultLocationId ?? locationOptions[0]?.id ?? '');

  useEffect(() => {
    if (!open) return;
    setFile(undefined);
    setLocationId(defaultLocationId ?? locationOptions[0]?.id ?? '');
  }, [defaultLocationId, locationOptions, open]);

  const submit = () => {
    if (!file || !locationId || submitting) return;
    onSubmit({ file, locationId });
  };

  return (
    <Dialog
      closeLabel="关闭导入工作簿"
      description="导入后会创建一个全新的工作簿，不会覆盖当前文件"
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button disabled={!file || !locationId} loading={submitting} onClick={submit} size="sm" variant="brand">开始导入</Button></>}
      maxWidth="md"
      onClose={onClose}
      open={open}
      title="导入 Excel 文件"
      testId="import-workbook-dialog"
    >
      <Stack gap="md">
        <Stack gap="sm" className="items-start rounded-lg border border-dashed border-brand-line bg-brand-pale p-5">
          <Text size="sm" weight="medium">选择 .xlsx 文件</Text>
          <Text size="xs" tone="muted">文件大小上限 50 MiB。导入完成后会显示兼容性报告。</Text>
          <FileButton accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" icon="upload" onFile={setFile} size="sm" variant="secondary">选择文件</FileButton>
          {file ? <Text className="max-w-full truncate text-[12px] text-brand-dark" size="sm">已选择：{file.name}</Text> : null}
        </Stack>
        <Stack gap="xs">
          <Text as="label" htmlFor="import-workbook-location" size="sm" weight="medium">导入到</Text>
          <LocationPicker aria-label="导入到" id="import-workbook-location" onChange={(event) => setLocationId(event.target.value)} options={locationOptions} value={locationId} />
        </Stack>
      </Stack>
    </Dialog>
  );
}
