import { useEffect, useState } from 'react';
import { Button, Dialog, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface SaveAsDocumentDialogProps {
  open: boolean;
  currentFileName?: string;
  onClose: () => void;
  onSubmit: (fileName: string) => void;
  submitting?: boolean;
}

const formats = [
  { value: 'ssjson', label: 'SpreadJS SSJSON (.ssjson)' },
  { value: 'sjs', label: 'SpreadJS SJS (.sjs)' },
  { value: 'xlsx', label: 'Excel OOXML (.xlsx)' },
  { value: 'xlsm', label: 'Excel 宏工作簿 (.xlsm)' },
  { value: 'xltx', label: 'Excel 模板 (.xltx)' },
  { value: 'xltm', label: 'Excel 宏模板 (.xltm)' },
  { value: 'ods', label: 'OpenDocument (.ods)' },
  { value: 'csv', label: 'CSV (.csv)' },
  { value: 'xml', label: 'XML Spreadsheet 2003 (.xml)' },
] as const;

function extensionOf(fileName: string): string {
  return fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? 'ssjson';
}

export function SaveAsDocumentDialog({ open, currentFileName = 'workbook.ssjson', onClose, onSubmit, submitting = false }: SaveAsDocumentDialogProps) {
  const [fileName, setFileName] = useState(currentFileName);
  const [format, setFormat] = useState(extensionOf(currentFileName));

  useEffect(() => {
    if (!open) return;
    setFileName(currentFileName);
    setFormat(extensionOf(currentFileName));
  }, [currentFileName, open]);

  const chooseFormat = (next: string) => {
    setFormat(next);
    const stem = fileName.replace(/\.[^.]+$/, '');
    setFileName(`${stem || 'workbook'}.${next}`);
  };
  const submit = () => {
    const trimmed = fileName.trim();
    if (!trimmed || submitting) return;
    onSubmit(trimmed);
  };

  return (
    <Dialog
      closeLabel="关闭另存为对话框"
      description="另存为只创建目标格式副本，不改变当前工作簿的原生身份。"
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button disabled={!fileName.trim()} loading={submitting} onClick={submit} size="sm" variant="brand">导出副本</Button></>}
      maxWidth="sm"
      onClose={onClose}
      open={open}
      title="另存为原生文档"
      testId="save-as-document-dialog"
    >
      <Stack gap="md">
        <Stack gap="xs"><Text as="label" htmlFor="save-as-document-name" size="sm" weight="medium">目标文件名</Text><TextInput id="save-as-document-name" onChange={(event) => setFileName(event.currentTarget.value)} value={fileName} /></Stack>
        <Stack gap="xs"><Text as="label" htmlFor="save-as-document-format" size="sm" weight="medium">目标协议</Text><Select id="save-as-document-format" onChange={(event) => chooseFormat(event.currentTarget.value)} options={formats.map((entry) => ({ value: entry.value, label: entry.label }))} value={format} /></Stack>
        <Text size="xs" tone="muted">跨协议转换会在导出结果中报告可编辑、保留、投影和阻断的特性。</Text>
      </Stack>
    </Dialog>
  );
}
