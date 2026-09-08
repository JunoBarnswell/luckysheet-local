import { useEffect, useState } from 'react';
import { Button, Dialog, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface SaveAsDocumentDialogProps {
  open: boolean;
  currentFileName?: string;
  supportedFormats?: readonly SaveAsDocumentFormat[];
  onClose: () => void;
  onSubmit: (fileName: string) => void;
  submitting?: boolean;
}

export type SaveAsDocumentFormat = 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'xlam';

const formatLabels: Record<SaveAsDocumentFormat, string> = {
  xlsx: 'Excel OOXML (.xlsx)',
  xlsm: 'Excel 宏工作簿 (.xlsm)',
  xltx: 'Excel 模板 (.xltx)',
  xltm: 'Excel 宏模板 (.xltm)',
  xlam: 'Excel 加载项 (.xlam)',
};

function extensionOf(fileName: string): string | undefined {
  return fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
}

function fileNameForFormat(fileName: string, format: SaveAsDocumentFormat): string {
  const stem = fileName.replace(/\.[^.]+$/, '');
  return `${stem || 'workbook'}.${format}`;
}

export function SaveAsDocumentDialog({ open, currentFileName = 'workbook.xlsx', supportedFormats = ['xlsx'], onClose, onSubmit, submitting = false }: SaveAsDocumentDialogProps) {
  const [fileName, setFileName] = useState(currentFileName);
  const [format, setFormat] = useState<SaveAsDocumentFormat>(supportedFormats[0] ?? 'xlsx');
  const supportedFormatKey = supportedFormats.join(',');

  useEffect(() => {
    if (!open) return;
    const currentExtension = extensionOf(currentFileName);
    const nextFormat = supportedFormats.includes(currentExtension as SaveAsDocumentFormat)
      ? currentExtension as SaveAsDocumentFormat
      : supportedFormats[0] ?? 'xlsx';
    setFileName(currentFileName);
    setFormat(nextFormat);
    if (currentExtension !== nextFormat) setFileName(fileNameForFormat(currentFileName, nextFormat));
  }, [currentFileName, open, supportedFormatKey]);

  const chooseFormat = (next: string) => {
    if (!supportedFormats.includes(next as SaveAsDocumentFormat)) return;
    const nextFormat = next as SaveAsDocumentFormat;
    setFormat(nextFormat);
    setFileName(fileNameForFormat(fileName, nextFormat));
  };
  const submit = () => {
    const trimmed = fileName.trim();
    if (!trimmed || submitting) return;
    onSubmit(fileNameForFormat(trimmed, format));
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
        <Stack gap="xs"><Text as="label" htmlFor="save-as-document-format" size="sm" weight="medium">目标协议</Text><Select id="save-as-document-format" onChange={(event) => chooseFormat(event.currentTarget.value)} options={supportedFormats.map((entry) => ({ value: entry, label: formatLabels[entry] }))} value={format} /></Stack>
        <Text size="xs" tone="muted">原生导出会在结果中报告可编辑、保留、投影和阻断的特性。</Text>
      </Stack>
    </Dialog>
  );
}
