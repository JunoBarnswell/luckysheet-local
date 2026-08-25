import { useEffect, useState } from 'react';
import { Button, Dialog, Inline, RadioOption, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { RangeRef } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';

export interface CreateTableDialogProps {
  open: boolean;
  locale: Locale;
  sourceRange: RangeRef;
  onClose: () => void;
  onCreate: (request: { name: string; hasHeaderRow: boolean; styleName?: string }) => void;
}

function columnLabel(column: number): string {
  let result = '';
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  return result;
}

function rangeLabel(range: RangeRef): string {
  return `${columnLabel(range.startColumn)}${range.startRow + 1}:${columnLabel(range.endColumn)}${range.endRow + 1}`;
}

export function CreateTableDialog({ open, locale, sourceRange, onClose, onCreate }: CreateTableDialogProps) {
  const [name, setName] = useState('Table1');
  const [hasHeaderRow, setHasHeaderRow] = useState(true);
  const [styleName, setStyleName] = useState('TableStyleMedium2');
  const isZh = locale === 'zh-CN';
  useEffect(() => {
    if (open) {
      setName('Table1');
      setHasHeaderRow(true);
      setStyleName('TableStyleMedium2');
    }
  }, [open, sourceRange.startRow, sourceRange.startColumn, sourceRange.endRow, sourceRange.endColumn]);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={isZh ? '创建表格' : 'Create Table'}
      testId="create-table-dialog"
      footer={<Inline gap="sm" className="w-full justify-end"><Button size="sm" variant="primary" data-testid="create-table-confirm" disabled={!name.trim()} onClick={() => onCreate({ name: name.trim(), hasHeaderRow, styleName: styleName.trim() || undefined })}>{isZh ? '确定' : 'OK'}</Button><Button size="sm" variant="outline" onClick={onClose}>{isZh ? '取消' : 'Cancel'}</Button></Inline>}
    >
      <Stack gap="md">
        <Stack gap="xs"><Text size="sm" weight="medium">{isZh ? '表格范围' : 'Table range'}</Text><TextInput aria-label={isZh ? '表格范围' : 'Table range'} readOnly value={rangeLabel(sourceRange)} className="font-mono" /></Stack>
        <Stack gap="xs"><Text size="sm" weight="medium">{isZh ? '表格名称' : 'Table name'}</Text><TextInput aria-label={isZh ? '表格名称' : 'Table name'} data-testid="create-table-name" value={name} onChange={(event) => setName(event.target.value)} /></Stack>
        <Stack gap="xs"><Text size="sm" weight="medium">{isZh ? '表头' : 'Headers'}</Text><RadioOption checked={hasHeaderRow} label={isZh ? '我的表包含标题' : 'My table has headers'} name="table-header-mode" onChange={() => setHasHeaderRow(true)} /><RadioOption checked={!hasHeaderRow} label={isZh ? '第一行是数据' : 'The first row is data'} name="table-header-mode" onChange={() => setHasHeaderRow(false)} /></Stack>
        <Stack gap="xs"><Text size="sm" weight="medium">{isZh ? '表格样式' : 'Table style'}</Text><TextInput aria-label={isZh ? '表格样式' : 'Table style'} value={styleName} onChange={(event) => setStyleName(event.target.value)} placeholder="TableStyleMedium2" /></Stack>
      </Stack>
    </Dialog>
  );
}
