import { useEffect, useState } from 'react';
import { Box, Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CellEditorConfig, CellEditorKind } from '@react-sheets/core-model';

export interface CellEditorDialogProps {
  open: boolean;
  onApply: (editor?: CellEditorConfig) => void;
  onClose: () => void;
}

/** Selection editor configuration; values are committed through sheet.cellEditor.set. */
export function CellEditorDialog({ open, onApply, onClose }: CellEditorDialogProps) {
  const [kind, setKind] = useState<CellEditorKind>('text');
  const [values, setValues] = useState('');
  useEffect(() => { if (open) { setKind('text'); setValues(''); } }, [open]);
  const apply = () => onApply({ kind, ...(kind === 'list' ? { values: values.split(',').map((value) => value.trim()).filter(Boolean) } : {}) });
  return (
    <Dialog open={open} title="单元格编辑器" onClose={onClose} maxWidth="sm">
      <Stack gap="md"><Text size="sm" tone="muted">编辑器配置应用到当前选区，并作为工作簿数据保存。</Text><Box><Text size="xs" tone="muted">编辑器类型</Text><Select aria-label="单元格编辑器类型" sizeVariant="sm" value={kind} onChange={(event) => setKind(event.target.value as CellEditorKind)}><option value="text">文本</option><option value="number">数字</option><option value="date">日期</option><option value="list">下拉列表</option><option value="checkbox">复选框</option></Select></Box>{kind === 'list' ? <Box><Text size="xs" tone="muted">列表值（逗号分隔）</Text><TextInput aria-label="单元格编辑器列表值" value={values} onChange={(event) => setValues(event.target.value)} /></Box> : null}<Inline gap="sm" className="justify-end"><Button size="sm" variant="ghost" onClick={() => onApply(undefined)}>清除编辑器</Button><Button size="sm" variant="primary" onClick={apply}>应用</Button></Inline></Stack>
    </Dialog>
  );
}
