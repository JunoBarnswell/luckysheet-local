import { useEffect, useState } from 'react';
import { Box, Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CellEditorConfig, CellEditorKind } from '@react-sheets/core-model';

type ConfigurableCellEditorKind = Exclude<CellEditorKind, 'custom'>;

export interface CellEditorDialogProps {
  open: boolean;
  onApply: (editor?: CellEditorConfig) => void;
  onClose: () => void;
}

/** Selection editor configuration; values are committed through sheet.cellEditor.set. */
export function CellEditorDialog({ open, onApply, onClose }: CellEditorDialogProps) {
  const [kind, setKind] = useState<ConfigurableCellEditorKind>('text');
  const [values, setValues] = useState('');
  useEffect(() => { if (open) { setKind('text'); setValues(''); } }, [open]);
  const apply = () => {
    const items = values.split(',').map((value) => value.trim()).filter(Boolean).map((value) => ({ value }));
    const editor: CellEditorConfig = kind === 'datetime'
      ? { kind, mode: 'date' }
      : kind === 'combo-box'
        ? { kind, items, editable: true }
        : kind === 'mask'
          ? { kind, mask: values.trim() || '########' }
          : { kind };
    onApply(editor);
  };
  return (
    <Dialog open={open} title="单元格编辑器" onClose={onClose} maxWidth="sm">
      <Stack gap="md">
        <Text size="sm" tone="muted">编辑器配置应用到当前选区，并作为工作簿数据保存。</Text>
        <Box>
          <Text size="xs" tone="muted">编辑器类型</Text>
          <Select aria-label="单元格编辑器类型" sizeVariant="sm" value={kind} onChange={(event) => setKind(event.target.value as ConfigurableCellEditorKind)}>
            <option value="text">文本</option>
            <option value="number">数字</option>
            <option value="datetime">日期时间</option>
            <option value="validation-list">验证列表</option>
            <option value="combo-box">组合框</option>
            <option value="checkbox">复选框</option>
            <option value="mask">掩码</option>
            <option value="formula">公式</option>
            <option value="rich-text">富文本</option>
          </Select>
        </Box>
        {kind === 'combo-box' ? (
          <Box><Text size="xs" tone="muted">组合框值（逗号分隔）</Text><TextInput aria-label="单元格编辑器组合框值" value={values} onChange={(event) => setValues(event.target.value)} /></Box>
        ) : kind === 'mask' ? (
          <Box><Text size="xs" tone="muted">输入掩码</Text><TextInput aria-label="单元格编辑器输入掩码" value={values} onChange={(event) => setValues(event.target.value)} /></Box>
        ) : null}
        <Inline gap="sm" className="justify-end"><Button size="sm" variant="ghost" onClick={() => onApply(undefined)}>清除编辑器</Button><Button size="sm" variant="primary" onClick={apply}>应用</Button></Inline>
      </Stack>
    </Dialog>
  );
}
