import { useEffect, useState } from 'react';
import { Box, Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CellEditorKind, CellStyleTemplate } from '@react-sheets/core-model';

export interface CellTemplateDialogProps {
  open: boolean;
  templates: readonly CellStyleTemplate[];
  onApply: (templateId: string) => void;
  onClose: () => void;
  onRemove: (templateId: string) => void;
  onSave: (template: CellStyleTemplate) => void;
}

function createTemplateId(name: string): string {
  const normalized = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'template';
  return `template-${normalized}-${Date.now().toString(36)}`;
}

/** Workbook-native template dialog. Form draft is local; saved templates are canonical commands. */
export function CellTemplateDialog({ open, templates, onApply, onClose, onRemove, onSave }: CellTemplateDialogProps) {
  const [name, setName] = useState('强调样式');
  const [background, setBackground] = useState('#e3f1ff');
  const [textColor, setTextColor] = useState('#1f4e79');
  const [editorKind, setEditorKind] = useState<CellEditorKind>('text');
  const [values, setValues] = useState('');

  useEffect(() => {
    if (!open) return;
    setName('强调样式');
    setBackground('#e3f1ff');
    setTextColor('#1f4e79');
    setEditorKind('text');
    setValues('');
  }, [open]);

  const save = () => {
    const trimmed = name.trim();
    if (!trimmed) return;
    const editor = {
      kind: editorKind,
      ...(editorKind === 'list' ? { values: values.split(',').map((value) => value.trim()).filter(Boolean) } : {}),
    } satisfies CellStyleTemplate['editor'];
    onSave({
      id: createTemplateId(trimmed),
      name: trimmed,
      style: { background, textColor, bold: true },
      ...(editor ? { editor } : {}),
      ...(editorKind === 'list' && values.trim() ? {
        dataValidation: {
          type: 'list',
          listSource: { kind: 'values', values: values.split(',').map((value) => value.trim()).filter(Boolean) },
          allowBlank: true,
          showDropdown: true,
        },
      } : {}),
    });
  };

  return (
    <Dialog open={open} title="单元格样式模板" onClose={onClose} maxWidth="lg">
      <Stack gap="md">
        <Text size="sm" tone="muted">模板保存在当前工作簿中，可应用到任意工作表并参与撤销、协作与导出。</Text>
        <Inline gap="sm">
          <Box className="min-w-0 flex-1"><Text size="xs" tone="muted">模板名称</Text><TextInput aria-label="模板名称" value={name} onChange={(event) => setName(event.target.value)} /></Box>
          <Box className="w-32 shrink-0"><Text size="xs" tone="muted">编辑器</Text><Select aria-label="模板编辑器" sizeVariant="sm" value={editorKind} onChange={(event) => setEditorKind(event.target.value as CellEditorKind)}><option value="text">文本</option><option value="number">数字</option><option value="date">日期</option><option value="list">下拉列表</option><option value="checkbox">复选框</option></Select></Box>
        </Inline>
        {editorKind === 'list' ? <Box><Text size="xs" tone="muted">列表值（逗号分隔）</Text><TextInput aria-label="模板列表值" value={values} onChange={(event) => setValues(event.target.value)} placeholder="待处理, 进行中, 已完成" /></Box> : null}
        <Inline gap="sm"><Box className="min-w-0 flex-1"><Text size="xs" tone="muted">填充颜色</Text><TextInput aria-label="模板填充颜色" value={background} onChange={(event) => setBackground(event.target.value)} /></Box><Box className="min-w-0 flex-1"><Text size="xs" tone="muted">字体颜色</Text><TextInput aria-label="模板字体颜色" value={textColor} onChange={(event) => setTextColor(event.target.value)} /></Box></Inline>
        <Inline gap="sm" className="justify-end"><Button size="sm" variant="ghost" onClick={onClose}>关闭</Button><Button size="sm" variant="primary" disabled={!name.trim() || (editorKind === 'list' && values.split(',').map((value) => value.trim()).filter(Boolean).length === 0)} onClick={save}>保存模板</Button></Inline>
        <Box className="border-t border-slate-200 pt-3"><Text size="xs" tone="muted">当前工作簿模板</Text><Stack gap="xs" className="mt-2">{templates.length === 0 ? <Text size="sm" tone="subtle">尚无自定义模板</Text> : templates.map((template) => <Inline key={template.id} gap="sm" className="justify-between rounded border border-slate-200 px-2 py-1"><Text size="sm">{template.name}</Text><Inline gap="xs"><Button size="xs" variant="ghost" onClick={() => onApply(template.id)}>应用</Button><Button aria-label={`删除模板 ${template.name}`} icon="trash" iconOnly size="xs" variant="ghost" className="text-rose-600" onClick={() => onRemove(template.id)} /></Inline></Inline>)}</Stack></Box>
      </Stack>
    </Dialog>
  );
}
