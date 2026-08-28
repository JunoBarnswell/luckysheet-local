import { useEffect, useState } from 'react';
import { Box, Button, Dialog, FileButton, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { LocalObjectDialogKind } from '@react-sheets/spreadsheet-app';

export interface LocalObjectDialogProps {
  open: boolean;
  kind: LocalObjectDialogKind | null;
  onClose: () => void;
  onInsert: (input: { iconName?: string; text?: string; file?: File; layout?: 'process' | 'cycle' | 'hierarchy' | 'list'; signerName?: string; signerTitle?: string; signerEmail?: string; relationship?: 'embedded' | 'linked' }) => Promise<void>;
}

const TITLES: Record<LocalObjectDialogKind, string> = {
  icon: '本地图标', model3d: '本地 3D 模型', smartart: '本地 SmartArt', wordart: '本地艺术字', 'signature-line': '本地签名行', 'embedded-object': '本地文件对象', equation: '本地公式',
};

export function LocalObjectDialog({ open, kind, onClose, onInsert }: LocalObjectDialogProps) {
  const [text, setText] = useState('');
  const [iconName, setIconName] = useState('star');
  const [layout, setLayout] = useState<'process' | 'cycle' | 'hierarchy' | 'list'>('process');
  const [file, setFile] = useState<File>();
  const [signerName, setSignerName] = useState('');
  const [signerTitle, setSignerTitle] = useState('');
  const [signerEmail, setSignerEmail] = useState('');
  const [relationship, setRelationship] = useState<'embedded' | 'linked'>('embedded');
  const [error, setError] = useState<string>();
  useEffect(() => { if (open) { setError(undefined); setFile(undefined); } }, [kind, open]);
  if (!kind) return null;
  const submit = () => { setError(undefined); void onInsert({ iconName, text, file, layout, signerName, signerTitle, signerEmail, relationship }).then(onClose).catch((cause: unknown) => setError(cause instanceof Error ? cause.message : 'LOCAL_OBJECT_INSERT_FAILED')); };
  return (
    <Dialog open={open} title={TITLES[kind]} onClose={onClose} maxWidth="sm" footer={<Inline gap="sm" className="justify-end"><Button size="sm" variant="ghost" onClick={onClose}>取消</Button><Button size="sm" variant="primary" onClick={submit} disabled={kind === 'model3d' || kind === 'embedded-object' ? !file : kind === 'signature-line' ? !signerName.trim() : false}>插入</Button></Inline>}>
      <Stack gap="md">
        <Text size="sm" tone="muted">对象内容由当前工作簿本地模型保存，不调用操作系统、Office 或其他外部宿主。</Text>
        {error ? <Box className="rounded border border-rose-300 bg-rose-50 p-2"><Text size="xs" tone="danger">{error}</Text></Box> : null}
        {kind === 'icon' ? <Stack gap="xs"><Text size="xs" tone="muted">图标</Text><Select sizeVariant="sm" value={iconName} onChange={(event) => setIconName(event.target.value)}><option value="star">星形</option><option value="heart">心形</option><option value="checkmark">勾选</option><option value="circle">圆形</option></Select></Stack> : null}
        {kind === 'model3d' ? <FileButton accept=".obj,text/plain" icon="picture" size="sm" variant="primary" onFile={setFile}>选择本地 OBJ 文件</FileButton> : null}
        {kind === 'embedded-object' ? <Stack gap="sm"><FileButton accept="*/*" icon="table" size="sm" variant="primary" onFile={setFile}>选择本地文件</FileButton><Select sizeVariant="sm" aria-label="对象关系" value={relationship} onChange={(event) => setRelationship(event.target.value as 'embedded' | 'linked')}><option value="embedded">嵌入副本</option><option value="linked">工作簿内本地链接</option></Select></Stack> : null}
        {kind === 'smartart' ? <Stack gap="sm"><Select sizeVariant="sm" aria-label="SmartArt 布局" value={layout} onChange={(event) => setLayout(event.target.value as typeof layout)}><option value="process">流程</option><option value="cycle">循环</option><option value="hierarchy">层级</option><option value="list">列表</option></Select><TextInput aria-label="SmartArt 节点" value={text} onChange={(event) => setText(event.target.value)} placeholder="每行一个节点，例如：开始\n处理\n完成" /></Stack> : null}
        {kind === 'wordart' ? <TextInput aria-label="艺术字内容" value={text} onChange={(event) => setText(event.target.value)} placeholder="输入艺术字内容" /> : null}
        {kind === 'equation' ? <TextInput aria-label="公式表达式" value={text} onChange={(event) => setText(event.target.value)} placeholder="例如 a^2+b^2=c^2" /> : null}
        {kind === 'signature-line' ? <Stack gap="sm"><TextInput aria-label="签名人" value={signerName} onChange={(event) => setSignerName(event.target.value)} placeholder="签名人姓名" /><TextInput aria-label="职务" value={signerTitle} onChange={(event) => setSignerTitle(event.target.value)} placeholder="职务（可选）" /><TextInput aria-label="邮箱" value={signerEmail} onChange={(event) => setSignerEmail(event.target.value)} placeholder="邮箱（可选）" /></Stack> : null}
        {file ? <Text size="xs" tone="muted">已选择：{file.name}</Text> : null}
      </Stack>
    </Dialog>
  );
}
