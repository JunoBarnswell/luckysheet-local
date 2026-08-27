import { useEffect, useState } from 'react';
import type { CellPhoneticMetadata, PhoneticAlignment, PhoneticType } from '@react-sheets/core-model';
import { Button, CheckToggle, Dialog, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { FontFamilyControl } from '../FontFamilyControl';

export interface PhoneticGuideDialogProps {
  open: boolean;
  locale: Locale;
  sourceText: string;
  initial?: CellPhoneticMetadata;
  onClose: () => void;
  onApply: (metadata: CellPhoneticMetadata) => void;
}

export function PhoneticGuideDialog({ open, locale, sourceText, initial, onClose, onApply }: PhoneticGuideDialogProps) {
  const zh = locale === 'zh-CN';
  const [phoneticText, setPhoneticText] = useState('');
  const [visible, setVisible] = useState(true);
  const [type, setType] = useState<PhoneticType>('no-conversion');
  const [alignment, setAlignment] = useState<PhoneticAlignment>('center');
  const [fontFamily, setFontFamily] = useState('Microsoft YaHei');
  const [fontSize, setFontSize] = useState('8');
  useEffect(() => {
    if (!open) return;
    setPhoneticText(initial?.runs.map((run) => run.text).join('') ?? '');
    setVisible(initial?.visible ?? true);
    setType(initial?.type ?? 'no-conversion');
    setAlignment(initial?.alignment ?? 'center');
    setFontFamily(initial?.fontFamily ?? 'Microsoft YaHei');
    setFontSize(String(initial?.fontSizePx ?? 8));
  }, [open, initial]);
  const valid = sourceText.length > 0 && phoneticText.trim().length > 0 && Number(fontSize) >= 6 && Number(fontSize) <= 72;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={zh ? '拼音指南' : 'Phonetic Guide'}
      description={zh ? '编辑东亚文字的拼音或注音，并设置显示、类型、对齐和字体。' : 'Edit East Asian phonetic runs and their display properties.'}
      closeLabel={zh ? '关闭' : 'Close'}
      testId="phonetic-guide-dialog"
      footer={<><Button size="sm" variant="ghost" onClick={onClose}>{zh ? '取消' : 'Cancel'}</Button><Button size="sm" variant="primary" disabled={!valid} onClick={() => { onApply({ visible, type, alignment, fontFamily, fontSizePx: Number(fontSize), runs: [{ text: phoneticText.trim(), start: 0, end: sourceText.length }] }); onClose(); }}>{zh ? '确定' : 'OK'}</Button></>}
    >
      <Stack gap="md">
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '原文字' : 'Base text'}</Text><TextInput value={sourceText} disabled /></Stack>
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '拼音/注音' : 'Phonetic text'}</Text><TextInput aria-label={zh ? '拼音/注音' : 'Phonetic text'} value={phoneticText} onChange={(event) => setPhoneticText(event.target.value)} /></Stack>
        <CheckToggle checked={visible} label={zh ? '显示拼音指南' : 'Show phonetic guide'} onChange={(event) => setVisible(event.target.checked)} />
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '类型' : 'Type'}</Text><Select value={type} onChange={(event) => setType(event.target.value as PhoneticType)} sizeVariant="sm"><option value="no-conversion">{zh ? '不转换' : 'No conversion'}</option><option value="hiragana">Hiragana</option><option value="fullwidth-katakana">{zh ? '全角片假名' : 'Full-width Katakana'}</option><option value="halfwidth-katakana">{zh ? '半角片假名' : 'Half-width Katakana'}</option></Select></Stack>
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '对齐' : 'Alignment'}</Text><Select value={alignment} onChange={(event) => setAlignment(event.target.value as PhoneticAlignment)} sizeVariant="sm"><option value="left">{zh ? '左对齐' : 'Left'}</option><option value="center">{zh ? '居中' : 'Center'}</option><option value="distributed">{zh ? '分散对齐' : 'Distributed'}</option><option value="no-control">{zh ? '无控制' : 'No control'}</option></Select></Stack>
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '字体' : 'Font'}</Text><FontFamilyControl value={fontFamily} fallbackValue="Microsoft YaHei" label={zh ? '字体' : 'Font'} onCommit={setFontFamily} /></Stack>
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '字号' : 'Size'}</Text><TextInput aria-label={zh ? '字号' : 'Size'} inputMode="decimal" value={fontSize} onChange={(event) => setFontSize(event.target.value)} /></Stack>
      </Stack>
    </Dialog>
  );
}
