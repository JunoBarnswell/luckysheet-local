import { useState } from 'react';
import { Box, Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';

export interface SymbolDialogProps {
  open: boolean;
  locale: Locale;
  recent: readonly string[];
  onClose: () => void;
  onInsert: (symbol: string) => string | undefined;
}

const COMMON_SYMBOLS = ['©', '®', '™', '°', '±', '×', '÷', '≠', '≤', '≥', '∞', '∑', '√', 'π', 'Ω', 'µ', '€', '£', '¥', '→', '←', '↑', '↓', '✓', '★'] as const;

export function SymbolDialog({ open, locale, recent, onClose, onInsert }: SymbolDialogProps) {
  const zh = locale === 'zh-CN';
  const [codePoint, setCodePoint] = useState('');
  const [error, setError] = useState<string | undefined>();
  const insert = (symbol: string) => {
    const failure = onInsert(symbol);
    setError(failure);
    if (!failure) onClose();
  };
  const parsed = /^[0-9a-f]{1,6}$/i.test(codePoint) ? Number.parseInt(codePoint, 16) : Number.NaN;
  const codePointSymbol = Number.isInteger(parsed) && parsed >= 0 && parsed <= 0x10ffff ? String.fromCodePoint(parsed) : undefined;
  return (
    <Dialog open={open} onClose={onClose} title={zh ? '符号' : 'Symbol'} description={zh ? '插入 Unicode 符号到当前单元格或文本框编辑目标。' : 'Insert a Unicode symbol into the active cell or text-box editor.'} closeLabel={zh ? '关闭' : 'Close'} testId="symbol-dialog" footer={<Button size="sm" variant="ghost" onClick={onClose}>{zh ? '取消' : 'Cancel'}</Button>}>
      <Stack gap="md">
        {error ? <Box className="rounded border border-rose-300 bg-rose-50 p-2"><Text size="xs" tone="danger">{error}</Text></Box> : null}
        {recent.length ? <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '最近使用' : 'Recently used'}</Text><Inline gap="xs" className="flex-wrap">{recent.map((symbol) => <Button key={symbol} size="sm" variant="ghost" onClick={() => insert(symbol)}>{symbol}</Button>)}</Inline></Stack> : null}
        <Stack gap="xs"><Text size="xs" tone="muted">{zh ? '常用符号' : 'Common symbols'}</Text><Inline gap="xs" className="flex-wrap">{COMMON_SYMBOLS.map((symbol) => <Button key={symbol} size="sm" variant="ghost" onClick={() => insert(symbol)}>{symbol}</Button>)}</Inline></Stack>
        <Stack gap="xs"><Text size="xs" tone="muted">Unicode code point</Text><Inline gap="sm"><TextInput aria-label="Unicode code point" value={codePoint} onChange={(event) => setCodePoint(event.target.value.replace(/^U\+/i, ''))} placeholder="03A9" /><Button size="sm" variant="primary" disabled={!codePointSymbol} onClick={() => { if (codePointSymbol) insert(codePointSymbol); }}>{zh ? '插入' : 'Insert'}</Button></Inline></Stack>
      </Stack>
    </Dialog>
  );
}
