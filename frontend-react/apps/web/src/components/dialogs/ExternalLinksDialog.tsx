import { useState } from 'react';
import { Box, Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ExternalLinkBinding } from '@react-sheets/core-model';

export interface ExternalLinksDialogProps {
  open: boolean;
  onClose: () => void;
  onReadSheets: (unitId: string) => Promise<readonly { id: string; name: string }[]>;
  onBind: (link: ExternalLinkBinding) => Promise<void>;
  onRefresh: () => Promise<readonly { token: string; state: string; sourceRevision: number; error?: { code: string; message: string } }[]>;
}
export function ExternalLinksDialog({ open, onClose, onReadSheets, onBind, onRefresh }: ExternalLinksDialogProps) {
  const [sourceUnitId, setSource] = useState('');
  const [token, setToken] = useState('Source.xlsx');
  const [sheets, setSheets] = useState<readonly { id: string; name: string }[]>([]);
  const [sheetId, setSheetId] = useState('');
  const [status, setStatus] = useState<readonly { token: string; state: string; sourceRevision: number; error?: { code: string; message: string } }[]>([]);
  const [busy, setBusy] = useState(false), [error, setError] = useState<string>();
  const run = (work: () => Promise<void>) => { setBusy(true); setError(undefined); void work().catch(cause => setError(cause instanceof Error ? cause.message : String(cause))).finally(() => setBusy(false)); };
  return <Dialog open={open} title="跨工作簿链接" onClose={onClose} maxWidth="md" footer={<Inline gap="sm" className="justify-end"><Button onClick={onClose}>关闭</Button><Button disabled={busy} onClick={() => run(async () => setStatus(await onRefresh()))}>刷新链接</Button><Button variant="primary" disabled={busy || !sheetId || !token.trim()} onClick={() => run(async () => {
    const sheet = sheets.find(sheet => sheet.id === sheetId)!;
    await onBind({ id: `link-${crypto.randomUUID()}`, token: token.trim(), sourceUnitId: sourceUnitId.trim(), sheets: [{ token: sheet.name, sheetId: sheet.id }] });
    setStatus(await onRefresh());
  })}>绑定工作表</Button></Inline>}>
    <Stack gap="md">
      <Text size="sm">绑定后可使用 =SUM('[Source.xlsx]Sales'!B2:B10)。来源工作表更名后，已有绑定仍按工作表身份读取。</Text>
      <TextInput aria-label="来源工作簿 ID" placeholder="来源工作簿 ID" value={sourceUnitId} onChange={event => { setSource(event.target.value); setSheets([]); setSheetId(''); }} />
      <Button disabled={busy || !sourceUnitId.trim()} onClick={() => run(async () => { const result = await onReadSheets(sourceUnitId.trim()); setSheets(result); setSheetId(result[0]?.id ?? ''); })}>读取来源工作表</Button>
      <TextInput aria-label="公式工作簿名称" value={token} onChange={event => setToken(event.target.value)} />
      <Select aria-label="来源工作表" value={sheetId} onChange={event => setSheetId(event.target.value)} options={sheets.map(sheet => ({ value: sheet.id, label: sheet.name }))} />
      {status.map(link => <Box key={link.token}><Text size="sm">{link.token} · {({ connected: '已连接', refreshing: '刷新中', stale: '旧缓存', denied: '无读取权限', unavailable: '来源不可用', broken: '链接已失效' } as Record<string, string>)[link.state] ?? link.state} · 来源修订 {link.sourceRevision}</Text>{link.error ? <Text tone="danger" size="sm">{link.error.code}: {link.error.message}</Text> : null}</Box>)}
      {error ? <Text tone="danger" size="sm">{error}</Text> : null}
    </Stack>
  </Dialog>;
}
