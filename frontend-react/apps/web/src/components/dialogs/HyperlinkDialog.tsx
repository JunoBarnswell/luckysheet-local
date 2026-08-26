import React, { useEffect, useState } from 'react';
import type { CellHyperlink, DefinedNameModel, HyperlinkTarget } from '@react-sheets/core-model';
import { Button, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface HyperlinkSheetOption {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
}

export interface HyperlinkDialogProps {
  open: boolean;
  initial?: CellHyperlink;
  sheets: readonly HyperlinkSheetOption[];
  definedNames: readonly DefinedNameModel[];
  onClose: () => void;
  onApply: (target: HyperlinkTarget, tooltip?: string) => void;
  onRemove?: () => void;
}

type TargetKind = HyperlinkTarget['kind'];

function initialKind(link?: CellHyperlink): TargetKind {
  return link?.target.kind ?? 'url';
}

function initialSheet(link: CellHyperlink | undefined, sheets: readonly HyperlinkSheetOption[]): string {
  if (link?.target.kind === 'sheet' && sheets.some((sheet) => sheet.id === link.target.sheetId)) return link.target.sheetId;
  return sheets[0]?.id ?? '';
}

function initialAddress(link?: CellHyperlink): string {
  if (link?.target.kind !== 'sheet') return '';
  if (link.target.address) return link.target.address;
  if (link.target.row !== undefined && link.target.column !== undefined) {
    let column = link.target.column + 1;
    let label = '';
    while (column > 0) {
      const remainder = (column - 1) % 26;
      label = String.fromCharCode(65 + remainder) + label;
      column = Math.floor((column - 1) / 26);
    }
    return `${label}${link.target.row + 1}`;
  }
  return '';
}

/** Typed hyperlink authoring boundary; it never exposes internal serialization prefixes. */
export function HyperlinkDialog({ open, initial, sheets, definedNames, onClose, onApply, onRemove }: HyperlinkDialogProps): React.ReactElement | null {
  const [kind, setKind] = useState<TargetKind>('url');
  const [url, setUrl] = useState('');
  const [email, setEmail] = useState('');
  const [subject, setSubject] = useState('');
  const [sheetId, setSheetId] = useState('');
  const [address, setAddress] = useState('');
  const [name, setName] = useState('');
  const [tooltip, setTooltip] = useState('');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    const target = initial?.target;
    setKind(initialKind(initial));
    setUrl(target?.kind === 'url' ? target.url : '');
    setEmail(target?.kind === 'email' ? target.address : '');
    setSubject(target?.kind === 'email' ? target.subject ?? '' : '');
    setSheetId(initialSheet(initial, sheets));
    setAddress(initialAddress(initial));
    setName(target?.kind === 'name' ? target.name : '');
    setTooltip(initial?.tooltip ?? '');
    setError(null);
  }, [open, initial, sheets]);

  if (!open) return null;

  const buildTarget = (): HyperlinkTarget => {
    if (kind === 'url') {
      const value = url.trim();
      let parsed: URL;
      try { parsed = new URL(value); } catch { throw new Error('Enter a valid web URL.'); }
      if (!['http:', 'https:', 'ftp:'].includes(parsed.protocol)) throw new Error('Use an HTTP, HTTPS, or FTP URL.');
      return { kind, url: value };
    }
    if (kind === 'email') {
      const value = email.trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) throw new Error('Enter a valid email address.');
      return { kind, address: value, ...(subject.trim() ? { subject: subject.trim() } : {}) };
    }
    if (kind === 'sheet') {
      if (!sheetId) throw new Error('Choose a worksheet.');
      if (!address.trim()) throw new Error('Enter a worksheet address such as A1.');
      return { kind, sheetId, address: address.trim() };
    }
    const value = name.trim();
    if (!value) throw new Error('Choose a defined name.');
    if (!definedNames.some((entry) => entry.name.toLowerCase() === value.toLowerCase())) throw new Error(`Defined name not found: ${value}`);
    return { kind, name: value };
  };

  const apply = () => {
    try {
      const target = buildTarget();
      onApply(target, tooltip.trim() || undefined);
      setError(null);
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason));
    }
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={initial ? 'Edit Hyperlink' : 'Insert Hyperlink'}
      description="Choose a typed destination. Internal worksheet and defined-name references are selected here and are never entered as serialized prefixes."
      closeLabel="Close hyperlink dialog"
      testId="hyperlink-dialog"
      maxWidth="md"
      footer={(
        <Inline gap="sm" className="justify-end">
          <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
          {initial && onRemove ? <Button size="sm" variant="ghost" onClick={() => { onRemove(); onClose(); }}>Remove</Button> : null}
          <Button size="sm" variant="primary" data-testid="hyperlink-apply" onClick={apply}>Apply</Button>
        </Inline>
      )}
    >
      <Stack gap="sm">
        <Select aria-label="Hyperlink destination type" data-testid="hyperlink-kind" value={kind} onChange={(event) => { setKind(event.target.value as TargetKind); setError(null); }}>
          <option value="url">Web page</option>
          <option value="email">Email address</option>
          <option value="sheet">Place in this workbook</option>
          <option value="name">Defined name</option>
        </Select>
        {kind === 'url' ? <TextInput aria-label="Web URL" data-testid="hyperlink-url" placeholder="https://example.com" value={url} onChange={(event) => setUrl(event.target.value)} /> : null}
        {kind === 'email' ? <>
          <TextInput aria-label="Email address" data-testid="hyperlink-email" placeholder="person@example.com" value={email} onChange={(event) => setEmail(event.target.value)} />
          <TextInput aria-label="Email subject" data-testid="hyperlink-subject" placeholder="Subject (optional)" value={subject} onChange={(event) => setSubject(event.target.value)} />
        </> : null}
        {kind === 'sheet' ? <>
          <Select aria-label="Target worksheet" data-testid="hyperlink-sheet" value={sheetId} onChange={(event) => setSheetId(event.target.value)}>
            {sheets.map((sheet) => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}
          </Select>
          <TextInput aria-label="Worksheet address" data-testid="hyperlink-address" placeholder="A1" value={address} onChange={(event) => setAddress(event.target.value)} />
        </> : null}
        {kind === 'name' ? <Select aria-label="Defined name" data-testid="hyperlink-name" value={name} onChange={(event) => setName(event.target.value)}>
          <option value="">Choose a defined name</option>
          {definedNames.map((entry) => <option key={`${entry.scope}:${entry.sheetId ?? ''}:${entry.name}`} value={entry.name}>{entry.name}{entry.scope === 'sheet' ? ' (sheet)' : ''}</option>)}
        </Select> : null}
        <TextInput aria-label="ScreenTip" data-testid="hyperlink-tooltip" placeholder="ScreenTip (optional)" value={tooltip} onChange={(event) => setTooltip(event.target.value)} />
        {error ? <Text role="alert" size="sm" tone="danger">{error}</Text> : null}
      </Stack>
    </Dialog>
  );
}

export default HyperlinkDialog;
