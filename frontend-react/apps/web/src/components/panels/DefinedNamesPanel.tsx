import { useMemo, useState } from 'react';
import { Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DefinedNameModel } from '@react-sheets/core-model';

export interface DefinedNamesPanelProps {
  sheetId: string;
  names: readonly DefinedNameModel[];
  onSave: (input: DefinedNameModel) => void;
  onRemove: (name: DefinedNameModel) => void;
}

export function DefinedNamesPanel({ names, onRemove, onSave, sheetId }: DefinedNamesPanelProps) {
  const [name, setName] = useState('');
  const [formula, setFormula] = useState('');
  const [scope, setScope] = useState<DefinedNameModel['scope']>('workbook');
  const visible = useMemo(() => [...names].sort((left, right) => left.name.localeCompare(right.name)), [names]);
  const submit = () => {
    const normalizedName = name.trim();
    const normalizedFormula = formula.trim();
    if (!normalizedName || !normalizedFormula) return;
    onSave({ name: normalizedName, formula: normalizedFormula, scope, ...(scope === 'sheet' ? { sheetId } : {}) });
    setName('');
    setFormula('');
  };
  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-line/80 px-4">
        <PanelTitle size="sm">Defined Names</PanelTitle>
      </PanelHeader>
      <PanelBody className="min-h-0 overflow-auto p-4">
        <Stack gap="md">
          <Stack gap="xs" className="rounded-lg border border-line/80 bg-white/70 p-3">
            <Text size="xs" weight="semibold" tone="muted">NEW OR UPDATE NAME</Text>
            <TextInput aria-label="Defined name" placeholder="SalesTotal" value={name} onChange={(event) => setName(event.target.value)} />
            <TextInput aria-label="Defined name formula" placeholder="=Sheet1!$A$1:$A$10" value={formula} onChange={(event) => setFormula(event.target.value)} />
            <Inline gap="xs" className="items-center">
              <Select aria-label="Defined name scope" sizeVariant="sm" value={scope} onChange={(event) => setScope(event.target.value as DefinedNameModel['scope'])}>
                <option value="workbook">Workbook</option>
                <option value="sheet">This worksheet</option>
              </Select>
              <Button size="sm" variant="primary" disabled={!name.trim() || !formula.trim()} onClick={submit}>Save name</Button>
            </Inline>
          </Stack>
          <Stack gap="xs">
            {visible.length === 0 ? <Text size="xs" tone="subtle">No defined names in this workbook.</Text> : null}
            {visible.map((entry) => (
              <Inline key={`${entry.scope}:${entry.sheetId ?? ''}:${entry.name}`} gap="sm" className="items-center rounded-lg border border-line/80 bg-white px-3 py-2">
                <Stack gap="none" className="min-w-0 flex-1">
                  <Text size="xs" weight="semibold" className="truncate">{entry.name}</Text>
                  <Text size="xs" tone="subtle" className="truncate">{entry.formula} · {entry.scope === 'sheet' ? 'Worksheet' : 'Workbook'}</Text>
                </Stack>
                <Button size="xs" variant="ghost" onClick={() => { setName(entry.name); setFormula(entry.formula); setScope(entry.scope); }}>Edit</Button>
                <Button size="xs" variant="danger" onClick={() => onRemove(entry)}>Remove</Button>
              </Inline>
            ))}
          </Stack>
        </Stack>
      </PanelBody>
    </Panel>
  );
}
