import { useState } from 'react';
import { Box, Button, CheckToggle, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, StatePanel, Text, TextInput } from '@react-sheets/ui-system';
import type { ReportBinding, ReportSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';

export interface ReportDesignerPanelProps {
  definition?: ReportSheetDefinition;
  tables: readonly WorkbookTableModel[];
  activeCell: string;
  onUpdate: (definition: ReportSheetDefinition) => void;
}

function cellAddress(value: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(value.trim().toUpperCase());
  if (!match) return undefined;
  let column = 0;
  for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

export function ReportDesignerPanel({ definition, tables, activeCell, onUpdate }: ReportDesignerPanelProps) {
  const [query, setQuery] = useState('');
  if (!definition) return <StatePanel kind="error" title="ReportSheet definition unavailable" description="The workbook does not contain a canonical ReportSheet definition." />;
  const table = definition.tableId ? tables.find((candidate) => candidate.id === definition.tableId) : undefined;
  if (definition.tableId && !table) return <StatePanel kind="error" title="Binding table unavailable" description={`The binding table ${definition.tableId} is not present in the workbook data model.`} />;
  const fields = (table?.fields ?? []).filter((field) => !query.trim() || field.name.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()));
  const update = (next: ReportSheetDefinition) => onUpdate(structuredClone(next));
  const bindField = (fieldId: string) => {
    const cell = cellAddress(activeCell);
    if (!cell) return;
    const next = structuredClone(definition);
    next.bindings = next.bindings.filter((binding) => binding.cell.row !== cell.row || binding.cell.column !== cell.column);
    const binding: ReportBinding = { cell, expression: fieldId, kind: 'field', direction: 'vertical', fill: 'down' };
    next.bindings.push(binding);
    update(next);
  };
  const setMode = (renderMode: ReportSheetDefinition['renderMode']) => update({ ...structuredClone(definition), renderMode });
  return (
    <Stack gap="md" data-testid="report-sheet-designer">
      <Panel tone="accent" className="shadow-none">
        <PanelHeader><Stack gap="none"><PanelTitle as="h3" size="sm">ReportSheet Designer</PanelTitle><Text size="xs" tone="muted">Template and data-entry configuration</Text></Stack></PanelHeader>
        <PanelBody className="space-y-3">
          <Select aria-label="Report binding table" sizeVariant="sm" value={definition.tableId ?? ''} onChange={(event) => update({ ...structuredClone(definition), tableId: event.currentTarget.value || undefined, bindings: [], dataEntry: [] })} options={[{ value: '', label: 'No binding table' }, ...tables.map((candidate) => ({ value: candidate.id, label: candidate.name }))]} />
          <Select aria-label="Report render mode" sizeVariant="sm" value={definition.renderMode} onChange={(event) => setMode(event.currentTarget.value as ReportSheetDefinition['renderMode'])} options={[{ value: 'design', label: 'Design' }, { value: 'preview', label: 'Preview' }, { value: 'paginated', label: 'Paginated' }]} />
          <TextInput aria-label="Search report fields" placeholder="Search fields" value={query} onChange={(event) => setQuery(event.currentTarget.value)} />
        </PanelBody>
      </Panel>
      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Report fields</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          <Text size="xs" tone="muted">Select a template cell ({activeCell}), then bind a field.</Text>
          {fields.map((field) => <Inline key={field.id} gap="sm" className="items-center justify-between rounded border border-slate-100 px-2 py-1.5"><Text size="xs">{field.name} · {field.type}</Text><Button size="xs" variant="soft" onClick={() => bindField(field.id)}>Bind</Button></Inline>)}
          {fields.length === 0 ? <Text size="xs" tone="muted">No fields match the search.</Text> : null}
        </PanelBody>
      </Panel>
      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Pagination and layout</PanelTitle></PanelHeader>
        <PanelBody className="space-y-3">
          <CheckToggle aria-label="Enable report pagination" checked={definition.pagination.enabled} label="Enable pagination" onChange={(event) => update({ ...structuredClone(definition), pagination: { ...definition.pagination, enabled: event.currentTarget.checked } })} />
          <TextInput aria-label="Report rows per page" type="number" min="1" value={String(definition.pagination.rowsPerPage ?? 50)} onChange={(event) => update({ ...structuredClone(definition), pagination: { ...definition.pagination, rowsPerPage: Number(event.currentTarget.value) } })} />
          <Select aria-label="Report orientation" sizeVariant="sm" value={definition.layout.orientation} onChange={(event) => update({ ...structuredClone(definition), layout: { ...definition.layout, orientation: event.currentTarget.value as ReportSheetDefinition['layout']['orientation'] } })} options={[{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }]} />
        </PanelBody>
      </Panel>
      <Box className="rounded border border-slate-200 bg-slate-50 px-3 py-2"><Text size="xs" tone="muted">Preview and paginated cells are derived overlays. They never become a second cell-storage owner.</Text></Box>
    </Stack>
  );
}
