import { useState } from 'react';
import {
  Box,
  Button,
  CheckToggle,
  Inline,
  Panel,
  PanelBody,
  PanelHeader,
  PanelTitle,
  Select,
  Stack,
  StatePanel,
  Text,
  TextInput,
} from '@react-sheets/ui-system';
import type { DataRelationship, TableSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';

export interface TableSheetDesignerPanelProps {
  definition?: TableSheetDefinition;
  tables: readonly WorkbookTableModel[];
  relationships: readonly DataRelationship[];
  onUpdate: (definition: TableSheetDefinition) => void;
}

function cloneDefinition(definition: TableSheetDefinition): TableSheetDefinition {
  return structuredClone(definition);
}

export function TableSheetDesignerPanel({ definition, relationships, tables, onUpdate }: TableSheetDesignerPanelProps) {
  const [query, setQuery] = useState('');
  if (!definition) {
    return <StatePanel kind="error" title="TableSheet definition unavailable" description="The workbook does not contain a canonical TableSheet definition." />;
  }

  const table = tables.find((candidate) => candidate.id === definition.viewId);
  if (!table) {
    return <StatePanel kind="error" title="Binding table unavailable" description={`The binding table ${definition.viewId} is not present in the workbook data model.`} />;
  }

  const visibleIds = new Set(definition.columns.map((column) => column.fieldId));
  const groupedIds = new Set(definition.grouping.map((group) => group.fieldId));
  const sortById = new Map((definition.sortState ?? []).map((sort) => [sort.fieldId, sort.direction]));
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const fields = table.fields.filter((field) => !normalizedQuery || field.name.toLocaleLowerCase().includes(normalizedQuery));
  const tableRelationships = relationships.filter((relationship) => relationship.fromTableId === table.id || relationship.toTableId === table.id);

  const update = (next: TableSheetDefinition) => onUpdate(cloneDefinition(next));
  const setVisible = (fieldId: string, enabled: boolean) => {
    const next = cloneDefinition(definition);
    if (enabled) {
      if (!next.columns.some((column) => column.fieldId === fieldId)) {
        const field = table.fields.find((candidate) => candidate.id === fieldId);
        if (!field) return;
        next.columns.push({ fieldId, caption: field.name, type: field.type });
      }
    } else {
      if (next.columns.length <= 1) return;
      next.columns = next.columns.filter((column) => column.fieldId !== fieldId);
      next.grouping = next.grouping.filter((group) => group.fieldId !== fieldId);
      next.sortState = next.sortState?.filter((sort) => sort.fieldId !== fieldId);
    }
    update(next);
  };
  const toggleGroup = (fieldId: string) => {
    const next = cloneDefinition(definition);
    const index = next.grouping.findIndex((group) => group.fieldId === fieldId);
    if (index >= 0) next.grouping.splice(index, 1);
    else next.grouping.push({ fieldId, collapsed: false });
    update(next);
  };
  const cycleSort = (fieldId: string) => {
    const next = cloneDefinition(definition);
    const index = next.sortState?.findIndex((sort) => sort.fieldId === fieldId) ?? -1;
    if (index < 0) (next.sortState ??= []).push({ fieldId, direction: 'asc' });
    else if (next.sortState![index]!.direction === 'asc') next.sortState![index] = { fieldId, direction: 'desc' };
    else next.sortState!.splice(index, 1);
    if (next.sortState?.length === 0) next.sortState = undefined;
    update(next);
  };
  const bindTable = (viewId: string) => {
    const target = tables.find((candidate) => candidate.id === viewId);
    if (!target) return;
    update({
      viewId,
      columns: target.fields.map((field) => ({ fieldId: field.id, caption: field.name, type: field.type })),
      grouping: [],
      sortState: undefined,
    });
  };

  return (
    <Stack gap="md" data-testid="table-sheet-designer">
      <Panel tone="accent" className="shadow-none">
        <PanelHeader>
          <Stack gap="none">
            <PanelTitle as="h3" size="sm">TableSheet Designer</PanelTitle>
            <Text size="xs" tone="muted">Canonical view settings for {table.name}</Text>
          </Stack>
        </PanelHeader>
        <PanelBody className="space-y-3">
          <Stack gap="xs">
            <Text size="xs" weight="semibold">Binding Table</Text>
            <Select aria-label="Binding Table" sizeVariant="sm" value={definition.viewId} onChange={(event) => bindTable(event.currentTarget.value)} options={tables.map((candidate) => ({ value: candidate.id, label: candidate.name }))} />
          </Stack>
          <TextInput aria-label="Search fields" placeholder="Search fields" value={query} onChange={(event) => setQuery(event.currentTarget.value)} />
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Field tree</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          {fields.map((field) => (
            <Inline key={field.id} gap="xs" className="items-center justify-between rounded border border-slate-100 px-2 py-1.5">
              <CheckToggle aria-label={`Visible ${field.name}`} checked={visibleIds.has(field.id)} onChange={(event) => setVisible(field.id, event.currentTarget.checked)} label={`${field.name} · ${field.type}`} />
              <Inline gap="xs">
                <Button size="xs" variant={groupedIds.has(field.id) ? 'primary' : 'ghost'} onClick={() => toggleGroup(field.id)} aria-label={`Group by ${field.name}`}>Group</Button>
                <Button size="xs" variant={sortById.has(field.id) ? 'soft' : 'ghost'} onClick={() => cycleSort(field.id)} aria-label={`Sort by ${field.name}`}>{sortById.get(field.id) === 'desc' ? 'Z→A' : sortById.has(field.id) ? 'A→Z' : 'Sort'}</Button>
              </Inline>
            </Inline>
          ))}
          {fields.length === 0 ? <Text size="xs" tone="muted">No fields match the search.</Text> : null}
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Visible columns &amp; column setting</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          {definition.columns.map((column) => (
            <Inline key={column.fieldId} gap="xs" className="items-center">
              <Text size="xs" className="min-w-0 flex-1 truncate">{column.caption}</Text>
              <TextInput aria-label={`Width ${column.caption}`} className="w-20" defaultValue={String(column.widthPx ?? '')} placeholder="width" onBlur={(event) => {
                const raw = event.currentTarget.value.trim();
                const widthPx = raw ? Number(raw) : undefined;
                if (widthPx !== undefined && (!Number.isFinite(widthPx) || widthPx <= 0)) return;
                const next = cloneDefinition(definition);
                const target = next.columns.find((entry) => entry.fieldId === column.fieldId);
                if (target) target.widthPx = widthPx;
                update(next);
              }} />
            </Inline>
          ))}
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Relationship hierarchy</PanelTitle></PanelHeader>
        <PanelBody>
          {tableRelationships.length > 0 ? tableRelationships.map((relationship) => (
            <Box key={relationship.id} className="border-b border-slate-100 py-2 last:border-0">
              <Text size="xs" weight="semibold">{relationship.fromTableId} → {relationship.toTableId}</Text>
              <Text size="xs" tone="muted">{relationship.fromFieldId} → {relationship.toFieldId} · {relationship.cardinality}</Text>
            </Box>
          )) : <Text size="xs" tone="muted">No relationships are defined for this binding table.</Text>}
        </PanelBody>
      </Panel>
    </Stack>
  );
}
