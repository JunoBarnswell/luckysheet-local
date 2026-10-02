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
  onConfigureRecord?: (input: { tableId: string; identityFieldId: string; fieldId?: string; calculation?: import('@react-sheets/core-model').RecordFieldCalculation }) => void;
  onSetRelationship?: (relationship: DataRelationship) => void;
  onUpdate: (definition: TableSheetDefinition) => void;
}

function cloneDefinition(definition: TableSheetDefinition): TableSheetDefinition {
  return structuredClone(definition);
}

export function TableSheetDesignerPanel({ definition, relationships, tables, onUpdate, onConfigureRecord, onSetRelationship }: TableSheetDesignerPanelProps) {
  const [query, setQuery] = useState('');
  const [identityFieldId, setIdentity] = useState('');
  const [calculatedFieldId, setCalculatedField] = useState('');
  const [calculationKind, setCalculationKind] = useState<'formula' | 'lookup' | 'rollup'>('formula');
  const [formula, setFormula] = useState('');
  const [relationshipId, setRelationshipId] = useState('');
  const [targetFieldId, setTargetFieldId] = useState('');
  const [aggregate, setAggregate] = useState<'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'>('SUM');
  const [foreignFieldId, setForeignFieldId] = useState('');
  const [targetTableId, setTargetTableId] = useState('');
  const [error, setError] = useState<string>();
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

  const chosenRelation = tableRelationships.find(relation => relation.id === relationshipId);
  const direction = chosenRelation?.fromTableId === table.id ? 'forward' : 'reverse';
  const lookupTable = chosenRelation ? tables.find(table => table.id === (direction === 'forward' ? chosenRelation.toTableId : chosenRelation.fromTableId)) : undefined;
  const perform = (work: () => void) => { setError(undefined); try { work(); } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); } };
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

      {onConfigureRecord ? <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Record / Field 计算</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          <Text size="xs">记录 ID 字段在排序后保持不变。首次启用会为该字段的空值分配 ID。</Text>
          <Select aria-label="记录 ID 字段" value={identityFieldId || table.recordIdFieldId || ''} onChange={event => setIdentity(event.target.value)} options={[{ value: '', label: '选择记录 ID 字段' }, ...table.fields.filter(field => field.type === 'text' && !field.calculation).map(field => ({ value: field.id, label: field.name }))]} />
          <Button onClick={() => perform(() => onConfigureRecord({ tableId: table.id, identityFieldId: identityFieldId || table.recordIdFieldId || '' }))}>启用记录身份</Button>
          <Select aria-label="计算字段" value={calculatedFieldId} onChange={event => setCalculatedField(event.target.value)} options={[{ value: '', label: '选择计算字段' }, ...table.fields.filter(field => field.id !== table.recordIdFieldId).map(field => ({ value: field.id, label: field.name }))]} />
          <Select aria-label="计算方式" value={calculationKind} onChange={event => setCalculationKind(event.target.value as typeof calculationKind)} options={[{ value: 'formula', label: '字段公式' }, { value: 'lookup', label: 'Lookup 关联值' }, { value: 'rollup', label: 'Rollup 汇总' }]} />
          {calculationKind === 'formula' ? <TextInput aria-label="字段公式" placeholder="=[数量]*[单价]" value={formula} onChange={event => setFormula(event.target.value)} /> : <>
            <Select aria-label="计算关联" value={relationshipId} onChange={event => { setRelationshipId(event.target.value); setTargetFieldId(''); }} options={[{ value: '', label: '选择关联' }, ...tableRelationships.map(relation => ({ value: relation.id, label: `${tables.find(table => table.id === relation.fromTableId)?.name} → ${tables.find(table => table.id === relation.toTableId)?.name}` }))]} />
            <Select aria-label="关联取值字段" value={targetFieldId} onChange={event => setTargetFieldId(event.target.value)} options={[{ value: '', label: '选择取值字段' }, ...(lookupTable?.fields ?? []).map(field => ({ value: field.id, label: field.name }))]} />
            {calculationKind === 'rollup' ? <Select aria-label="汇总函数" value={aggregate} onChange={event => setAggregate(event.target.value as typeof aggregate)} options={['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX'].map(value => ({ value, label: value }))} /> : null}
          </>}
          <Button variant="primary" disabled={!calculatedFieldId || calculationKind === 'formula' && !formula.trim() || calculationKind !== 'formula' && (!relationshipId || !targetFieldId)} onClick={() => perform(() => onConfigureRecord({ tableId: table.id, identityFieldId: identityFieldId || table.recordIdFieldId || '', fieldId: calculatedFieldId, calculation: calculationKind === 'formula' ? { kind: 'formula', formula } : calculationKind === 'lookup' ? { kind: 'lookup', relationshipId, targetFieldId, direction } : { kind: 'rollup', relationshipId, targetFieldId, direction, aggregate } }))}>保存计算字段</Button>
          {error ? <Text tone="danger" size="xs">{error}</Text> : null}
        </PanelBody>
      </Panel> : null}
      {onSetRelationship && table.recordIdFieldId ? <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">创建记录关联</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          <Select aria-label="关联 ID 字段" value={foreignFieldId} onChange={event => setForeignFieldId(event.target.value)} options={[{ value: '', label: '存放目标记录 ID 的字段' }, ...table.fields.filter(field => !field.calculation).map(field => ({ value: field.id, label: field.name }))]} />
          <Select aria-label="关联目标表" value={targetTableId} onChange={event => setTargetTableId(event.target.value)} options={[{ value: '', label: '选择目标表' }, ...tables.filter(table => table.recordIdFieldId).map(table => ({ value: table.id, label: table.name }))]} />
          <Button disabled={!foreignFieldId || !targetTableId} onClick={() => perform(() => onSetRelationship({ id: `relation-${crypto.randomUUID()}`, fromTableId: table.id, fromFieldId: foreignFieldId, toTableId: targetTableId, toFieldId: tables.find(table => table.id === targetTableId)!.recordIdFieldId!, cardinality: 'many-to-one' }))}>保存关联</Button>
        </PanelBody>
      </Panel> : null}
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
