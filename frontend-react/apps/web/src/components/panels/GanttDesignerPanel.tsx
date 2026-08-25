import { useState } from 'react';
import {
  Box,
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
import type { GanttSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';

export interface GanttDesignerPanelProps {
  definition?: GanttSheetDefinition;
  tables: readonly WorkbookTableModel[];
  onUpdate: (definition: GanttSheetDefinition) => void;
}

const mappingKeys = ['id', 'title', 'start', 'end', 'progress', 'parentId', 'dependencies'] as const;
type MappingKey = typeof mappingKeys[number];

export function GanttDesignerPanel({ definition, tables, onUpdate }: GanttDesignerPanelProps) {
  const [query, setQuery] = useState('');
  if (!definition) return <StatePanel kind="error" title="GanttSheet definition unavailable" description="The workbook does not contain a canonical GanttSheet definition." />;
  const table = tables.find((candidate) => candidate.id === definition.viewId);
  if (!table) return <StatePanel kind="error" title="Binding table unavailable" description={`The binding table ${definition.viewId} is not present in the workbook data model.`} />;
  const fields = table.fields.filter((field) => !query.trim() || field.name.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()));
  const update = (next: GanttSheetDefinition) => onUpdate(structuredClone(next));
  const setMapping = (key: MappingKey, fieldId: string) => {
    const next = structuredClone(definition);
    if (key === 'parentId' || key === 'dependencies') {
      if (fieldId) next.fieldMap[key] = fieldId;
      else delete next.fieldMap[key];
    } else next.fieldMap[key] = fieldId;
    update(next);
  };
  const setCalendar = (patch: Partial<GanttSheetDefinition['calendar']>) => update({ ...structuredClone(definition), calendar: { ...definition.calendar, ...patch } });
  const setTimeline = (patch: Partial<GanttSheetDefinition['timeline']>) => update({ ...structuredClone(definition), timeline: { ...definition.timeline, ...patch } });
  const selectedDay = (day: number) => definition.calendar.workingDays.includes(day);
  return (
    <Stack gap="md" data-testid="gantt-sheet-designer">
      <Panel tone="accent" className="shadow-none">
        <PanelHeader><Stack gap="none"><PanelTitle as="h3" size="sm">GanttSheet Designer</PanelTitle><Text size="xs" tone="muted">Canonical task mapping for {table.name}</Text></Stack></PanelHeader>
        <PanelBody className="space-y-3">
          <Text size="xs" weight="semibold">Binding Table</Text>
          <Select aria-label="Gantt binding table" sizeVariant="sm" value={definition.viewId} onChange={(event) => {
            const target = tables.find((candidate) => candidate.id === event.currentTarget.value);
            if (!target) return;
            const ids = new Set(target.fields.map((field) => field.id));
            const next = structuredClone(definition);
            next.viewId = target.id;
            for (const key of mappingKeys) {
              const value = next.fieldMap[key];
              if (value && !ids.has(value)) delete next.fieldMap[key];
            }
            update(next);
          }} options={tables.map((candidate) => ({ value: candidate.id, label: candidate.name }))} />
          <TextInput aria-label="Search Gantt fields" placeholder="Search fields" value={query} onChange={(event) => setQuery(event.currentTarget.value)} />
        </PanelBody>
      </Panel>
      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Task field mapping</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2">
          {mappingKeys.map((key) => (
            <Inline key={key} gap="sm" className="items-center justify-between">
              <Text size="xs" className="capitalize">{key}</Text>
              <Select aria-label={`Gantt ${key} field`} sizeVariant="sm" value={definition.fieldMap[key] ?? ''} onChange={(event) => setMapping(key, event.currentTarget.value)} options={[...(key === 'parentId' || key === 'dependencies' ? [{ value: '', label: 'Not mapped' }] : []), ...fields.map((field) => ({ value: field.id, label: `${field.name} · ${field.type}` }))]} />
            </Inline>
          ))}
        </PanelBody>
      </Panel>
      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Calendar and timeline</PanelTitle></PanelHeader>
        <PanelBody className="space-y-3">
          <Inline gap="xs" className="flex-wrap">{[0, 1, 2, 3, 4, 5, 6].map((day) => <CheckToggle key={day} aria-label={`Working day ${day}`} checked={selectedDay(day)} label={['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][day]!} onChange={(event) => setCalendar({ workingDays: event.currentTarget.checked ? [...new Set([...definition.calendar.workingDays, day])].sort() : definition.calendar.workingDays.filter((item) => item !== day) })} />)}</Inline>
          <Inline gap="sm"><TextInput aria-label="Day start hour" type="number" min="0" max="23" value={String(definition.calendar.dayStartHour)} onChange={(event) => setCalendar({ dayStartHour: Number(event.currentTarget.value) })} /><TextInput aria-label="Day end hour" type="number" min="1" max="24" value={String(definition.calendar.dayEndHour)} onChange={(event) => setCalendar({ dayEndHour: Number(event.currentTarget.value) })} /></Inline>
          <Select aria-label="Gantt timeline unit" sizeVariant="sm" value={definition.timeline.unit} onChange={(event) => setTimeline({ unit: event.currentTarget.value as GanttSheetDefinition['timeline']['unit'] })} options={['day', 'week', 'month', 'quarter'].map((value) => ({ value, label: value }))} />
          <Inline gap="sm"><TextInput aria-label="Timeline start" type="date" value={definition.timeline.start ?? ''} onChange={(event) => setTimeline({ start: event.currentTarget.value || undefined })} /><TextInput aria-label="Timeline end" type="date" value={definition.timeline.end ?? ''} onChange={(event) => setTimeline({ end: event.currentTarget.value || undefined })} /></Inline>
        </PanelBody>
      </Panel>
      <Panel className="shadow-none">
        <PanelHeader><PanelTitle as="h3" size="sm">Dependency format</PanelTitle></PanelHeader>
        <PanelBody className="space-y-2"><Inline gap="sm"><TextInput aria-label="Dependency color" type="color" value={definition.dependencyStyle.color} onChange={(event) => update({ ...structuredClone(definition), dependencyStyle: { ...definition.dependencyStyle, color: event.currentTarget.value } })} /><TextInput aria-label="Dependency width" type="number" min="0.5" step="0.5" value={String(definition.dependencyStyle.width)} onChange={(event) => update({ ...structuredClone(definition), dependencyStyle: { ...definition.dependencyStyle, width: Number(event.currentTarget.value) } })} /></Inline></PanelBody>
      </Panel>
      <Box className="rounded border border-slate-200 bg-slate-50 px-3 py-2"><Text size="xs" tone="muted">Task bars and dependencies are derived from the bound table. Invalid mappings are rejected by the canonical command.</Text></Box>
    </Stack>
  );
}
