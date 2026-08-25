import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DataChartBindingArea, DataChartDrawingPayload, DataChartFieldBinding, DataChartPlotType, DrawingObject, DrawingPayload, RangeRef, WorkbookTableModel } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import { parseRangeInput } from '../../domain/range-input';

export interface DataChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  sheet: CanvasSheetSnapshot;
  tables: readonly WorkbookTableModel[];
  onCommand: (descriptor: CommandDescriptor) => void;
}

const plotTypes: readonly DataChartPlotType[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar', 'treemap', 'funnel'];
const bindingAreas: readonly DataChartBindingArea[] = ['values', 'category', 'details', 'color', 'size', 'tooltip', 'filter'];
const areaLabels: Record<DataChartBindingArea, string> = { values: 'Values', category: 'Category', details: 'Details', color: 'Color', size: 'Size', tooltip: 'Tooltip', filter: 'Filter' };

function formatRange(range?: RangeRef): string {
  if (!range) return '';
  const column = (value: number) => {
    let label = '';
    for (let current = value + 1; current > 0; current = Math.floor((current - 1) / 26)) label = String.fromCharCode(65 + ((current - 1) % 26)) + label;
    return label;
  };
  return `${column(range.startColumn)}${range.startRow + 1}:${column(range.endColumn)}${range.endRow + 1}`;
}

function cloneBindings(payload: DataChartDrawingPayload): DataChartDrawingPayload['bindings'] {
  return Object.fromEntries(bindingAreas.map((area) => [area, structuredClone(payload.bindings[area] ?? [])])) as DataChartDrawingPayload['bindings'];
}

export function DataChartPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], sheet, tables, onCommand }: DataChartPanelProps) {
  const entries = drawings
    .filter((drawing) => drawing.kind === 'data-chart')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'data-chart' }> } => entry.payload?.kind === 'data-chart');
  const selected = entries.find(({ drawing }) => selectedDrawingIds.includes(drawing.id)) ?? entries[0];
  const payload = selected?.payload;
  const [title, setTitle] = useState('');
  const [plotType, setPlotType] = useState<DataChartPlotType>('column');
  const [sourceKind, setSourceKind] = useState<'table' | 'report-sheet'>('table');
  const [tableId, setTableId] = useState('');
  const [rangeInput, setRangeInput] = useState('');

  useEffect(() => {
    if (!payload) return;
    setTitle(payload.inspector.title ?? '');
    setPlotType(payload.plotType);
    setSourceKind(payload.source.kind);
    setTableId(payload.source.kind === 'table' ? payload.source.tableId : '');
    setRangeInput(payload.source.kind === 'report-sheet' ? formatRange(payload.source.range) : '');
  }, [payload?.source.kind, payload?.source.kind === 'table' ? payload.source.tableId : payload?.source.kind === 'report-sheet' ? formatRange(payload.source.range) : '', payload?.plotType, payload?.inspector.title]);

  const sourceRange = useMemo(() => {
    if (sourceKind !== 'report-sheet') return undefined;
    const parsed = parseRangeInput(rangeInput, sheetId);
    return parsed ? { sheetId, ...parsed } : undefined;
  }, [rangeInput, sheetId, sourceKind]);

  const table = tables.find((entry) => entry.id === (payload?.source.kind === 'table' ? payload.source.tableId : tableId));
  const reportFields = sourceRange ? Array.from({ length: sourceRange.endColumn - sourceRange.startColumn + 1 }, (_, offset) => ({ id: `report-column-${offset}`, name: String(sheet.getCell(sourceRange.startRow, sourceRange.startColumn + offset)?.value ?? `Column ${offset + 1}`), ordinal: offset })) : [];
  const fields = table?.fields ?? reportFields;

  const update = (next: DataChartDrawingPayload): void => {
    if (!selected) return;
    onCommand({ commandId: 'dataChart.update', params: { sheetId, drawingId: selected.drawing.id, payload: next } });
  };

  const withInspector = (patch: Partial<DataChartDrawingPayload['inspector']>): DataChartDrawingPayload | undefined => payload ? { ...payload, inspector: { ...payload.inspector, ...structuredClone(patch) } } : undefined;
  const updateBinding = (area: DataChartBindingArea, index: number, patch: Partial<DataChartFieldBinding>): void => {
    if (!payload) return;
    const bindings = cloneBindings(payload);
    bindings[area] = bindings[area].map((entry, entryIndex) => entryIndex === index ? { ...entry, ...patch, area } : entry);
    update({ ...payload, bindings });
  };
  const addBinding = (area: DataChartBindingArea): void => {
    if (!payload || !fields[0]) return;
    const bindings = cloneBindings(payload);
    bindings[area] = [...bindings[area], { area, fieldId: fields[0].id, aggregate: area === 'values' ? 'sum' : 'none' }];
    update({ ...payload, bindings });
  };
  const removeBinding = (area: DataChartBindingArea, index: number): void => {
    if (!payload) return;
    const bindings = cloneBindings(payload);
    bindings[area] = bindings[area].filter((_entry, entryIndex) => entryIndex !== index);
    update({ ...payload, bindings });
  };
  const remapBindings = (nextFields: readonly { id: string }[]): DataChartDrawingPayload['bindings'] => {
    const bindings = cloneBindings(payload!);
    bindings.values = bindings.values.map((binding, index) => ({ ...binding, fieldId: nextFields[Math.min(index + 1, Math.max(0, nextFields.length - 1))]?.id ?? binding.fieldId }));
    bindings.category = bindings.category.map((binding) => ({ ...binding, fieldId: nextFields[0]?.id ?? binding.fieldId }));
    for (const area of bindingAreas.filter((entry) => entry !== 'values' && entry !== 'category')) bindings[area] = bindings[area].map((binding, index) => ({ ...binding, fieldId: nextFields[index % Math.max(1, nextFields.length)]?.id ?? binding.fieldId }));
    return bindings;
  };

  if (!payload || !selected) {
    return <Panel className="border-0 bg-transparent shadow-none"><PanelHeader><PanelTitle size="sm">Data Chart</PanelTitle></PanelHeader><PanelBody><Text size="sm" tone="muted">Insert a Data Chart from the Insert ribbon, then select it to open Inspector and Data Binding.</Text></PanelBody></Panel>;
  }

  return (
    <Panel className="border-0 bg-transparent shadow-none">
      <PanelHeader><PanelTitle size="sm">Data Chart · Inspector / Data Binding</PanelTitle></PanelHeader>
      <PanelBody><Stack gap="md">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Inspector · Title</Text><TextInput value={title} onChange={(event) => setTitle(event.target.value)} onBlur={() => { const next = withInspector({ title }); if (next) update(next); }} /></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Inspector · Chart type</Text><Select value={plotType} onChange={(event) => { const next = { ...payload, plotType: event.target.value as DataChartPlotType }; setPlotType(next.plotType); update(next); }} sizeVariant="sm">{plotTypes.map((type) => <option key={type} value={type}>{type}</option>)}</Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Inspector · Legend</Text><Select value={payload.inspector.legendPosition} onChange={(event) => { const next = withInspector({ legendPosition: event.target.value as DataChartDrawingPayload['inspector']['legendPosition'] }); if (next) update(next); }} sizeVariant="sm">{['none', 'top', 'bottom', 'left', 'right'].map((position) => <option key={position} value={position}>{position}</option>)}</Select></Box>
        <Button variant="secondary" size="sm" onClick={() => { const next = withInspector({ showDataLabels: !payload.inspector.showDataLabels }); if (next) update(next); }}>{payload.inspector.showDataLabels ? 'Hide data labels' : 'Show data labels'}</Button>
        <Button variant="secondary" size="sm" onClick={() => { const next = withInspector({ showHiddenData: !payload.inspector.showHiddenData }); if (next) update(next); }}>{payload.inspector.showHiddenData ? 'Hide worksheet hidden data' : 'Show worksheet hidden data'}</Button>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Data Binding · Source mode</Text><Select value={sourceKind} onChange={(event) => { const value = event.target.value as 'table' | 'report-sheet'; setSourceKind(value); if (value === 'table') { const nextTable = tables[0]; if (nextTable) update({ ...payload, source: { kind: 'table', tableId: nextTable.id }, bindings: remapBindings(nextTable.fields) }); } else if (sourceRange) update({ ...payload, source: { kind: 'report-sheet', range: sourceRange }, bindings: remapBindings(reportFields) }); }} sizeVariant="sm"><option value="table">DataManager Table Binding</option><option value="report-sheet">ReportSheet Cell Binding</option></Select></Box>
        {sourceKind === 'table' ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Bound table</Text><Select value={tableId} onChange={(event) => { const nextTable = tables.find((entry) => entry.id === event.target.value); setTableId(event.target.value); if (nextTable) update({ ...payload, source: { kind: 'table', tableId: nextTable.id }, bindings: remapBindings(nextTable.fields) }); }} sizeVariant="sm">{tables.map((entry) => <option key={entry.id} value={entry.id}>{entry.name}</option>)}</Select></Box> : <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">ReportSheet range</Text><TextInput value={rangeInput} onChange={(event) => setRangeInput(event.target.value)} onBlur={() => { if (sourceRange) update({ ...payload, source: { kind: 'report-sheet', range: sourceRange }, bindings: remapBindings(reportFields) }); }} placeholder="A1:F20" /></Box>}
        {bindingAreas.map((area) => <Box key={area} className="rounded border border-slate-200 p-2"><Text size="xs" weight="semibold" className="mb-2 text-slate-700">{areaLabels[area]}</Text><Stack gap="xs">{payload.bindings[area].map((binding, index) => <Box key={`${area}-${index}`} className="flex items-center gap-1"><Select value={binding.fieldId} onChange={(event) => updateBinding(area, index, { fieldId: event.target.value })} sizeVariant="sm">{fields.map((field) => <option key={field.id} value={field.id}>{field.name}</option>)}</Select>{area === 'values' ? <><Select value={binding.aggregate} onChange={(event) => updateBinding(area, index, { aggregate: event.target.value as DataChartFieldBinding['aggregate'] })} sizeVariant="sm"><option value="sum">Sum</option><option value="average">Average</option><option value="count">Count</option><option value="min">Min</option><option value="max">Max</option><option value="none">None</option></Select><Select value={binding.sort ?? 'none'} onChange={(event) => updateBinding(area, index, { sort: event.target.value === 'none' ? undefined : event.target.value as 'asc' | 'desc' })} sizeVariant="sm"><option value="none">No sort</option><option value="asc">Ascending</option><option value="desc">Descending</option></Select><TextInput value={binding.format ?? ''} placeholder="Format" onChange={(event) => updateBinding(area, index, { format: event.target.value || undefined })} /></> : null}<Button icon="trash" iconOnly variant="ghost" size="xs" onClick={() => removeBinding(area, index)} /></Box>)}<Button variant="ghost" size="xs" onClick={() => addBinding(area)}>Add field</Button></Stack></Box>)}
      </Stack></PanelBody>
    </Panel>
  );
}
