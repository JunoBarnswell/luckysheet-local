import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ChartDrawingPayload, DrawingObject, DrawingPayload, RangeRef } from '@react-sheets/core-model';
import { parseRangeInput } from '../../domain/range-input';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface ChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  defaultRange?: string;
  onInsertChart: (type: ChartDrawingPayload['chartType'], sourceRange: RangeRef, title: string, stacked: NonNullable<ChartDrawingPayload['stacked']>) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

const chartTypes: readonly ChartDrawingPayload['chartType'][] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'];

function parseA1Range(value: string, sheetId: string): RangeRef | undefined {
  const range = parseRangeInput(value, sheetId);
  return range ? { sheetId, ...range } : undefined;
}

function columnLabel(column: number): string {
  let result = '';
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  return result;
}

function formatRange(range: RangeRef | undefined): string {
  if (!range) return '';
  return `${columnLabel(range.startColumn)}${range.startRow + 1}:${columnLabel(range.endColumn)}${range.endRow + 1}`;
}

export function ChartPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], defaultRange, onInsertChart, onCommand, onClose }: ChartPanelProps) {
  const chartEntries = drawings
    .filter((drawing) => drawing.kind === 'chart')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'chart' }> } => entry.payload?.kind === 'chart');
  const selectedEntry = chartEntries.find(({ drawing }) => selectedDrawingIds.includes(drawing.id));
  const selectedPayload = selectedEntry?.payload;
  const [type, setType] = useState<ChartDrawingPayload['chartType']>('column');
  const [stacked, setStacked] = useState<NonNullable<ChartDrawingPayload['stacked']>>('none');
  const [title, setTitle] = useState('Sales Overview');
  const [rangeInput, setRangeInput] = useState(defaultRange ?? 'A1:C5');
  const [categoryInput, setCategoryInput] = useState('');
  const sourceRange = parseA1Range(rangeInput, sheetId);
  const selectedId = selectedPayload?.chartId;
  const selectedCategoryRange = useMemo(() => parseA1Range(categoryInput, sheetId), [categoryInput, sheetId]);

  useEffect(() => {
    if (!selectedPayload) return;
    setType(selectedPayload.chartType);
    setStacked(selectedPayload.stacked ?? 'none');
    setTitle(selectedPayload.elements.title ?? '');
    const source = selectedPayload.sourceRanges[0];
    setRangeInput(formatRange(source));
    const category = selectedPayload.categoryRange;
    setCategoryInput(formatRange(category));
  }, [selectedId, selectedPayload]);

  const updateElements = (elements: Partial<ChartDrawingPayload['elements']>) => {
    if (!selectedPayload) return;
    onCommand({ commandId: 'chart.setElements', params: { sheetId, chartId: selectedPayload.chartId, elements } });
  };
  const updateSeries = (series: NonNullable<ChartDrawingPayload['series']>, categoryRange = selectedCategoryRange) => {
    if (!selectedPayload) return;
    onCommand({ commandId: 'chart.setSeries', params: { sheetId, chartId: selectedPayload.chartId, sourceRanges: selectedPayload.sourceRanges, series, categoryRange } });
  };
  const handleCreate = () => {
    if (!sourceRange) return;
    onInsertChart(type, sourceRange, title, stacked);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Chart Design & Format</PanelTitle></PanelHeader>
      <PanelBody className="p-4"><Stack gap="md">
        {selectedPayload ? <>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Chart Title</Text><TextInput value={title} onChange={(event) => setTitle(event.target.value)} onBlur={() => updateElements({ title })} /></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Chart Type</Text><Select value={selectedPayload.chartType} onChange={(event) => onCommand({ commandId: 'chart.setType', params: { sheetId, chartId: selectedPayload.chartId, chartType: event.target.value } })} sizeVariant="sm">{chartTypes.map((entry) => <option key={entry} value={entry}>{entry[0]!.toUpperCase() + entry.slice(1)}</option>)}</Select></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Legend</Text><Select value={selectedPayload.elements.legend?.visible ? selectedPayload.elements.legend.position : 'none'} onChange={(event) => updateElements({ legend: event.target.value === 'none' ? { visible: false, position: 'bottom' } : { visible: true, position: event.target.value as 'top' | 'bottom' | 'left' | 'right' } })} sizeVariant="sm">{['none', 'top', 'bottom', 'left', 'right'].map((entry) => <option key={entry} value={entry}>{entry}</option>)}</Select></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Data Labels</Text><Button variant="secondary" size="sm" onClick={() => updateElements({ dataLabels: { ...(selectedPayload.elements.dataLabels ?? { visible: false }), visible: !selectedPayload.elements.dataLabels?.visible } })}>{selectedPayload.elements.dataLabels?.visible ? 'Hide labels' : 'Show labels'}</Button></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Hidden/Filtered Data</Text><Select value={selectedPayload.elements.hiddenData} onChange={(event) => updateElements({ hiddenData: event.target.value as ChartDrawingPayload['elements']['hiddenData'] })} sizeVariant="sm"><option value="show">Show all data</option><option value="hideRows">Hide hidden rows</option><option value="hideColumns">Hide hidden columns</option></Select></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Value Axis</Text><Stack gap="xs"><TextInput value={selectedPayload.elements.valueAxis?.title ?? ''} placeholder="Axis title" onChange={(event) => updateElements({ valueAxis: { ...(selectedPayload.elements.valueAxis ?? { id: 'value', position: 'left' }), title: event.target.value } })} /><Button variant="secondary" size="sm" onClick={() => updateElements({ valueAxis: { ...(selectedPayload.elements.valueAxis ?? { id: 'value', position: 'left' }), majorGridlines: { visible: selectedPayload.elements.valueAxis?.majorGridlines?.visible === false } } })}>{selectedPayload.elements.valueAxis?.majorGridlines?.visible === false ? 'Show gridlines' : 'Hide gridlines'}</Button></Stack></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Select Data / Series</Text><Stack gap="xs">{(selectedPayload.series ?? []).map((series, index, all) => <Box key={`${series.name}-${index}`} className="flex items-center gap-1 rounded border border-slate-200 p-1"><TextInput value={series.name} onChange={(event) => updateSeries(all.map((entry, entryIndex) => entryIndex === index ? { ...entry, name: event.target.value } : entry))} /><Select value={series.chartType ?? (selectedPayload.chartType === 'combo' ? 'column' : selectedPayload.chartType)} onChange={(event) => updateSeries(all.map((entry, entryIndex) => entryIndex === index ? { ...entry, chartType: event.target.value as Exclude<ChartDrawingPayload['chartType'], 'combo'> } : entry))} sizeVariant="sm"><option value="column">Column</option><option value="bar">Bar</option><option value="line">Line</option><option value="area">Area</option><option value="scatter">Scatter</option></Select><Select value={series.axis ?? 'primary'} onChange={(event) => updateSeries(all.map((entry, entryIndex) => entryIndex === index ? { ...entry, axis: event.target.value as 'primary' | 'secondary' } : entry))} sizeVariant="sm"><option value="primary">Primary</option><option value="secondary">Secondary</option></Select><Button icon="arrow-up" iconOnly variant="ghost" size="xs" disabled={index === 0} onClick={() => { const next = [...all]; [next[index - 1], next[index]] = [next[index]!, next[index - 1]!]; updateSeries(next); }} /><Button icon="arrow-down" iconOnly variant="ghost" size="xs" disabled={index === all.length - 1} onClick={() => { const next = [...all]; [next[index], next[index + 1]] = [next[index + 1]!, next[index]!]; updateSeries(next); }} /><Button icon="trash" iconOnly variant="ghost" size="xs" onClick={() => updateSeries(all.filter((_entry, entryIndex) => entryIndex !== index))} /></Box>)}<Button variant="secondary" size="sm" onClick={() => { const range = selectedPayload.sourceRanges[0]; if (range) updateSeries([...(selectedPayload.series ?? []), { name: `Series ${(selectedPayload.series?.length ?? 0) + 1}`, range: structuredClone(range), chartType: selectedPayload.chartType === 'combo' ? 'column' : selectedPayload.chartType }]); }}>Add series</Button></Stack></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Category Range</Text><TextInput value={categoryInput} onChange={(event) => setCategoryInput(event.target.value)} onBlur={() => updateSeries(selectedPayload.series ?? [], selectedCategoryRange)} placeholder="A2:A5" /></Box>
        </> : <>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Chart Title</Text><TextInput value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Monthly Revenue" /></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Chart Type</Text><Select value={type} onChange={(event) => setType(event.target.value as ChartDrawingPayload['chartType'])} sizeVariant="sm">{chartTypes.map((entry) => <option key={entry} value={entry}>{entry[0]!.toUpperCase() + entry.slice(1)}</option>)}</Select></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Stacking</Text><Select value={stacked} onChange={(event) => setStacked(event.target.value as NonNullable<ChartDrawingPayload['stacked']>)} sizeVariant="sm"><option value="none">Grouped</option><option value="stacked">Stacked</option><option value="percent">100% Stacked</option></Select></Box>
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Data Source Range</Text><TextInput value={rangeInput} onChange={(event) => setRangeInput(event.target.value)} placeholder="e.g. A1:F6" /></Box>
          <Button variant="primary" size="sm" icon="plus" disabled={!sourceRange} onClick={handleCreate}>Insert Chart to Canvas</Button>
        </>}
        {chartEntries.length > 0 ? <Box className="mt-4 border-t border-slate-200 pt-3"><Text size="xs" weight="semibold" className="mb-2 text-slate-700">Worksheet Charts ({chartEntries.length})</Text><Stack gap="xs">{chartEntries.map(({ drawing, payload }) => <Box key={drawing.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"><Stack gap="none"><Text size="sm" weight="medium" className="text-slate-800">{payload.elements.title || payload.chartType}</Text><Text size="xs" tone="subtle">{payload.chartType.toUpperCase()} Chart</Text></Stack><Button variant="ghost" size="xs" icon="trash" iconOnly onClick={() => onCommand({ commandId: 'chart.remove', params: { sheetId, chartId: payload.chartId } })} className="text-rose-600 hover:bg-rose-50" /></Box>)}</Stack></Box> : null}
      </Stack></PanelBody>
      {onClose ? <PanelFooter className="border-t border-slate-200 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button></PanelFooter> : null}
    </Panel>
  );
}
