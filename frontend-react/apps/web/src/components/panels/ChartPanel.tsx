import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { CHART_SUBTYPES_BY_TYPE, defaultChartSubtype, type ChartAxisModel, type ChartDrawingPayload, type ChartSeriesModel, type DrawingObject, type DrawingPayload, type RangeRef } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { ChartElementSelection } from '@react-sheets/spreadsheet-app';
import { parseRangeInput } from '../../domain/range-input';

export interface ChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  selectedChartElement?: ChartElementSelection | null;
  defaultRange?: string;
  onInsertChart: (type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype'], sourceRange: RangeRef, title: string, stacked: NonNullable<ChartDrawingPayload['stacked']>) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

const chartTypes: readonly ChartDrawingPayload['chartType'][] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble', 'treemap', 'sunburst', 'histogram', 'pareto', 'box-whisker', 'waterfall', 'funnel', 'stock', 'surface', 'radar', 'map', 'combo'];
const seriesTypes: readonly Exclude<ChartDrawingPayload['chartType'], 'combo'>[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble', 'treemap', 'sunburst', 'histogram', 'pareto', 'box-whisker', 'waterfall', 'funnel', 'stock', 'surface', 'radar', 'map'];

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

function chartLabel(type: ChartDrawingPayload['chartType']): string {
  return type === 'box-whisker' ? 'Box & Whisker' : type[0]!.toUpperCase() + type.slice(1);
}

function seriesId(entry: ChartSeriesModel, index: number): string { return entry.id ?? `series:${index + 1}`; }

function Field({ label, children }: { label: string; children: React.ReactNode }): React.ReactNode {
  return <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">{label}</Text>{children}</Box>;
}

function axisWith(axis: ChartAxisModel | undefined, position: ChartAxisModel['position']): ChartAxisModel {
  return { id: axis?.id ?? position, position, axisType: axis?.axisType ?? 'value', ...(axis ?? {}) };
}

function fillValue(fill: NonNullable<ChartDrawingPayload['elements']['chartArea']>['fill'] | undefined): string {
  return typeof fill === 'string' ? fill : fill?.color ?? '';
}

export function ChartPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], selectedChartElement = null, defaultRange, onInsertChart, onCommand, onClose }: ChartPanelProps) {
  const chartEntries = drawings
    .filter((drawing) => drawing.kind === 'chart')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'chart' }> } => entry.payload?.kind === 'chart');
  const selectedEntry = chartEntries.find(({ drawing }) => selectedDrawingIds.includes(drawing.id));
  const selectedPayload = selectedEntry?.payload;
  const [type, setType] = useState<ChartDrawingPayload['chartType']>('column');
  const [subtype, setSubtype] = useState<ChartDrawingPayload['subtype']>('clustered');
  const [stacked, setStacked] = useState<NonNullable<ChartDrawingPayload['stacked']>>('none');
  const [title, setTitle] = useState('Sales Overview');
  const [rangeInput, setRangeInput] = useState(defaultRange ?? 'A1:C5');
  const [categoryInput, setCategoryInput] = useState('');
  const sourceRange = parseA1Range(rangeInput, sheetId);
  const selectedId = selectedPayload?.chartId;
  const selectedCategoryRange = useMemo(() => parseA1Range(categoryInput, sheetId), [categoryInput, sheetId]);
  const primaryWorksheetRange = selectedPayload?.source.kind === 'worksheet-ranges' ? selectedPayload.source.ranges[0] : undefined;

  useEffect(() => {
    if (!selectedPayload) return;
    setType(selectedPayload.chartType);
    setSubtype(selectedPayload.subtype);
    setStacked(selectedPayload.stacked ?? 'none');
    setTitle(selectedPayload.elements.title ?? '');
    const source = selectedPayload.source.kind === 'worksheet-ranges' ? selectedPayload.source.ranges[0] : selectedPayload.source.kind === 'report-range' ? selectedPayload.source.range : undefined;
    setRangeInput(formatRange(source));
    setCategoryInput(formatRange(selectedPayload.categoryRange));
  }, [selectedId, selectedPayload]);

  const updateElements = (elements: Partial<ChartDrawingPayload['elements']>) => {
    if (!selectedPayload) return;
    onCommand({ commandId: 'chart.setElements', params: { sheetId, chartId: selectedPayload.chartId, elements } });
  };
  const updateSeries = (series: NonNullable<ChartDrawingPayload['series']>, categoryRange = selectedCategoryRange) => {
    if (!selectedPayload) return;
    onCommand({ commandId: 'chart.setSeries', params: { sheetId, chartId: selectedPayload.chartId, source: selectedPayload.source, series: series.map((entry, index) => ({ ...entry, id: seriesId(entry, index) })), categoryRange } });
  };
  const updateAxis = (key: 'valueAxis' | 'categoryAxis' | 'secondaryValueAxis', patch: Partial<ChartAxisModel>) => {
    if (!selectedPayload) return;
    const current = selectedPayload.elements[key];
    updateElements({ [key]: { ...axisWith(current, key === 'categoryAxis' ? 'bottom' : key === 'secondaryValueAxis' ? 'right' : 'left'), ...patch } });
  };
  const handleCreate = () => { if (sourceRange) onInsertChart(type, subtype, sourceRange, title, stacked); };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Chart Design & Format</PanelTitle></PanelHeader>
      <PanelBody className="p-4"><Stack gap="md">
        {selectedChartElement ? <Box className="rounded border border-emerald-200 bg-emerald-50 px-3 py-2"><Text size="xs" weight="semibold" className="text-emerald-800">Selected element: {selectedChartElement.kind}{selectedChartElement.kind === 'point' || selectedChartElement.kind === 'series' ? ` · ${selectedChartElement.seriesId}` : ''}</Text></Box> : null}
        {selectedPayload ? <>
          <Field label="Chart Title"><TextInput value={title} onChange={(event) => setTitle(event.target.value)} onBlur={() => updateElements({ title })} /></Field>
          <Field label="Chart Type"><Select value={selectedPayload.chartType} onChange={(event) => { const chartType = event.target.value as ChartDrawingPayload['chartType']; onCommand({ commandId: 'chart.setType', params: { sheetId, chartId: selectedPayload.chartId, chartType, subtype: defaultChartSubtype(chartType) } }); }} sizeVariant="sm">{chartTypes.map((entry) => <option key={entry} value={entry}>{chartLabel(entry)}</option>)}</Select></Field>
          <Field label="Chart Subtype"><Select value={selectedPayload.subtype} onChange={(event) => onCommand({ commandId: 'chart.setType', params: { sheetId, chartId: selectedPayload.chartId, chartType: selectedPayload.chartType, subtype: event.target.value } })} sizeVariant="sm">{CHART_SUBTYPES_BY_TYPE[selectedPayload.chartType].map((entry) => <option key={entry} value={entry}>{entry}</option>)}</Select></Field>
          <Field label="Legend"><Select value={selectedPayload.elements.legend?.visible ? selectedPayload.elements.legend.position : 'none'} onChange={(event) => updateElements({ legend: event.target.value === 'none' ? { visible: false, position: 'bottom' } : { visible: true, position: event.target.value as NonNullable<ChartDrawingPayload['elements']['legend']>['position'] } })} sizeVariant="sm"><option value="none">None</option><option value="top">Top</option><option value="top-right">Top Right</option><option value="bottom">Bottom</option><option value="left">Left</option><option value="right">Right</option></Select></Field>
          <Field label="Data Labels"><Stack gap="xs"><CheckToggle label="Show labels" checked={selectedPayload.elements.dataLabels?.visible === true} onChange={(event) => updateElements({ dataLabels: { ...(selectedPayload.elements.dataLabels ?? { visible: false }), visible: event.currentTarget.checked } })} /><CheckToggle label="Value" checked={selectedPayload.elements.dataLabels?.showValue !== false} onChange={(event) => updateElements({ dataLabels: { ...(selectedPayload.elements.dataLabels ?? { visible: true }), showValue: event.currentTarget.checked } })} /><CheckToggle label="Category / Series" checked={selectedPayload.elements.dataLabels?.showCategoryName === true || selectedPayload.elements.dataLabels?.showSeriesName === true} onChange={(event) => updateElements({ dataLabels: { ...(selectedPayload.elements.dataLabels ?? { visible: true }), showCategoryName: event.currentTarget.checked, showSeriesName: event.currentTarget.checked } })} /></Stack></Field>
          <Field label="Data Table"><CheckToggle label="Show chart data table" checked={selectedPayload.elements.dataTable?.visible === true} onChange={(event) => onCommand({ commandId: 'chart.setDataTable', params: { sheetId, chartId: selectedPayload.chartId, dataTable: { ...(selectedPayload.elements.dataTable ?? { visible: false }), visible: event.currentTarget.checked } } })} /></Field>
          <Field label="Hidden / Empty Data"><Stack gap="xs"><Select value={selectedPayload.elements.hiddenData} onChange={(event) => updateElements({ hiddenData: event.target.value as ChartDrawingPayload['elements']['hiddenData'] })} sizeVariant="sm"><option value="show">Show all data</option><option value="hideRows">Do not plot hidden rows</option><option value="hideColumns">Do not plot hidden columns</option></Select><Select value={selectedPayload.elements.emptyCells ?? 'gap'} onChange={(event) => updateElements({ emptyCells: event.target.value as ChartDrawingPayload['elements']['emptyCells'] })} sizeVariant="sm"><option value="gap">Empty cells: gap</option><option value="zero">Empty cells: zero</option><option value="connect">Empty cells: connect</option></Select></Stack></Field>
          <Field label="Chart Area Format"><Stack gap="xs"><TextInput value={fillValue(selectedPayload.elements.chartArea?.fill)} placeholder="#ffffff" onChange={(event) => updateElements({ chartArea: { ...(selectedPayload.elements.chartArea ?? {}), fill: event.target.value } })} /><TextInput value={selectedPayload.elements.chartArea?.border ?? ''} placeholder="#cbd5e1" onChange={(event) => updateElements({ chartArea: { ...(selectedPayload.elements.chartArea ?? {}), border: event.target.value } })} /></Stack></Field>
          <Field label="Value Axis"><Stack gap="xs"><TextInput value={selectedPayload.elements.valueAxis?.title ?? ''} placeholder="Axis title" onChange={(event) => updateAxis('valueAxis', { title: event.target.value })} /><Select value={selectedPayload.elements.valueAxis?.scale ?? 'linear'} onChange={(event) => updateAxis('valueAxis', { scale: event.target.value as ChartAxisModel['scale'] })} sizeVariant="sm"><option value="linear">Linear scale</option><option value="logarithmic">Logarithmic scale</option></Select><Inline gap="xs"><TextInput value={selectedPayload.elements.valueAxis?.minimum?.toString() ?? ''} placeholder="Automatic min" onChange={(event) => updateAxis('valueAxis', { minimum: event.target.value === '' ? undefined : Number(event.target.value), automaticMinimum: event.target.value === '' })} /><TextInput value={selectedPayload.elements.valueAxis?.maximum?.toString() ?? ''} placeholder="Automatic max" onChange={(event) => updateAxis('valueAxis', { maximum: event.target.value === '' ? undefined : Number(event.target.value), automaticMaximum: event.target.value === '' })} /></Inline><CheckToggle label="Major gridlines" checked={selectedPayload.elements.valueAxis?.majorGridlines?.visible !== false} onChange={(event) => updateAxis('valueAxis', { majorGridlines: { ...(selectedPayload.elements.valueAxis?.majorGridlines ?? {}), visible: event.currentTarget.checked } })} /></Stack></Field>
          <Field label="Select Data"><Stack gap="xs"><TextInput value={formatRange(primaryWorksheetRange)} placeholder="Chart data range" onChange={(event) => setRangeInput(event.target.value)} onBlur={() => primaryWorksheetRange && onCommand({ commandId: 'chart.selectData', params: { sheetId, chartId: selectedPayload.chartId, source: { kind: 'worksheet-ranges', ranges: [parseA1Range(rangeInput, sheetId) ?? primaryWorksheetRange] }, categoryRange: selectedCategoryRange, series: selectedPayload.series } })} /><TextInput value={categoryInput} onChange={(event) => setCategoryInput(event.target.value)} onBlur={() => updateSeries(selectedPayload.series ?? [], selectedCategoryRange)} placeholder="Horizontal axis labels: A2:A5" /><Button variant="secondary" size="sm" onClick={() => onCommand({ commandId: 'chart.selectData', params: { sheetId, chartId: selectedPayload.chartId, source: selectedPayload.source, series: selectedPayload.series, categoryRange: selectedCategoryRange, switchRowColumn: true } })}>Switch Row / Column</Button></Stack></Field>
          <Field label="Series"><Stack gap="xs">{(selectedPayload.series ?? []).map((entry, index, all) => { const id = seriesId(entry, index); return <Box key={id} className="rounded border border-slate-200 bg-white p-2"><Stack gap="xs"><TextInput value={entry.name} onChange={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, name: event.target.value } : candidate))} /><Inline gap="xs"><Select value={entry.chartType ?? (selectedPayload.chartType === 'combo' ? 'column' : selectedPayload.chartType)} onChange={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, chartType: event.target.value as Exclude<ChartDrawingPayload['chartType'], 'combo'> } : candidate))} sizeVariant="sm">{seriesTypes.map((entryType) => <option key={entryType} value={entryType}>{chartLabel(entryType)}</option>)}</Select><Select value={entry.axis ?? 'primary'} onChange={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, axis: event.target.value as 'primary' | 'secondary' } : candidate))} sizeVariant="sm"><option value="primary">Primary axis</option><option value="secondary">Secondary axis</option></Select></Inline>{selectedPayload.chartType === 'scatter' || selectedPayload.chartType === 'bubble' ? <Stack gap="xs"><TextInput value={formatRange(entry.xRange)} placeholder="X values range" onBlur={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, xRange: parseA1Range(event.target.value, sheetId) } : candidate))} /><TextInput value={formatRange(entry.yRange ?? entry.range)} placeholder="Y values range" onBlur={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, yRange: parseA1Range(event.target.value, sheetId) } : candidate))} /><TextInput value={formatRange(entry.sizeRange)} placeholder="Bubble size range" onBlur={(event) => updateSeries(all.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, sizeRange: parseA1Range(event.target.value, sheetId) } : candidate))} /></Stack> : null}<Inline gap="xs"><Button icon="arrow-up" iconOnly variant="ghost" size="xs" disabled={index === 0} onClick={() => onCommand({ commandId: 'chart.series.move', params: { sheetId, chartId: selectedPayload.chartId, seriesId: id, direction: 'up' } })} /><Button icon="arrow-down" iconOnly variant="ghost" size="xs" disabled={index === all.length - 1} onClick={() => onCommand({ commandId: 'chart.series.move', params: { sheetId, chartId: selectedPayload.chartId, seriesId: id, direction: 'down' } })} /><Button icon="trash" iconOnly variant="ghost" size="xs" disabled={all.length <= 1} onClick={() => onCommand({ commandId: 'chart.series.remove', params: { sheetId, chartId: selectedPayload.chartId, seriesId: id } })} /></Inline></Stack></Box>; })}{primaryWorksheetRange ? <Button variant="secondary" size="sm" onClick={() => onCommand({ commandId: 'chart.series.add', params: { sheetId, chartId: selectedPayload.chartId, series: { id: `series:${(selectedPayload.series?.length ?? 0) + 1}`, name: `Series ${(selectedPayload.series?.length ?? 0) + 1}`, range: structuredClone(primaryWorksheetRange), chartType: selectedPayload.chartType === 'combo' ? 'column' : selectedPayload.chartType } } })}>Add series</Button> : null}</Stack></Field>
          {(selectedPayload.series ?? []).length > 0 ? <Field label="Analysis"><Stack gap="xs"><Button variant="secondary" size="sm" onClick={() => { const first = selectedPayload.series![0]!; onCommand({ commandId: 'chart.setTrendlines', params: { sheetId, chartId: selectedPayload.chartId, seriesId: seriesId(first, 0), trendlines: [...(first.trendlines ?? []), { type: 'linear', displayEquation: true, displayRSquared: true, color: first.color ?? '#2563eb', width: 1.5 }] } }); }}>Add linear trendline</Button><Button variant="secondary" size="sm" onClick={() => { const first = selectedPayload.series![0]!; onCommand({ commandId: 'chart.setErrorBars', params: { sheetId, chartId: selectedPayload.chartId, seriesId: seriesId(first, 0), errorBars: { type: 'standard-error', direction: 'vertical', endStyle: 'cap', color: first.color ?? '#2563eb', width: 1 } } }); }}>Add standard error bars</Button></Stack></Field> : null}
          <Box><Text size="xs" tone="muted">Source: {selectedPayload.source.kind} · {selectedPayload.nativeIdentity?.status === 'preserved-native' ? 'preserved-native' : 'canonical owned'}</Text></Box>
        </> : <>
          <Field label="Chart Title"><TextInput value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Monthly Revenue" /></Field>
          <Field label="Chart Type"><Select value={type} onChange={(event) => { const next = event.target.value as ChartDrawingPayload['chartType']; setType(next); setSubtype(defaultChartSubtype(next)); }} sizeVariant="sm">{chartTypes.map((entry) => <option key={entry} value={entry}>{chartLabel(entry)}</option>)}</Select></Field>
          <Field label="Chart Subtype"><Select value={subtype} onChange={(event) => setSubtype(event.target.value as ChartDrawingPayload['subtype'])} sizeVariant="sm">{CHART_SUBTYPES_BY_TYPE[type].map((entry) => <option key={entry} value={entry}>{entry}</option>)}</Select></Field>
          <Field label="Stacking"><Select value={stacked} onChange={(event) => setStacked(event.target.value as NonNullable<ChartDrawingPayload['stacked']>)} sizeVariant="sm"><option value="none">Grouped</option><option value="stacked">Stacked</option><option value="percent">100% Stacked</option></Select></Field>
          <Field label="Data Source Range"><TextInput value={rangeInput} onChange={(event) => setRangeInput(event.target.value)} placeholder="e.g. A1:F6" /></Field>
          <Button variant="primary" size="sm" icon="plus" disabled={!sourceRange} onClick={handleCreate}>Insert Chart to Canvas</Button>
        </>}
        {chartEntries.length > 0 ? <Box className="mt-4 border-t border-slate-200 pt-3"><Text size="xs" weight="semibold" className="mb-2 text-slate-700">Worksheet Charts ({chartEntries.length})</Text><Stack gap="xs">{chartEntries.map(({ drawing, payload }) => <Box key={drawing.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"><Stack gap="none"><Text size="sm" weight="medium" className="text-slate-800">{payload.elements.title || chartLabel(payload.chartType)}</Text><Text size="xs" tone="subtle">{payload.chartType.toUpperCase()} · {payload.subtype}</Text></Stack><Button variant="ghost" size="xs" icon="trash" iconOnly aria-label={`Delete ${payload.chartId}`} onClick={() => onCommand({ commandId: 'chart.remove', params: { sheetId, chartId: payload.chartId } })} className="text-rose-600 hover:bg-rose-50" /></Box>)}</Stack></Box> : null}
      </Stack></PanelBody>
      {onClose ? <PanelFooter className="border-t border-slate-200 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button></PanelFooter> : null}
    </Panel>
  );
}
