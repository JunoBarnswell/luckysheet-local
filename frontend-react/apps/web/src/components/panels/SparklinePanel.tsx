import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, CheckToggle, ColorPicker, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { RangeRef, SparklineGroup, SparklineModel } from '@react-sheets/core-model';
import { parseAddress } from '@react-sheets/spreadsheet-app';
import { parseRangeInput } from '../../domain/range-input';

export interface SparklinePanelProps {
  sheetId: string;
  activeCell: string;
  sparklines: readonly SparklineModel[];
  sparklineGroups: readonly SparklineGroup[];
  defaultRange?: string;
  onAddSparkline: (sparkline: SparklineModel) => void;
  onRemoveSparkline: (id: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

function columnName(column: number): string {
  let value = column + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function rangeName(range: RangeRef): string {
  const start = `${columnName(range.startColumn)}${range.startRow + 1}`;
  const end = `${columnName(range.endColumn)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function parseTargetCell(input: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(input.trim().toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const char of match[1]) column = column * 26 + char.charCodeAt(0) - 64;
  const row = Number(match[2]) - 1;
  if (!Number.isSafeInteger(row) || row < 0 || column < 1) return undefined;
  return { row, column: column - 1 };
}

function defaultSparkline(id: string, sheetId: string, sourceRange: RangeRef, location: { row: number; column: number }, type: SparklineModel['type'], color: string): SparklineModel {
  return {
    id,
    sheetId,
    anchor: location,
    sourceRange,
    type,
    color,
    negativeColor: '#ef4444',
    highlightMax: true,
    highlightMin: true,
    showAxis: false,
    showMarkers: false,
  };
}

export function SparklinePanel({
  sheetId,
  activeCell,
  sparklines,
  sparklineGroups,
  defaultRange,
  onAddSparkline,
  onRemoveSparkline,
  onCommand,
  onClose,
}: SparklinePanelProps) {
  const activeAddress = parseAddress(activeCell);
  const anchoredSparkline = sparklines.find((entry) => entry.anchor.row === activeAddress?.row && entry.anchor.column === activeAddress?.column);
  const [selectedIds, setSelectedIds] = useState<string[]>(anchoredSparkline ? [anchoredSparkline.id] : []);
  const selectedSparkline = selectedIds.length === 1 ? sparklines.find((entry) => entry.id === selectedIds[0]) : undefined;
  const selectedGroup = selectedSparkline?.groupId ? sparklineGroups.find((group) => group.id === selectedSparkline.groupId) : undefined;
  const [type, setType] = useState<SparklineModel['type']>(selectedSparkline?.type ?? 'line');
  const [sourceRange, setSourceRange] = useState(defaultRange ?? 'B2:E2');
  const [targetCell, setTargetCell] = useState('F2');
  const [color, setColor] = useState('#2563eb');
  const [negativeColor, setNegativeColor] = useState('#ef4444');
  const [highlightMax, setHighlightMax] = useState(true);
  const [highlightMin, setHighlightMin] = useState(true);
  const [highlightFirst, setHighlightFirst] = useState(false);
  const [highlightLast, setHighlightLast] = useState(false);
  const [highlightNegative, setHighlightNegative] = useState(false);
  const [showAxis, setShowAxis] = useState(false);
  const [showMarkers, setShowMarkers] = useState(false);
  const parsedSourceRange = parseRangeInput(sourceRange, sheetId);

  useEffect(() => {
    if (anchoredSparkline) setSelectedIds([anchoredSparkline.id]);
    else setSelectedIds((current) => current.filter((id) => sparklines.some((entry) => entry.id === id)));
  }, [anchoredSparkline?.id, sparklines]);

  useEffect(() => {
    if (!selectedSparkline) {
      setType('line');
      setSourceRange(defaultRange ?? 'B2:E2');
      setTargetCell('F2');
      setColor('#2563eb');
      setNegativeColor('#ef4444');
      setHighlightMax(true);
      setHighlightMin(true);
      setHighlightFirst(false);
      setHighlightLast(false);
      setHighlightNegative(false);
      setShowAxis(false);
      setShowMarkers(false);
      return;
    }
    setType(selectedSparkline.type);
    setSourceRange(rangeName(selectedSparkline.sourceRange));
    setTargetCell(`${columnName(selectedSparkline.anchor.column)}${selectedSparkline.anchor.row + 1}`);
    setColor(selectedSparkline.color);
    setNegativeColor(selectedSparkline.negativeColor ?? '#ef4444');
    setHighlightMax(selectedSparkline.highlightMax === true);
    setHighlightMin(selectedSparkline.highlightMin === true);
    setHighlightFirst(selectedSparkline.highlightFirst === true);
    setHighlightLast(selectedSparkline.highlightLast === true);
    setHighlightNegative(selectedSparkline.highlightNegative === true);
    setShowAxis(selectedSparkline.showAxis === true);
    setShowMarkers(selectedSparkline.showMarkers === true);
  }, [defaultRange, selectedSparkline?.id]);

  const selectedRangeIds = useMemo(() => new Set(selectedIds), [selectedIds]);
  const toggleSelected = (id: string): void => setSelectedIds((current) => current.includes(id) ? current.filter((entry) => entry !== id) : [...current, id]);

  const handleCommit = (): void => {
    const parsed = parsedSourceRange;
    const location = parseTargetCell(targetCell);
    if (!parsed || !location) return;
    const source: RangeRef = { sheetId, ...parsed };
    const patch: Partial<SparklineModel> = {
      sourceRange: source,
      anchor: location,
      color,
      negativeColor,
      highlightMax,
      highlightMin,
      highlightFirst,
      highlightLast,
      highlightNegative,
    };
    if (!selectedGroup) {
      patch.type = type;
      patch.showAxis = showAxis;
      patch.showMarkers = showMarkers;
    }
    if (selectedSparkline) {
      onCommand({ commandId: 'sparkline.update', params: { sheetId, sparklineId: selectedSparkline.id, patch } });
      return;
    }
    onAddSparkline({
      ...defaultSparkline(`spark-${Date.now().toString(36)}`, sheetId, source, location, type, color),
      negativeColor,
      highlightMax,
      highlightMin,
      highlightFirst,
      highlightLast,
      highlightNegative,
      showAxis,
      showMarkers,
    });
  };

  const createGroup = (): void => {
    if (selectedIds.length < 2) return;
    const first = sparklines.find((entry) => entry.id === selectedIds[0]);
    if (!first) return;
    onCommand({ commandId: 'sparkline.group.create', params: { sheetId, group: { id: `sparkline-group-${Date.now().toString(36)}`, sheetId, type: first.type, sparklineIds: [...selectedIds], showAxis: first.showAxis, showMarkers: first.showMarkers } } });
  };

  const updateGroup = (patch: Partial<SparklineGroup>): void => {
    if (!selectedGroup) return;
    onCommand({ commandId: 'sparkline.group.update', params: { sheetId, groupId: selectedGroup.id, patch } });
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Sparkline Design</PanelTitle></PanelHeader>
      <PanelBody className="p-4"><Stack gap="md">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Sparkline</Text><Select value={type} disabled={Boolean(selectedGroup)} onChange={(event) => setType(event.target.value as SparklineModel['type'])} sizeVariant="sm"><option value="line">Line</option><option value="column">Column</option><option value="win-loss">Win / Loss</option></Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Data source range</Text><TextInput value={sourceRange} onChange={(event) => setSourceRange(event.target.value)} placeholder="e.g. B2:E2" /></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Target cell</Text><TextInput value={targetCell} onChange={(event) => setTargetCell(event.target.value)} placeholder="e.g. F2" /></Box>
        <Box className="grid grid-cols-2 gap-2"><Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Series color</Text><ColorPicker color={color} onChange={setColor} /><TextInput value={color} onChange={(event) => setColor(event.target.value)} className="mt-1 h-8 text-xs font-mono" /></Box><Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Negative color</Text><ColorPicker color={negativeColor} onChange={setNegativeColor} /><TextInput value={negativeColor} onChange={(event) => setNegativeColor(event.target.value)} className="mt-1 h-8 text-xs font-mono" /></Box></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Points and axis</Text><Inline gap="xs" className="flex-wrap"><CheckToggle label="High" checked={highlightMax} onChange={(event) => setHighlightMax(event.currentTarget.checked)} /><CheckToggle label="Low" checked={highlightMin} onChange={(event) => setHighlightMin(event.currentTarget.checked)} /><CheckToggle label="First" checked={highlightFirst} onChange={(event) => setHighlightFirst(event.currentTarget.checked)} /><CheckToggle label="Last" checked={highlightLast} onChange={(event) => setHighlightLast(event.currentTarget.checked)} /><CheckToggle label="Negative" checked={highlightNegative} onChange={(event) => setHighlightNegative(event.currentTarget.checked)} /><CheckToggle label="Markers" checked={showMarkers} disabled={Boolean(selectedGroup)} onChange={(event) => setShowMarkers(event.currentTarget.checked)} /><CheckToggle label="Axis" checked={showAxis} disabled={Boolean(selectedGroup)} onChange={(event) => setShowAxis(event.currentTarget.checked)} /></Inline></Box>
        {selectedGroup ? <Box className="rounded border border-slate-200 bg-white p-2"><Text size="xs" weight="medium" className="mb-1 text-slate-700">Group settings · {selectedGroup.id}</Text><Stack gap="xs"><Select value={selectedGroup.type} onChange={(event) => updateGroup({ type: event.target.value as SparklineModel['type'] })} sizeVariant="sm"><option value="line">Line</option><option value="column">Column</option><option value="win-loss">Win / Loss</option></Select><CheckToggle label="Shared markers" checked={selectedGroup.showMarkers === true} onChange={(event) => updateGroup({ showMarkers: event.currentTarget.checked })} /><CheckToggle label="Shared axis" checked={selectedGroup.showAxis === true} onChange={(event) => updateGroup({ showAxis: event.currentTarget.checked })} /><Button variant="secondary" size="sm" onClick={() => onCommand({ commandId: 'sparkline.group.remove', params: { sheetId, groupId: selectedGroup.id } })}>Ungroup selected</Button></Stack></Box> : null}
        <Inline gap="xs"><Button variant="primary" size="sm" icon="sparkline" disabled={!parsedSourceRange || !parseTargetCell(targetCell)} onClick={handleCommit}>{selectedSparkline ? 'Update Sparkline' : 'Insert Sparkline'}</Button><Button variant="secondary" size="sm" disabled={selectedIds.length < 2} onClick={createGroup}>Group selected</Button></Inline>
        {sparklines.length > 0 ? <Box className="border-t border-slate-200 pt-3"><Text size="xs" weight="semibold" className="mb-2 text-slate-700">Cell Sparklines ({sparklines.length})</Text><Stack gap="xs">{sparklines.map((sparkline) => <Box key={sparkline.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"><CheckToggle aria-label={`Select ${sparkline.id}`} checked={selectedRangeIds.has(sparkline.id)} onChange={() => toggleSelected(sparkline.id)} /><Button variant="ghost" size="xs" className="min-w-0 flex-1 justify-start px-2" onClick={() => setSelectedIds([sparkline.id])}>{sparkline.type.toUpperCase()} · {columnName(sparkline.anchor.column)}{sparkline.anchor.row + 1}{sparkline.groupId ? ` · ${sparkline.groupId}` : ''}</Button><Button variant="ghost" size="xs" icon="trash" iconOnly aria-label={`Delete ${sparkline.id}`} onClick={() => onRemoveSparkline(sparkline.id)} className="text-rose-600 hover:bg-rose-50" /></Box>)}</Stack></Box> : null}
      </Stack></PanelBody>
      {onClose ? <Box className="border-t border-slate-200 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button></Box> : null}
    </Panel>
  );
}
