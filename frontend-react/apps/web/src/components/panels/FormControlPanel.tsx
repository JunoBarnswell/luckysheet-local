import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { DrawingObject, DrawingPayload, FormControlDrawingPayload, RangeRef } from '@react-sheets/core-model';

export interface FormControlPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  onCommand: (descriptor: CommandDescriptor) => void;
}

function numeric(value: string, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function columnNumber(label: string): number {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function rangeText(range: RangeRef): string {
  const address = (row: number, column: number): string => {
    let label = '';
    let current = column + 1;
    while (current > 0) {
      const modulo = (current - 1) % 26;
      label = String.fromCharCode(65 + modulo) + label;
      current = Math.floor((current - 1) / 26);
    }
    return `${label}${row + 1}`;
  };
  const start = address(range.startRow, range.startColumn);
  const end = address(range.endRow, range.endColumn);
  return start === end ? start : `${start}:${end}`;
}

function parseRangeText(value: string, sheetId: string): RangeRef | undefined {
  const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(value.trim());
  if (!match) return undefined;
  const startColumn = columnNumber(match[1]!);
  const startRow = Number(match[2]) - 1;
  const endColumn = columnNumber(match[3] ?? match[1]!);
  const endRow = Number(match[4] ?? match[2]) - 1;
  if (![startColumn, startRow, endColumn, endRow].every((entry) => Number.isSafeInteger(entry) && entry >= 0)) return undefined;
  return { sheetId, startRow: Math.min(startRow, endRow), endRow: Math.max(startRow, endRow), startColumn: Math.min(startColumn, endColumn), endColumn: Math.max(startColumn, endColumn) };
}

export function FormControlPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], onCommand }: FormControlPanelProps) {
  const selected = useMemo(() => {
    const id = selectedDrawingIds[0];
    const drawing = id ? drawings.find((entry) => entry.id === id && entry.kind === 'form-control') : undefined;
    const payload = drawing ? drawingPayloads.get(drawing.payloadId) : undefined;
    return drawing && payload?.kind === 'form-control' ? { drawing, payload } : undefined;
  }, [drawingPayloads, drawings, selectedDrawingIds]);
  const [text, setText] = useState('');
  const [enabled, setEnabled] = useState('true');
  const [minValue, setMinValue] = useState('0');
  const [maxValue, setMaxValue] = useState('100');
  const [step, setStep] = useState('1');
  const [pageChange, setPageChange] = useState('10');
  const [selectionType, setSelectionType] = useState<'single' | 'multiple'>('single');
  const [dropDownLines, setDropDownLines] = useState('8');
  const [inputRangeText, setInputRangeText] = useState('');

  useEffect(() => {
    const payload = selected?.payload;
    setText(payload?.text ?? '');
    setEnabled(String(payload?.enabled ?? true));
    if (payload?.controlType === 'spin-button' || payload?.controlType === 'scrollbar') {
      setMinValue(String(payload.minValue));
      setMaxValue(String(payload.maxValue));
      setStep(String(payload.step));
      if (payload.controlType === 'scrollbar') setPageChange(String(payload.pageChange));
    }
    if (payload?.controlType === 'list-box') setSelectionType(payload.selectionType);
    if (payload?.controlType === 'combo-box') setDropDownLines(String(payload.dropDownLines));
    if (payload?.controlType === 'list-box' || payload?.controlType === 'combo-box') setInputRangeText(rangeText(payload.inputRange));
  }, [selected?.drawing.id, selected?.payload]);

  if (!selected) {
    return <Panel className="h-full border-0 bg-transparent shadow-none"><PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Form Control Properties</PanelTitle></PanelHeader><PanelBody className="p-4"><Text size="sm" tone="muted">Select a form control to edit its typed properties.</Text></PanelBody></Panel>;
  }

  const apply = (patch: Partial<FormControlDrawingPayload>): void => {
    const payload = structuredClone(selected.payload) as FormControlDrawingPayload;
    Object.assign(payload, patch);
    onCommand({ commandId: 'formControl.update', params: { sheetId, drawingId: selected.drawing.id, payload } });
  };
  const applyNumeric = (field: 'minValue' | 'maxValue' | 'step' | 'pageChange', value: string): void => {
    const current = selected.payload;
    if (current.controlType !== 'spin-button' && current.controlType !== 'scrollbar') return;
    if (field === 'pageChange') {
      if (current.controlType !== 'scrollbar') return;
      apply({ pageChange: numeric(value, current.pageChange) });
      return;
    }
    const fallback = current[field];
    apply({ [field]: numeric(value, fallback) } as Partial<FormControlDrawingPayload>);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Form Control Properties</PanelTitle></PanelHeader>
      <PanelBody className="p-4"><Stack gap="md">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Type</Text><Text size="sm" weight="semibold">{selected.payload.controlType}</Text></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Label</Text><TextInput value={text} onChange={(event) => setText(event.target.value)} onBlur={() => apply({ text })} /></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Enabled</Text><Select value={enabled} onChange={(event) => { setEnabled(event.target.value); apply({ enabled: event.target.value === 'true' }); }} sizeVariant="sm"><option value="true">Enabled</option><option value="false">Disabled</option></Select></Box>
        {selected.payload.controlType === 'spin-button' || selected.payload.controlType === 'scrollbar' ? (
          <Box className="grid grid-cols-3 gap-2">
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Minimum</Text><TextInput value={minValue} onChange={(event) => setMinValue(event.target.value)} onBlur={() => applyNumeric('minValue', minValue)} /></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Maximum</Text><TextInput value={maxValue} onChange={(event) => setMaxValue(event.target.value)} onBlur={() => applyNumeric('maxValue', maxValue)} /></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Step</Text><TextInput value={step} onChange={(event) => setStep(event.target.value)} onBlur={() => applyNumeric('step', step)} /></Box>
          </Box>
        ) : null}
        {selected.payload.controlType === 'scrollbar' ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Page change</Text><TextInput value={pageChange} onChange={(event) => setPageChange(event.target.value)} onBlur={() => applyNumeric('pageChange', pageChange)} /></Box> : null}
        {selected.payload.controlType === 'list-box' ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Input range</Text><TextInput value={inputRangeText} onChange={(event) => setInputRangeText(event.target.value)} onBlur={() => { const range = parseRangeText(inputRangeText, sheetId); if (range) apply({ inputRange: range }); }} placeholder="A1:A5" /><Text size="xs" weight="medium" className="mb-1 mt-2 text-slate-700">Selection</Text><Select value={selectionType} onChange={(event) => { const next = event.target.value as 'single' | 'multiple'; setSelectionType(next); apply({ selectionType: next }); }} sizeVariant="sm"><option value="single">Single</option><option value="multiple">Multiple</option></Select></Box> : null}
        {selected.payload.controlType === 'combo-box' ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Input range</Text><TextInput value={inputRangeText} onChange={(event) => setInputRangeText(event.target.value)} onBlur={() => { const range = parseRangeText(inputRangeText, sheetId); if (range) apply({ inputRange: range }); }} placeholder="A1:A5" /><Text size="xs" weight="medium" className="mb-1 mt-2 text-slate-700">Drop-down lines</Text><TextInput value={dropDownLines} onChange={(event) => setDropDownLines(event.target.value)} onBlur={() => apply({ dropDownLines: Math.max(1, Math.min(100, Math.trunc(numeric(dropDownLines, selected.payload.controlType === 'combo-box' ? selected.payload.dropDownLines : 8)))) } as Partial<FormControlDrawingPayload>) } /></Box> : null}
        {selected.payload.controlType === 'button' ? <Text size="xs" tone="muted">Click event: {selected.payload.action.eventId}</Text> : null}
        <Button variant="secondary" size="sm" onClick={() => apply({ text })}>Apply properties</Button>
      </Stack></PanelBody>
    </Panel>
  );
}
