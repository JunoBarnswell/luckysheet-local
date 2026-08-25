import { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { DrawingObject, DrawingPayload, PivotControlStyle, PivotSlicerDrawingPayload, PivotSlicerSettings } from '@react-sheets/core-model';

export interface SlicerEditorPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  onCommand: (descriptor: CommandDescriptor) => void;
}

const STYLE_GALLERY: Record<string, PivotControlStyle> = {
  Light: { theme: 'light', fill: '#ffffff', border: '#94a3b8', textColor: '#0f172a', accentColor: '#e2e8f0', selectedFill: '#dbeafe', fontSize: 12 },
  Accent: { theme: 'accent', fill: '#eff6ff', border: '#2563eb', textColor: '#0f172a', accentColor: '#2563eb', selectedFill: '#bfdbfe', fontSize: 12 },
  Dark: { theme: 'dark', fill: '#1e293b', border: '#475569', textColor: '#f8fafc', accentColor: '#334155', selectedFill: '#475569', fontSize: 12 },
};

function selectedSlicer(drawings: readonly DrawingObject[], payloads: ReadonlyMap<string, DrawingPayload>, ids: readonly string[]): { drawing: DrawingObject; payload: PivotSlicerDrawingPayload } | undefined {
  const drawing = ids.map((id) => drawings.find((entry) => entry.id === id)).find((entry): entry is DrawingObject => entry?.kind === 'slicer');
  const payload = drawing ? payloads.get(drawing.payloadId) : undefined;
  return drawing && payload?.kind === 'slicer' ? { drawing, payload } : undefined;
}

export function SlicerEditorPanel({ drawingPayloads, drawings, onCommand, selectedDrawingIds = [], sheetId }: SlicerEditorPanelProps) {
  const selected = useMemo(() => selectedSlicer(drawings, drawingPayloads, selectedDrawingIds), [drawingPayloads, drawings, selectedDrawingIds]);
  const [settings, setSettings] = useState<PivotSlicerSettings | undefined>(selected?.payload.settings);
  const [caption, setCaption] = useState(selected?.payload.settings.caption ?? '');

  useEffect(() => {
    setSettings(selected?.payload.settings);
    setCaption(selected?.payload.settings.caption ?? '');
  }, [selected?.drawing.id, selected?.payload.settings]);

  if (!selected || !settings) {
    return <Panel className="h-full border-0 bg-transparent shadow-none"><PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Slicer Design</PanelTitle></PanelHeader><PanelBody className="p-4"><Text size="sm" tone="muted">Select a floating Slicer to edit its header, item layout, no-data and style settings.</Text></PanelBody></Panel>;
  }

  const applySettings = (patch: Partial<PivotSlicerSettings>): void => {
    const next = { ...settings, ...patch };
    if (!next.caption.trim() || next.columnCount < 1 || next.columnCount > 32 || next.itemHeight < 16 || next.itemHeight > 96) return;
    setSettings(next);
    onCommand({ commandId: 'pivot.control.slicer.settings.set', params: { sheetId, drawingId: selected.drawing.id, settings: next } });
  };
  const applyStyle = (style: PivotControlStyle): void => onCommand({ commandId: 'pivot.control.style.set', params: { sheetId, drawingId: selected.drawing.id, style } });

  return <Panel className="h-full border-0 bg-transparent shadow-none">
    <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Slicer Design</PanelTitle></PanelHeader>
    <PanelBody className="p-4"><Stack gap="md">
      <Text size="xs" tone="muted">{selected.payload.fieldId} · {selected.payload.pivotId}</Text>
      <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Caption</Text><TextInput value={caption} onChange={(event) => setCaption(event.target.value)} onBlur={() => applySettings({ caption })} /></Box>
      <Box className="grid grid-cols-2 gap-2">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Header</Text><Select value={String(settings.showHeader)} onChange={(event) => applySettings({ showHeader: event.target.value === 'true' })} sizeVariant="sm"><option value="true">Show</option><option value="false">Hide</option></Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Selection</Text><Select value={String(settings.multiSelect)} onChange={(event) => applySettings({ multiSelect: event.target.value === 'true' })} sizeVariant="sm"><option value="true">Multi-select</option><option value="false">Single-select</option></Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Item sort</Text><Select value={settings.sort} onChange={(event) => applySettings({ sort: event.target.value as PivotSlicerSettings['sort'] })} sizeVariant="sm"><option value="ascending">Ascending</option><option value="descending">Descending</option></Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Columns</Text><Select value={String(settings.columnCount)} onChange={(event) => applySettings({ columnCount: Number(event.target.value) })} sizeVariant="sm">{[1, 2, 3, 4].map((value) => <option key={value} value={value}>{value}</option>)}</Select></Box>
      </Box>
      <Box className="grid grid-cols-2 gap-2">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Item height</Text><TextInput type="number" min={16} max={96} value={String(settings.itemHeight)} onChange={(event) => applySettings({ itemHeight: Number(event.target.value) })} /></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">No-data style</Text><Select value={String(settings.showNoDataStyle)} onChange={(event) => applySettings({ showNoDataStyle: event.target.value === 'true' })} sizeVariant="sm"><option value="true">Indicate</option><option value="false">Normal</option></Select></Box>
      </Box>
      <Box className="grid grid-cols-2 gap-2">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">No-data items</Text><Select value={String(settings.showNoDataItems)} onChange={(event) => applySettings({ showNoDataItems: event.target.value === 'true' })} sizeVariant="sm"><option value="true">Show</option><option value="false">Hide</option></Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">No-data order</Text><Select value={String(settings.noDataItemsLast)} onChange={(event) => applySettings({ noDataItemsLast: event.target.value === 'true' })} sizeVariant="sm"><option value="true">Last</option><option value="false">Natural</option></Select></Box>
      </Box>
      <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Slicer style</Text><Box className="grid grid-cols-3 gap-1">{Object.entries(STYLE_GALLERY).map(([name, style]) => <Button key={name} variant="secondary" size="xs" onClick={() => applyStyle(style)}>{name}</Button>)}</Box></Box>
      <Button variant="secondary" size="sm" onClick={() => applySettings({ showHeader: !settings.showHeader })}>{settings.showHeader ? 'Hide header' : 'Show header'}</Button>
    </Stack></PanelBody>
  </Panel>;
}
