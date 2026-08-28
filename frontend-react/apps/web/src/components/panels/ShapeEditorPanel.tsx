import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, ColorPicker, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ConnectorDrawingPayload, DrawingGroup, DrawingObject, DrawingPayload, ShapeDrawingPayload, WorksheetSnapSettings } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { SHAPE_DRAWING_PRESETS } from '@react-sheets/core-model';
import { SHAPE_META } from '../insert-ribbon-catalog';
import { insertText, type Locale } from '../../i18n';

export interface ShapeEditorPanelProps {
  locale?: Locale;
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  drawingGroups?: readonly DrawingGroup[];
  snapSettings?: WorksheetSnapSettings;
  selectedDrawingIds?: readonly string[];
  onSelectDrawing?: (drawingId: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
}

/** Selected-shape projection. Creation lives in the categorized Insert gallery. */
export function ShapeEditorPanel({
  locale,
  sheetId,
  drawings,
  drawingPayloads,
  selectedDrawingIds = [],
  onSelectDrawing,
  onCommand,
  drawingGroups = [],
  snapSettings,
}: ShapeEditorPanelProps) {
  const shapeEntries = useMemo(() => drawings
    .filter((drawing) => drawing.kind === 'shape')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'shape' }> } => entry.payload?.kind === 'shape'), [drawings, drawingPayloads]);
  const selectedEntry = shapeEntries.find((entry) => selectedDrawingIds.includes(entry.drawing.id)) ?? (selectedDrawingIds.length === 0 && shapeEntries.length === 1 ? shapeEntries[0] : undefined);
  const connectorEntries = useMemo(() => drawings
    .filter((drawing) => drawing.kind === 'connector')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: ConnectorDrawingPayload } => entry.payload?.kind === 'connector'), [drawings, drawingPayloads]);
  const selectedConnector = connectorEntries.find((entry) => selectedDrawingIds.includes(entry.drawing.id));
  const [draftText, setDraftText] = useState('');

  useEffect(() => {
    setDraftText(selectedEntry?.payload.text ?? '');
  }, [selectedEntry?.drawing.id, selectedEntry?.payload.text]);

  const updatePayload = (patch: Partial<ShapeDrawingPayload>) => {
    if (!selectedEntry) return;
    const before = structuredClone(selectedEntry.payload);
    const after: ShapeDrawingPayload = { ...before, ...patch };
    onCommand({ commandId: 'drawing.payload.update', params: { sheetId, payloadId: selectedEntry.drawing.payloadId, before, after } });
  };

  const commitText = () => {
    if (!selectedEntry || draftText === (selectedEntry.payload.text ?? '')) return;
    updatePayload({ text: draftText });
  };

  const updateConnector = (patch: Partial<ConnectorDrawingPayload>) => {
    if (!selectedConnector) return;
    const before = structuredClone(selectedConnector.payload);
    onCommand({ commandId: 'drawing.connector.update', params: { sheetId, drawingId: selectedConnector.drawing.id, before, after: { ...before, ...patch } } });
  };

  const selectedGroup = drawingGroups.find((group) => selectedDrawingIds.length === 1 && group.memberDrawingIds.includes(selectedDrawingIds[0]!));
  const canGroup = selectedDrawingIds.length >= 2 && selectedDrawingIds.every((id) => drawings.some((drawing) => drawing.id === id));

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Shape Format</PanelTitle></PanelHeader>
      <PanelBody className="p-4">
        <Stack gap="md">
          {selectedEntry ? (
            <>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">Shape Type</Text>
                <Select value={selectedEntry.payload.type} onChange={(event) => updatePayload({ type: event.target.value as ShapeDrawingPayload['type'] })} sizeVariant="sm">
                  {SHAPE_DRAWING_PRESETS.map((preset) => <option key={preset.type} value={preset.type}>{locale ? insertText(locale, SHAPE_META[preset.type].labelKey) : preset.type}</option>)}
                </Select>
              </Box>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">Shape Text</Text>
                <TextInput value={draftText} onChange={(event) => setDraftText(event.target.value)} onBlur={commitText} onKeyDown={(event) => { if (event.key === 'Enter') commitText(); }} />
              </Box>
              <Box className="grid grid-cols-2 gap-2">
                <Box>
                  <Text size="xs" weight="medium" className="mb-1 text-slate-700">Text direction</Text>
                  <Select value={selectedEntry.payload.textDirection ?? 'horizontal'} onChange={(event) => updatePayload({ textDirection: event.target.value as ShapeDrawingPayload['textDirection'] })} sizeVariant="sm">
                    <option value="horizontal">Horizontal</option>
                    <option value="vertical">Vertical</option>
                  </Select>
                </Box>
                <Box>
                  <Text size="xs" weight="medium" className="mb-1 text-slate-700">Text alignment</Text>
                  <Select value={selectedEntry.payload.textAlignment ?? 'center'} onChange={(event) => updatePayload({ textAlignment: event.target.value as ShapeDrawingPayload['textAlignment'] })} sizeVariant="sm">
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                  </Select>
                </Box>
              </Box>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">Shape hyperlink</Text>
                <TextInput value={selectedEntry.payload.hyperlink?.kind === 'url' ? selectedEntry.payload.hyperlink.url : ''} placeholder="https://..." onBlur={(event) => {
                  const url = event.target.value.trim();
                  updatePayload({ hyperlink: url ? { kind: 'url', url } : undefined });
                }} />
              </Box>
              <Button size="xs" variant={selectedEntry.payload.effects?.shadow ? 'soft' : 'ghost'} onClick={() => updatePayload({ effects: selectedEntry.payload.effects?.shadow ? undefined : { shadow: { color: '#00000055', blur: 8, offsetX: 2, offsetY: 2, opacity: 0.35 } } })}>Shadow effect</Button>
              <Box className="grid grid-cols-2 gap-2">
                <Box>
                  <Text size="xs" weight="medium" className="mb-1 text-slate-700">Fill</Text>
                  <Stack gap="xs"><ColorPicker color={selectedEntry.payload.fill} onChange={(value) => updatePayload({ fill: value })} /><TextInput value={selectedEntry.payload.fill} onChange={(event) => updatePayload({ fill: event.target.value })} className="h-8 text-xs font-mono" /></Stack>
                </Box>
                <Box>
                  <Text size="xs" weight="medium" className="mb-1 text-slate-700">Outline</Text>
                  <Stack gap="xs"><ColorPicker color={selectedEntry.payload.stroke} onChange={(value) => updatePayload({ stroke: value })} /><TextInput value={selectedEntry.payload.stroke} onChange={(event) => updatePayload({ stroke: event.target.value })} className="h-8 text-xs font-mono" /></Stack>
                </Box>
              </Box>
            </>
          ) : selectedConnector ? (
            <>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">Connector Type</Text>
                <Select value={selectedConnector.payload.connectorType} onChange={(event) => updateConnector({ connectorType: event.target.value as ConnectorDrawingPayload['connectorType'] })} sizeVariant="sm">
                  <option value="straight">Straight</option>
                  <option value="elbow">Elbow</option>
                  <option value="curved">Curved</option>
                </Select>
              </Box>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">End arrow</Text>
                <Select value={selectedConnector.payload.endArrowhead} onChange={(event) => updateConnector({ endArrowhead: event.target.value as ConnectorDrawingPayload['endArrowhead'] })} sizeVariant="sm">
                  <option value="none">None</option>
                  <option value="triangle">Triangle</option>
                  <option value="stealth">Stealth</option>
                  <option value="diamond">Diamond</option>
                  <option value="oval">Oval</option>
                </Select>
              </Box>
            </>
          ) : (
            <Box className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-3">
              <Text size="sm" weight="medium" className="text-slate-700">Select a shape or connector</Text>
              <Text size="xs" tone="subtle" className="mt-1">Connector routes, group ownership and worksheet snap settings use the canonical drawing commands.</Text>
            </Box>
          )}
          <Box className="grid grid-cols-2 gap-2 border-t border-slate-200 pt-3">
            <Button size="xs" variant="soft" disabled={!canGroup} onClick={() => onCommand({ commandId: 'drawing.group', params: { sheetId, group: { id: `${sheetId}:group:${selectedDrawingIds.join(',')}`, sheetId, memberDrawingIds: [...selectedDrawingIds] } } })}>Group</Button>
            <Button size="xs" variant="ghost" disabled={!selectedGroup} onClick={() => selectedGroup && onCommand({ commandId: 'drawing.ungroup', params: { sheetId, groupId: selectedGroup.id } })}>Ungroup</Button>
            <Button size="xs" variant={snapSettings?.enabled ? 'soft' : 'ghost'} onClick={() => snapSettings && onCommand({ commandId: 'drawing.snapSettings.set', params: { sheetId, before: snapSettings, after: { ...snapSettings, enabled: !snapSettings.enabled } } })}>Snap {snapSettings?.enabled ? 'On' : 'Off'}</Button>
          </Box>
          <Box className="border-t border-slate-200 pt-3">
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Shapes and connectors ({shapeEntries.length + connectorEntries.length})</Text>
            {shapeEntries.length + connectorEntries.length === 0 ? <Text size="xs" tone="subtle">No shapes or connectors on this worksheet.</Text> : (
              <Stack gap="xs">
                {shapeEntries.map(({ drawing, payload }) => (
                  <Box key={drawing.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs">
                    <Button variant={selectedDrawingIds.includes(drawing.id) ? 'soft' : 'ghost'} size="xs" className="min-w-0 flex-1 justify-start" onClick={() => onSelectDrawing?.(drawing.id)}>
                      <Stack gap="none" className="text-left"><Text size="sm" weight="medium" className="text-slate-800">{payload.text || payload.type}</Text><Text size="xs" tone="subtle">{payload.type}</Text></Stack>
                    </Button>
                    <Button variant="ghost" size="xs" icon="trash" iconOnly onClick={() => onCommand({ commandId: 'drawing.remove', params: { sheetId, drawingId: drawing.id } })} className="text-rose-600 hover:bg-rose-50" />
                  </Box>
                ))}
                {connectorEntries.map(({ drawing, payload }) => (
                  <Box key={drawing.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs">
                    <Button variant={selectedDrawingIds.includes(drawing.id) ? 'soft' : 'ghost'} size="xs" className="min-w-0 flex-1 justify-start" onClick={() => onSelectDrawing?.(drawing.id)}>
                      <Stack gap="none" className="text-left"><Text size="sm" weight="medium" className="text-slate-800">{payload.connectorType} connector</Text><Text size="xs" tone="subtle">{payload.start.drawingId} → {payload.end.drawingId}</Text></Stack>
                    </Button>
                    <Button variant="ghost" size="xs" icon="trash" iconOnly onClick={() => onCommand({ commandId: 'drawing.remove', params: { sheetId, drawingId: drawing.id } })} className="text-rose-600 hover:bg-rose-50" />
                  </Box>
                ))}
              </Stack>
            )}
          </Box>
        </Stack>
      </PanelBody>
    </Panel>
  );
}
