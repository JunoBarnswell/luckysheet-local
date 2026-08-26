import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, ColorPicker, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DrawingObject, DrawingPayload, ShapeDrawingPayload } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface ShapeEditorPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  onSelectDrawing?: (drawingId: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
}

/** Selected-shape projection. Creation lives in the categorized Insert gallery. */
export function ShapeEditorPanel({
  sheetId,
  drawings,
  drawingPayloads,
  selectedDrawingIds = [],
  onSelectDrawing,
  onCommand,
}: ShapeEditorPanelProps) {
  const shapeEntries = useMemo(() => drawings
    .filter((drawing) => drawing.kind === 'shape')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'shape' }> } => entry.payload?.kind === 'shape'), [drawings, drawingPayloads]);
  const selectedEntry = shapeEntries.find((entry) => selectedDrawingIds.includes(entry.drawing.id)) ?? (selectedDrawingIds.length === 0 && shapeEntries.length === 1 ? shapeEntries[0] : undefined);
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
                  <option value="rectangle">Rectangle</option>
                  <option value="rounded-rectangle">Rounded rectangle</option>
                  <option value="ellipse">Ellipse</option>
                  <option value="line">Line</option>
                  <option value="arrow">Arrow</option>
                  <option value="callout">Callout</option>
                  <option value="star">Star</option>
                </Select>
              </Box>
              <Box>
                <Text size="xs" weight="medium" className="mb-1 text-slate-700">Shape Text</Text>
                <TextInput value={draftText} onChange={(event) => setDraftText(event.target.value)} onBlur={commitText} onKeyDown={(event) => { if (event.key === 'Enter') commitText(); }} />
              </Box>
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
          ) : (
            <Box className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-3">
              <Text size="sm" weight="medium" className="text-slate-700">Select a supported shape</Text>
              <Text size="xs" tone="subtle" className="mt-1">Insert shapes from the categorized Shapes gallery. Connector, group, snap, and shape-link features are unavailable until their canonical model contracts exist.</Text>
            </Box>
          )}
          <Box className="border-t border-slate-200 pt-3">
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Shapes ({shapeEntries.length})</Text>
            {shapeEntries.length === 0 ? <Text size="xs" tone="subtle">No shapes on this worksheet.</Text> : (
              <Stack gap="xs">
                {shapeEntries.map(({ drawing, payload }) => (
                  <Box key={drawing.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs">
                    <Button variant={selectedDrawingIds.includes(drawing.id) ? 'soft' : 'ghost'} size="xs" className="min-w-0 flex-1 justify-start" onClick={() => onSelectDrawing?.(drawing.id)}>
                      <Stack gap="none" className="text-left"><Text size="sm" weight="medium" className="text-slate-800">{payload.text || payload.type}</Text><Text size="xs" tone="subtle">{payload.type}</Text></Stack>
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
