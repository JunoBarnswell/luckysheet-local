import React, { useState } from 'react';
import { Box, Button, ColorPicker, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DrawingObject, DrawingPayload, ShapeDrawingPayload } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface ShapeEditorPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

export function ShapeEditorPanel({
  sheetId,
  drawings,
  drawingPayloads,
  onCommand,
  onClose,
}: ShapeEditorPanelProps) {
  const [type, setType] = useState<ShapeDrawingPayload['type']>('rounded-rectangle');
  const [text, setText] = useState('Process Step');
  const [fill, setFill] = useState('#dbeafe');
  const [stroke, setStroke] = useState('#2563eb');
  const shapeEntries = drawings
    .filter((drawing) => drawing.kind === 'shape')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'shape' }> } => entry.payload?.kind === 'shape');

  const createId = (prefix: string): string => {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `${prefix}-${crypto.randomUUID()}`;
    return `${prefix}-${Date.now().toString(36)}`;
  };

  const handleCreate = () => {
    const payloadId = createId('shape');
    onCommand({
      commandId: 'drawing.add',
      params: {
        sheetId,
        drawing: {
          id: createId('drawing'),
          sheetId,
          kind: 'shape',
          anchor: { kind: 'absolute' },
          transform: { x: 120, y: 120, width: 160, height: 60, rotation: 0 },
          zIndex: 0,
          payloadId,
        },
        payload: {
          kind: 'shape',
          type,
          text,
          fill,
          stroke,
          strokeWidth: 2,
          textColor: '#1e3a8a',
          fontSize: 13,
        },
      },
    });
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Shape & Diagram Tools</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Shape Type
            </Text>
            <Select
              value={type}
              onChange={(e) => setType(e.target.value as ShapeDrawingPayload['type'])}
              sizeVariant="sm"
            >
              <option value="rounded-rectangle">Rounded Rectangle</option>
              <option value="rectangle">Rectangle</option>
              <option value="ellipse">Ellipse / Circle</option>
              <option value="arrow">Arrow</option>
              <option value="line">Line</option>
              <option value="callout">Callout / Note</option>
              <option value="star">Star Badge</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Shape Label
            </Text>
            <TextInput
              value={text}
              onChange={(e) => setText(e.target.value)}
              placeholder="e.g. Callout text"
            />
          </Box>

          <Box className="grid grid-cols-2 gap-2">
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Fill Color
              </Text>
              <Stack gap="xs">
                <ColorPicker color={fill} onChange={setFill} />
                <TextInput value={fill} onChange={(e) => setFill(e.target.value)} className="h-8 text-xs font-mono" />
              </Stack>
            </Box>
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Border Stroke
              </Text>
              <Stack gap="xs">
                <ColorPicker color={stroke} onChange={setStroke} />
                <TextInput value={stroke} onChange={(e) => setStroke(e.target.value)} className="h-8 text-xs font-mono" />
              </Stack>
            </Box>
          </Box>

          <Button variant="primary" size="sm" icon="plus" onClick={handleCreate}>
            Place Shape on Canvas
          </Button>

          {shapeEntries.length > 0 ? (
            <Box className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Floating Shapes ({shapeEntries.length})
              </Text>
              <Stack gap="xs">
                {shapeEntries.map(({ drawing, payload }) => (
                  <Box
                    key={drawing.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <Stack gap="none">
                      <Text size="sm" weight="medium" className="text-slate-800">{payload.text || payload.type}</Text>
                      <Text size="xs" tone="subtle">{payload.type}</Text>
                    </Stack>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onCommand({ commandId: 'drawing.remove', params: { sheetId, drawingId: drawing.id } })}
                      className="text-rose-600 hover:bg-rose-50"
                    />
                  </Box>
                ))}
              </Stack>
            </Box>
          ) : null}
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>
            Close Panel
          </Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
