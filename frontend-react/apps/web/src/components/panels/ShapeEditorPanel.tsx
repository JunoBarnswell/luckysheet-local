import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DrawingObject, DrawingPayload, ShapeDrawingPayload } from '@react-sheets/core-model';
import type { CommandDescriptor } from '../../domain/command-descriptor';

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
          <div>
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
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Shape Label
            </Text>
            <TextInput
              value={text}
              onChange={(e) => setText(e.target.value)}
              placeholder="e.g. Callout text"
            />
          </div>

          <div className="grid grid-cols-2 gap-2">
            <div>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Fill Color
              </Text>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  value={fill}
                  onChange={(e) => setFill(e.target.value)}
                  className="h-8 w-8 cursor-pointer rounded border border-slate-300 p-0"
                />
                <TextInput value={fill} onChange={(e) => setFill(e.target.value)} className="h-8 text-xs font-mono" />
              </div>
            </div>
            <div>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Border Stroke
              </Text>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  value={stroke}
                  onChange={(e) => setStroke(e.target.value)}
                  className="h-8 w-8 cursor-pointer rounded border border-slate-300 p-0"
                />
                <TextInput value={stroke} onChange={(e) => setStroke(e.target.value)} className="h-8 text-xs font-mono" />
              </div>
            </div>
          </div>

          <Button variant="primary" size="sm" icon="plus" onClick={handleCreate}>
            Place Shape on Canvas
          </Button>

          {shapeEntries.length > 0 ? (
            <div className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Floating Shapes ({shapeEntries.length})
              </Text>
              <Stack gap="xs">
                {shapeEntries.map(({ drawing, payload }) => (
                  <div
                    key={drawing.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <div>
                      <div className="font-medium text-slate-800">{payload.text || payload.type}</div>
                      <div className="text-[10px] text-slate-500">{payload.type}</div>
                    </div>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onCommand({ commandId: 'drawing.remove', params: { sheetId, drawingId: drawing.id } })}
                      className="text-rose-600 hover:bg-rose-50"
                    />
                  </div>
                ))}
              </Stack>
            </div>
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
