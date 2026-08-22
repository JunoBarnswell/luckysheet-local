import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ShapeModel } from '@react-sheets/core-model';

export interface ShapeEditorPanelProps {
  sheetId: string;
  shapes: ShapeModel[];
  onAddShape: (shape: ShapeModel) => void;
  onRemoveShape: (id: string) => void;
  onClose?: () => void;
}

export function ShapeEditorPanel({
  sheetId,
  shapes,
  onAddShape,
  onRemoveShape,
  onClose,
}: ShapeEditorPanelProps) {
  const [type, setType] = useState<ShapeModel['type']>('rounded-rectangle');
  const [text, setText] = useState('Process Step');
  const [fill, setFill] = useState('#dbeafe');
  const [stroke, setStroke] = useState('#2563eb');

  const handleCreate = () => {
    const newShape: ShapeModel = {
      id: 'shape-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      type,
      text,
      fill,
      stroke,
      strokeWidth: 2,
      textColor: '#1e3a8a',
      fontSize: 13,
      bounds: {
        x: 120,
        y: 120,
        width: 160,
        height: 60,
      },
    };
    onAddShape(newShape);
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
              onChange={(e) => setType(e.target.value as ShapeModel['type'])}
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

          {shapes.length > 0 ? (
            <div className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Floating Shapes ({shapes.length})
              </Text>
              <Stack gap="xs">
                {shapes.map((s) => (
                  <div
                    key={s.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <div>
                      <div className="font-medium text-slate-800">{s.text || s.type}</div>
                      <div className="text-[10px] text-slate-500">{s.type}</div>
                    </div>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemoveShape(s.id)}
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
