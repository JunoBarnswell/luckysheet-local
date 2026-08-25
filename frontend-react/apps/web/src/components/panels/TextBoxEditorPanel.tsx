import React, { useEffect, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput, Textarea } from '@react-sheets/ui-system';
import type { DrawingObject, DrawingPayload, TextBoxDrawingPayload, TextBoxTextFrame } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface TextBoxEditorPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

function clonePayload(payload: TextBoxDrawingPayload): TextBoxDrawingPayload {
  return structuredClone(payload);
}

export function TextBoxEditorPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], onCommand, onClose }: TextBoxEditorPanelProps) {
  const entry = (() => {
    const selected = selectedDrawingIds.map((id) => drawings.find((drawing) => drawing.id === id)).find((drawing): drawing is DrawingObject => drawing?.kind === 'textbox');
    const drawing = selected ?? drawings.find((candidate) => candidate.kind === 'textbox');
    const payload = drawing ? drawingPayloads.get(drawing.payloadId) : undefined;
    return drawing && payload?.kind === 'textbox' ? { drawing, payload } : null;
  })();
  const [draft, setDraft] = useState<TextBoxDrawingPayload | null>(entry ? clonePayload(entry.payload) : null);

  useEffect(() => setDraft(entry ? clonePayload(entry.payload) : null), [entry?.drawing.id, entry?.payload]);

  if (!entry || !draft) {
    return <Panel className="h-full border-0 bg-transparent shadow-none"><PanelBody><Text tone="muted">Select a text box to edit its text frame.</Text></PanelBody></Panel>;
  }

  const updateFrame = (patch: Partial<TextBoxTextFrame>) => setDraft((current) => current ? { ...current, textFrame: { ...current.textFrame, ...patch } } : current);
  const apply = () => onCommand({ commandId: 'drawing.textbox.update', params: { sheetId, drawingId: entry.drawing.id, payload: draft } });

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Text Box Format</PanelTitle></PanelHeader>
      <PanelBody className="p-4">
        <Stack gap="md">
          <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Text</Text><Textarea value={draft.text} rows={4} onChange={(event) => setDraft({ ...draft, text: event.target.value })} /></Box>
          <Box className="grid grid-cols-2 gap-2">
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Font family</Text><TextInput value={draft.textFrame.fontFamily} onChange={(event) => updateFrame({ fontFamily: event.target.value })} /></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Size</Text><TextInput type="number" min={1} value={draft.textFrame.fontSize} onChange={(event) => updateFrame({ fontSize: Math.max(1, Number(event.target.value) || 1) })} /></Box>
          </Box>
          <Box className="grid grid-cols-2 gap-2">
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Direction</Text><Select value={draft.textFrame.direction} onChange={(event) => updateFrame({ direction: event.target.value as TextBoxTextFrame['direction'] })}><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></Select></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Wrap</Text><Button size="sm" variant={draft.textFrame.wrap ? 'primary' : 'ghost'} onClick={() => updateFrame({ wrap: !draft.textFrame.wrap })}>{draft.textFrame.wrap ? 'On' : 'Off'}</Button></Box>
          </Box>
          <Box className="grid grid-cols-4 gap-2">
            {(['top', 'right', 'bottom', 'left'] as const).map((side) => <TextInput key={side} aria-label={`Margin ${side}`} type="number" min={0} value={draft.textFrame.margin[side]} onChange={(event) => setDraft({ ...draft, textFrame: { ...draft.textFrame, margin: { ...draft.textFrame.margin, [side]: Math.max(0, Number(event.target.value) || 0) } } })} />)}
          </Box>
          <Box className="grid grid-cols-2 gap-2">
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Horizontal</Text><Select value={draft.textFrame.horizontalAlignment} onChange={(event) => updateFrame({ horizontalAlignment: event.target.value as TextBoxTextFrame['horizontalAlignment'] })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></Select></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Vertical</Text><Select value={draft.textFrame.verticalAlignment} onChange={(event) => updateFrame({ verticalAlignment: event.target.value as TextBoxTextFrame['verticalAlignment'] })}><option value="top">Top</option><option value="middle">Middle</option><option value="bottom">Bottom</option></Select></Box>
          </Box>
          <Box className="grid grid-cols-2 gap-2">
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Text color</Text><TextInput value={draft.textFrame.textColor} onChange={(event) => updateFrame({ textColor: event.target.value })} /></Box>
            <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Autofit</Text><Select value={draft.textFrame.autofit} onChange={(event) => updateFrame({ autofit: event.target.value as TextBoxTextFrame['autofit'] })}><option value="none">None</option><option value="shrink-text">Shrink text</option><option value="resize-shape">Resize shape</option></Select></Box>
          </Box>
          <Box className="grid grid-cols-3 gap-2">
            <Button size="sm" variant={draft.textFrame.bold ? 'primary' : 'ghost'} onClick={() => updateFrame({ bold: !draft.textFrame.bold })}>Bold</Button>
            <Button size="sm" variant={draft.textFrame.italic ? 'primary' : 'ghost'} onClick={() => updateFrame({ italic: !draft.textFrame.italic })}>Italic</Button>
            <Button size="sm" variant={draft.textFrame.underline ? 'primary' : 'ghost'} onClick={() => updateFrame({ underline: !draft.textFrame.underline })}>Underline</Button>
          </Box>
          <Button variant="primary" size="sm" onClick={apply}>Apply text-frame</Button>
        </Stack>
      </PanelBody>
      {onClose ? <PanelFooter className="border-t border-slate-200 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button></PanelFooter> : null}
    </Panel>
  );
}
