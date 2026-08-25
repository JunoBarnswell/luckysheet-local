import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { ImageCellPresentation, ImageDrawingPayload, ImageEffects } from '@react-sheets/core-model';
import { parseAddress, type CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';

export interface PicturePanelProps {
  sheetId: string;
  activeCell: string;
  sheet: CanvasSheetSnapshot;
  selectedDrawingIds?: readonly string[];
  onCommand: (descriptor: CommandDescriptor) => void;
}

function idFor(prefix: string): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Date.now().toString(36)}`;
}

function percent(value: number | undefined): string {
  return value === undefined ? '0' : String(Math.round(value * 100));
}

function fraction(value: string): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed / 100 : Number.NaN;
}

function cropOf(source: ImageCellPresentation | ImageDrawingPayload): { left: string; top: string; right: string; bottom: string } {
  return {
    left: percent(source.crop?.left),
    top: percent(source.crop?.top),
    right: percent(source.crop?.right),
    bottom: percent(source.crop?.bottom),
  };
}

function effectsOf(source: ImageCellPresentation | ImageDrawingPayload): { brightness: string; contrast: string; transparency: string } {
  return {
    brightness: percent(source.effects?.brightness),
    contrast: percent(source.effects?.contrast),
    transparency: percent(source.effects?.transparency),
  };
}

export function PicturePanel({ sheetId, activeCell, sheet, selectedDrawingIds = [], onCommand }: PicturePanelProps) {
  const activeAddress = parseAddress(activeCell);
  const selectedEntry = useMemo(() => {
    const selectedId = selectedDrawingIds[0];
    const drawing = selectedId ? sheet.drawings.find((entry) => entry.id === selectedId && entry.kind === 'image') : undefined;
    const payload = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
    return drawing && payload?.kind === 'image' ? { drawing, payload } : undefined;
  }, [selectedDrawingIds, sheet.drawings, sheet.drawingPayloads]);
  const cell = activeAddress ? sheet.getCell(activeAddress.row, activeAddress.column) : undefined;
  const cellImage = cell?.presentation?.kind === 'image' ? cell.presentation : undefined;
  const source = selectedEntry?.payload ?? cellImage;
  const [altText, setAltText] = useState('');
  const [fit, setFit] = useState<ImageCellPresentation['fit']>('contain');
  const [left, setLeft] = useState('0');
  const [top, setTop] = useState('0');
  const [right, setRight] = useState('0');
  const [bottom, setBottom] = useState('0');
  const [brightness, setBrightness] = useState('0');
  const [contrast, setContrast] = useState('0');
  const [transparency, setTransparency] = useState('0');

  useEffect(() => {
    if (!source) {
      setAltText('');
      setFit('contain');
      setLeft('0');
      setTop('0');
      setRight('0');
      setBottom('0');
      setBrightness('0');
      setContrast('0');
      setTransparency('0');
      return;
    }
    setAltText(source.altText ?? '');
    setFit(source.kind === 'image' && 'fit' in source ? source.fit : 'contain');
    const crop = cropOf(source);
    setLeft(crop.left);
    setTop(crop.top);
    setRight(crop.right);
    setBottom(crop.bottom);
    const effects = effectsOf(source);
    setBrightness(effects.brightness);
    setContrast(effects.contrast);
    setTransparency(effects.transparency);
  }, [source?.src, selectedEntry?.drawing.id, activeAddress?.row, activeAddress?.column]);

  if (!source) {
    return (
      <Panel className="h-full border-0 bg-transparent shadow-none">
        <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Picture Format</PanelTitle></PanelHeader>
        <PanelBody className="p-4"><Text size="sm" tone="muted">Select a floating picture or a cell containing a picture.</Text></PanelBody>
      </Panel>
    );
  }

  const applyCellPresentation = (patch: Partial<ImageCellPresentation>): void => {
    if (!cellImage || !activeAddress) return;
    onCommand({ commandId: 'cell.image.apply', params: { sheetId, row: activeAddress.row, column: activeAddress.column, presentation: { ...structuredClone(cellImage), ...structuredClone(patch) } } });
  };
  const applyAltText = (): void => {
    if (selectedEntry) onCommand({ commandId: 'drawing.image.altText', params: { sheetId, drawingId: selectedEntry.drawing.id, altText } });
    else applyCellPresentation({ altText });
  };
  const readCrop = () => ({ left: fraction(left), top: fraction(top), right: fraction(right), bottom: fraction(bottom) });
  const readEffects = (): ImageEffects => ({ brightness: fraction(brightness), contrast: fraction(contrast), transparency: fraction(transparency) });
  const applyCrop = (): void => {
    const crop = readCrop();
    if (selectedEntry) onCommand({ commandId: 'drawing.image.crop', params: { sheetId, drawingId: selectedEntry.drawing.id, crop } });
    else applyCellPresentation({ crop });
  };
  const applyEffects = (): void => {
    const effects = readEffects();
    if (selectedEntry) onCommand({ commandId: 'drawing.image.effects', params: { sheetId, drawingId: selectedEntry.drawing.id, effects } });
    else applyCellPresentation({ effects });
  };
  const convertToCell = (): void => {
    if (!selectedEntry || !activeAddress) return;
    onCommand({ commandId: 'picture.convertToCell', params: { sheetId, drawingId: selectedEntry.drawing.id, row: activeAddress.row, column: activeAddress.column } });
  };
  const convertToFloating = (): void => {
    if (!cellImage || !activeAddress) return;
    onCommand({ commandId: 'picture.convertToFloating', params: { sheetId, row: activeAddress.row, column: activeAddress.column, drawingId: idFor('drawing'), payloadId: idFor('image') } });
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4"><PanelTitle size="sm">Picture Format</PanelTitle></PanelHeader>
      <PanelBody className="p-4"><Stack gap="md">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Alternative text</Text><TextInput value={altText} onChange={(event) => setAltText(event.target.value)} onBlur={applyAltText} placeholder="Describe this picture" /></Box>
        {!selectedEntry ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Cell fit</Text><Select value={fit} onChange={(event) => { const nextFit = event.target.value as ImageCellPresentation['fit']; setFit(nextFit); applyCellPresentation({ fit: nextFit }); }} sizeVariant="sm"><option value="contain">Contain</option><option value="cover">Cover</option><option value="stretch">Stretch</option></Select></Box> : null}
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Crop (%)</Text><Box className="grid grid-cols-4 gap-1"><TextInput value={left} onChange={(event) => setLeft(event.target.value)} aria-label="Crop left" /><TextInput value={top} onChange={(event) => setTop(event.target.value)} aria-label="Crop top" /><TextInput value={right} onChange={(event) => setRight(event.target.value)} aria-label="Crop right" /><TextInput value={bottom} onChange={(event) => setBottom(event.target.value)} aria-label="Crop bottom" /></Box><Button variant="secondary" size="sm" className="mt-2" onClick={applyCrop}>Apply crop</Button></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Adjustments (%)</Text><Box className="grid grid-cols-3 gap-1"><TextInput value={brightness} onChange={(event) => setBrightness(event.target.value)} aria-label="Brightness" /><TextInput value={contrast} onChange={(event) => setContrast(event.target.value)} aria-label="Contrast" /><TextInput value={transparency} onChange={(event) => setTransparency(event.target.value)} aria-label="Transparency" /></Box><Button variant="secondary" size="sm" className="mt-2" onClick={applyEffects}>Apply adjustments</Button></Box>
        <Box className="flex flex-wrap gap-2">{selectedEntry ? <Button variant="secondary" size="sm" onClick={convertToCell} disabled={!activeAddress}>Place in active cell</Button> : <Button variant="secondary" size="sm" onClick={convertToFloating} disabled={!activeAddress}>Float over cells</Button>}</Box>
      </Stack></PanelBody>
    </Panel>
  );
}
