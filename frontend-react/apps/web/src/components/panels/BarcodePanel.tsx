import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { BARCODE_SYMBOLOGIES, type BarcodeCellPresentation, type BarcodeLabelPosition, type BarcodeSymbology, type RangeRef } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { parseAddress, type CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';

export interface BarcodePanelProps {
  sheetId: string;
  sheet: CanvasSheetSnapshot;
  activeCell: string;
  selectedRange?: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  initialSymbology: BarcodeSymbology;
  onCommand: (descriptor: CommandDescriptor) => void;
}

const symbologies = BARCODE_SYMBOLOGIES;

function defaultPresentation(symbology: BarcodeSymbology): BarcodeCellPresentation {
  return {
    kind: 'barcode', symbology, source: { kind: 'cell-value' }, parameters: { symbology },
    options: { foreground: '#111827', background: '#ffffff', showText: !['qr', 'data-matrix'].includes(symbology), labelPosition: ['qr', 'data-matrix'].includes(symbology) ? 'none' : 'below', quietZone: 2 },
  };
}

export function BarcodePanel({ sheetId, sheet, activeCell, selectedRange, initialSymbology, onCommand }: BarcodePanelProps) {
  const address = parseAddress(activeCell);
  const selectedCell = address ? sheet.getCell(address.row, address.column) : undefined;
  const existing = selectedCell?.presentation?.kind === 'barcode' ? selectedCell.presentation : undefined;
  const [symbology, setSymbology] = useState<BarcodeSymbology>(existing?.symbology ?? initialSymbology);
  const [sourceKind, setSourceKind] = useState<'cell-value' | 'formula'>(existing?.source.kind ?? 'cell-value');
  const [formula, setFormula] = useState(existing?.source.kind === 'formula' ? existing.source.formula : '=A1');
  const [foreground, setForeground] = useState(existing?.options.foreground ?? '#111827');
  const [background, setBackground] = useState(existing?.options.background ?? '#ffffff');
  const [showText, setShowText] = useState(existing?.options.showText ?? true);
  const [labelPosition, setLabelPosition] = useState<BarcodeLabelPosition>(existing?.options.labelPosition ?? 'below');
  const [quietZone, setQuietZone] = useState(String(existing?.options.quietZone ?? 2));
  const [fontSize, setFontSize] = useState(String(existing?.options.fontSize ?? 10));
  const [addOnText, setAddOnText] = useState(existing?.parameters.symbology === symbology && 'addOnText' in existing.parameters ? existing.parameters.addOnText ?? '' : '');
  const [includeCheckDigit, setIncludeCheckDigit] = useState(existing?.parameters.symbology === symbology && 'includeCheckDigit' in existing.parameters ? existing.parameters.includeCheckDigit !== false : true);
  const [wideNarrowRatio, setWideNarrowRatio] = useState(String(existing?.parameters.symbology === symbology && 'wideNarrowRatio' in existing.parameters ? existing.parameters.wideNarrowRatio ?? 2 : 2));
  const [fullAscii, setFullAscii] = useState(existing?.parameters.symbology === symbology && 'fullAscii' in existing.parameters ? existing.parameters.fullAscii === true : false);
  const [errorCorrection, setErrorCorrection] = useState<'low' | 'medium' | 'quartile' | 'high'>(existing?.parameters.symbology === symbology && 'errorCorrection' in existing.parameters ? existing.parameters.errorCorrection ?? 'medium' : 'medium');
  const [pdf417SecurityLevel, setPdf417SecurityLevel] = useState(String(existing?.parameters.symbology === 'pdf417' && 'securityLevel' in existing.parameters ? existing.parameters.securityLevel ?? 2 : 2));

  useEffect(() => {
    const next = selectedCell?.presentation?.kind === 'barcode' ? selectedCell.presentation : undefined;
    const nextSymbology = next?.symbology ?? initialSymbology;
    const defaults = next ?? defaultPresentation(nextSymbology);
    setSymbology(nextSymbology);
    setSourceKind(defaults.source.kind);
    setFormula(defaults.source.kind === 'formula' ? defaults.source.formula : '=A1');
    setForeground(defaults.options.foreground);
    setBackground(defaults.options.background);
    setShowText(defaults.options.showText);
    setLabelPosition(defaults.options.labelPosition);
    setQuietZone(String(defaults.options.quietZone));
    setFontSize(String(defaults.options.fontSize ?? 10));
    setAddOnText('addOnText' in defaults.parameters ? defaults.parameters.addOnText ?? '' : '');
    setIncludeCheckDigit('includeCheckDigit' in defaults.parameters ? defaults.parameters.includeCheckDigit !== false : true);
    setWideNarrowRatio('wideNarrowRatio' in defaults.parameters ? String(defaults.parameters.wideNarrowRatio ?? 2) : '2');
    setFullAscii('fullAscii' in defaults.parameters ? defaults.parameters.fullAscii === true : false);
    setErrorCorrection('errorCorrection' in defaults.parameters ? defaults.parameters.errorCorrection ?? 'medium' : 'medium');
    setPdf417SecurityLevel('securityLevel' in defaults.parameters ? String(defaults.parameters.securityLevel ?? 2) : '2');
  }, [activeCell, initialSymbology, selectedCell?.presentation]);

  const range: RangeRef = useMemo(() => ({ sheetId, startRow: selectedRange?.startRow ?? address?.row ?? 0, endRow: selectedRange?.endRow ?? address?.row ?? 0, startColumn: selectedRange?.startColumn ?? address?.column ?? 0, endColumn: selectedRange?.endColumn ?? address?.column ?? 0 }), [address?.column, address?.row, selectedRange, sheetId]);
  const linear = ['code128', 'code39', 'code93', 'code49', 'codabar', 'gs1-128'].includes(symbology);
  const ean = ['ean13', 'ean8', 'upca'].includes(symbology);
  const qr = symbology === 'qr';
  const pdf417 = symbology === 'pdf417';
  const matrix = symbology === 'data-matrix';

  const apply = (): void => {
    const parameters: BarcodeCellPresentation['parameters'] = qr
      ? { symbology: 'qr', errorCorrection }
      : pdf417
        ? { symbology: 'pdf417', securityLevel: Number(pdf417SecurityLevel) }
        : matrix
          ? { symbology: 'data-matrix' }
      : ean
        ? { symbology: symbology as 'ean13' | 'ean8' | 'upca', addOnText: addOnText || undefined, includeCheckDigit }
        : { symbology: symbology as 'code128' | 'code39' | 'code93' | 'code49' | 'codabar' | 'gs1-128', fullAscii, includeCheckDigit, wideNarrowRatio: Number(wideNarrowRatio) };
    const presentation: BarcodeCellPresentation = {
      kind: 'barcode', symbology, source: sourceKind === 'formula' ? { kind: 'formula', formula } : { kind: 'cell-value' }, parameters,
      options: { foreground, background, showText, labelPosition, quietZone: Number(quietZone), fontSize: Number(fontSize) },
    };
    onCommand({ commandId: 'cell.barcode.apply', params: { sheetId, ranges: [range], presentation } });
  };

  return (
    <Panel className="border-0 bg-transparent shadow-none">
      <PanelHeader><PanelTitle size="sm">Barcode · Type and Parameters</PanelTitle></PanelHeader>
      <PanelBody><Stack gap="md">
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Symbology</Text><Select value={symbology} onChange={(event) => { const next = event.target.value as BarcodeSymbology; setSymbology(next); const defaults = defaultPresentation(next); setShowText(defaults.options.showText); setLabelPosition(defaults.options.labelPosition); setAddOnText(''); setIncludeCheckDigit(true); setWideNarrowRatio('2'); setFullAscii(false); setErrorCorrection('medium'); setPdf417SecurityLevel('2'); }} sizeVariant="sm">{symbologies.map((type) => <option key={type} value={type}>{type}</option>)}</Select></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Source</Text><Select value={sourceKind} onChange={(event) => setSourceKind(event.target.value as 'cell-value' | 'formula')} sizeVariant="sm"><option value="cell-value">Selected cell value</option><option value="formula">Formula source</option></Select>{sourceKind === 'formula' ? <TextInput className="mt-1" value={formula} onChange={(event) => setFormula(event.target.value)} placeholder={'=A1&"-"&B1'} /> : null}</Box>
        <Box className="grid grid-cols-2 gap-2"><Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Foreground</Text><TextInput type="color" value={foreground} onChange={(event) => setForeground(event.target.value)} /></Box><Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Background</Text><TextInput type="color" value={background} onChange={(event) => setBackground(event.target.value)} /></Box></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Label</Text><Stack gap="xs"><Select value={labelPosition} onChange={(event) => { const value = event.target.value as BarcodeLabelPosition; setLabelPosition(value); setShowText(value !== 'none'); }} sizeVariant="sm"><option value="none">Hidden</option><option value="below">Below</option><option value="above">Above</option></Select><TextInput type="number" value={fontSize} onChange={(event) => setFontSize(event.target.value)} placeholder="Font size" /></Stack></Box>
        <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Quiet zone</Text><TextInput type="number" value={quietZone} onChange={(event) => setQuietZone(event.target.value)} /></Box>
        {qr ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">QR error correction</Text><Select value={errorCorrection} onChange={(event) => setErrorCorrection(event.target.value as typeof errorCorrection)} sizeVariant="sm"><option value="low">Low</option><option value="medium">Medium</option><option value="quartile">Quartile</option><option value="high">High</option></Select></Box> : null}
        {pdf417 ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">PDF417 security level</Text><TextInput type="number" min={0} max={8} value={pdf417SecurityLevel} onChange={(event) => setPdf417SecurityLevel(event.target.value)} /></Box> : null}
        {ean ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">EAN/UPC parameters</Text><TextInput value={addOnText} onChange={(event) => setAddOnText(event.target.value)} placeholder="Add-on digits (optional)" /><Button className="mt-1" variant="secondary" size="sm" onClick={() => setIncludeCheckDigit((value) => !value)}>{includeCheckDigit ? 'Include check digit' : 'Exclude check digit'}</Button></Box> : null}
        {linear ? <Box><Text size="xs" weight="medium" className="mb-1 text-slate-700">Linear parameters</Text><Stack gap="xs"><Button variant="secondary" size="sm" onClick={() => setFullAscii((value) => !value)}>{fullAscii ? 'Full ASCII enabled' : 'Full ASCII disabled'}</Button><TextInput type="number" value={wideNarrowRatio} onChange={(event) => setWideNarrowRatio(event.target.value)} placeholder="Wide/narrow ratio" /><Button variant="secondary" size="sm" onClick={() => setIncludeCheckDigit((value) => !value)}>{includeCheckDigit ? 'Include check digit' : 'Exclude check digit'}</Button></Stack></Box> : null}
        <Button variant="primary" icon="check" onClick={apply}>Apply Barcode</Button>
        <Text size="xs" tone="muted">Applies to the current selection as one Undo/Redo transaction. Invalid source values are rejected before any cell changes.</Text>
      </Stack></PanelBody>
    </Panel>
  );
}
