import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text } from '@react-sheets/ui-system';
import type { PrintLayout } from '@react-sheets/pro-features';

export interface PrintPanelProps {
  onPrint: (layout: PrintLayout) => void;
  onExportPdf: (layout: PrintLayout) => void;
  onClose?: () => void;
}

export function PrintPanel({ onPrint, onExportPdf, onClose }: PrintPanelProps) {
  const [paper, setPaper] = useState<PrintLayout['paper']>('A4');
  const [orientation, setOrientation] = useState<PrintLayout['orientation']>('portrait');
  const [marginType, setMarginType] = useState<'normal' | 'narrow' | 'wide'>('normal');

  const getLayout = (): PrintLayout => {
    const margins = {
      normal: { top: 20, right: 20, bottom: 20, left: 20 },
      narrow: { top: 10, right: 10, bottom: 10, left: 10 },
      wide: { top: 30, right: 30, bottom: 30, left: 30 },
    }[marginType];

    return {
      paper,
      orientation,
      margin: margins,
    };
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Print & PDF Export</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Paper Size
            </Text>
            <Select
              value={paper}
              onChange={(e) => setPaper(e.target.value as PrintLayout['paper'])}
              sizeVariant="sm"
            >
              <option value="A4">A4 (210 × 297 mm)</option>
              <option value="Letter">Letter (8.5 × 11 in)</option>
              <option value="Legal">Legal (8.5 × 14 in)</option>
            </Select>
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Page Orientation
            </Text>
            <Select
              value={orientation}
              onChange={(e) => setOrientation(e.target.value as PrintLayout['orientation'])}
              sizeVariant="sm"
            >
              <option value="portrait">Portrait</option>
              <option value="landscape">Landscape</option>
            </Select>
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Margins
            </Text>
            <Select
              value={marginType}
              onChange={(e) => setMarginType(e.target.value as 'normal' | 'narrow' | 'wide')}
              sizeVariant="sm"
            >
              <option value="normal">Normal (0.75 in)</option>
              <option value="narrow">Narrow (0.25 in)</option>
              <option value="wide">Wide (1.0 in)</option>
            </Select>
          </div>

          <div className="rounded-lg border border-blue-100 bg-blue-50/60 p-3 text-xs text-blue-800">
            <div className="font-semibold">Print Output Layout</div>
            <div className="mt-0.5 text-[11px] text-blue-600">
              Automatic grid pagination with header row repeating is active.
            </div>
          </div>

          <div className="flex flex-col gap-2">
            <Button
              variant="primary"
              size="sm"
              icon="printer"
              onClick={() => onPrint(getLayout())}
            >
              Print Preview & Output
            </Button>
            <Button
              variant="outline"
              size="sm"
              icon="download"
              onClick={() => onExportPdf(getLayout())}
            >
              Print / Save as PDF
            </Button>
          </div>
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
