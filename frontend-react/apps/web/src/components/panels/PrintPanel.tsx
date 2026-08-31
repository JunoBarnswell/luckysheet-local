import React, { useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PrintLayout } from '@react-sheets/spreadsheet-app';

export interface PrintPanelProps {
  onPrint: (layout: PrintLayout, scope: PrintScope) => void;
  onExportPdf: (layout: PrintLayout, scope: PrintScope) => Promise<void>;
  pageCount?: number;
  onClose?: () => void;
  initialLayout?: PrintLayout;
}

export type PrintScope = 'saved-area' | 'selection' | 'active-sheet';

export function PrintPanel({ onPrint, onExportPdf, pageCount = 0, onClose, initialLayout }: PrintPanelProps) {
  const [paper, setPaper] = useState<PrintLayout['paper']>(initialLayout?.paper ?? 'A4');
  const [orientation, setOrientation] = useState<PrintLayout['orientation']>(initialLayout?.orientation ?? 'portrait');
  const [marginType, setMarginType] = useState<'normal' | 'narrow' | 'wide'>('normal');
  const [scope, setScope] = useState<PrintScope>('saved-area');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [headerText, setHeaderText] = useState(initialLayout?.headerText ?? '');
  const [footerText, setFooterText] = useState(initialLayout?.footerText ?? '');

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
      ...(headerText ? { headerText } : {}),
      ...(footerText ? { footerText } : {}),
    };
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Print & PDF Export</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Paper Size
            </Text>
            <Select
              value={paper}
              onChange={(e) => setPaper(e.target.value as PrintLayout['paper'])}
              sizeVariant="sm"
              options={[{ value: 'A4', label: 'A4 (210 × 297 mm)' }, { value: 'Letter', label: 'Letter (8.5 × 11 in)' }, { value: 'Legal', label: 'Legal (8.5 × 14 in)' }]}
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Header</Text>
            <TextInput aria-label="Print header" value={headerText} onChange={(event) => setHeaderText(event.currentTarget.value)} />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Footer</Text>
            <TextInput aria-label="Print footer" value={footerText} onChange={(event) => setFooterText(event.currentTarget.value)} />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Page Orientation
            </Text>
            <Select
              value={orientation}
              onChange={(e) => setOrientation(e.target.value as PrintLayout['orientation'])}
              sizeVariant="sm"
              options={[{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }]}
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Margins
            </Text>
            <Select
              value={marginType}
              onChange={(e) => setMarginType(e.target.value as 'normal' | 'narrow' | 'wide')}
              sizeVariant="sm"
              options={[{ value: 'normal', label: 'Normal (0.75 in)' }, { value: 'narrow', label: 'Narrow (0.25 in)' }, { value: 'wide', label: 'Wide (1.0 in)' }]}
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Print scope</Text>
            <Select
              value={scope}
              onChange={(event) => setScope(event.target.value as PrintScope)}
              sizeVariant="sm"
              options={[
                { value: 'saved-area', label: 'Saved print area (or used range)' },
                { value: 'selection', label: 'Current selection' },
                { value: 'active-sheet', label: 'Entire active sheet used range' },
              ]}
            />
          </Box>

          <Box className="rounded-lg border border-blue-100 bg-blue-50/60 p-3 text-xs text-blue-800">
            <Text size="sm" weight="semibold">Print Output Layout</Text>
            <Text size="xs" className="mt-0.5 text-blue-600">
              {pageCount > 0
                ? `${pageCount} page(s) will be generated from the current print area.`
                : 'Automatic grid pagination with header row repeating is active.'}
            </Text>
          </Box>

          <Stack gap="sm">
            {error ? <Text size="xs" tone="danger">{error}</Text> : null}
            <Button
              variant="primary"
              size="sm"
              icon="printer"
              disabled={busy}
              onClick={() => {
                setError(null);
                try { onPrint(getLayout(), scope); }
                catch (cause) { setError(cause instanceof Error ? cause.message : 'Print preview failed'); }
              }}
            >
              Print Preview & Output
            </Button>
            <Button
              variant="outline"
              size="sm"
              icon="download"
              disabled={busy}
              onClick={() => {
                setBusy(true);
                setError(null);
                void onExportPdf(getLayout(), scope).catch((cause) => setError(cause instanceof Error ? cause.message : 'PDF export failed')).finally(() => setBusy(false));
              }}
            >
              Print / Save as PDF
            </Button>
          </Stack>
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
