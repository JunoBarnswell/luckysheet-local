import { useEffect, useMemo, useState } from 'react';
import type { PrintProjection } from '@react-sheets/spreadsheet-app';
import { AssetIcon, Box, Button, Dialog, Inline, Stack, StatePanel, Text } from '@react-sheets/ui-system';

export interface PrintPreviewDialogProps {
  open: boolean;
  onClose: () => void;
  projections: readonly PrintProjection[];
  headerText?: string;
  footerText?: string;
}

function drawingLabel(projection: PrintProjection['drawings'][number]): string {
  const payload = projection.payload;
  if (payload.kind === 'shape') return payload.text?.trim() || `Shape · ${payload.type}`;
  if (payload.kind === 'textbox') return payload.text.trim() || 'Text box';
  if (payload.kind === 'image') return payload.altText?.trim() || `Image · ${payload.asset.assetId}`;
  if (payload.kind === 'chart') return payload.elements.title?.trim() || `Chart · ${payload.chartType}`;
  return payload.kind;
}

function PrintChart({ chart }: { chart: NonNullable<PrintProjection['drawings'][number]['chart']> }) {
  const values = chart.series.flatMap((series) => series.values.filter((value): value is number => value !== null && Number.isFinite(value)));
  const maximum = Math.max(1, ...values.map((value) => Math.abs(value)));
  return (
    <Box className="flex size-full items-end gap-px border border-slate-300 bg-white p-1" aria-label="Chart preview">
      {chart.series.flatMap((series) => series.values.map((value, index) => (
        <Box
          key={`${series.id}:${index}`}
          className="min-w-px flex-1"
          style={{ height: `${Math.max(2, Math.abs(value ?? 0) / maximum * 92)}%`, background: series.color ?? '#2563eb' }}
        />
      )))}
    </Box>
  );
}

function PrintDrawing({ entry }: { entry: PrintProjection['drawings'][number] }) {
  if (entry.image?.url) return <AssetIcon alt={drawingLabel(entry)} className="!size-full object-contain" size="xl" src={entry.image.url} />;
  if (entry.chart) return <PrintChart chart={entry.chart} />;
  const payload = entry.payload;
  const ellipse = payload.kind === 'shape' && payload.type === 'ellipse';
  const fill = payload.kind === 'shape' ? payload.fill : '#eff6ff';
  const borderColor = payload.kind === 'shape' ? payload.stroke : '#60a5fa';
  return (
    <Box
      className="flex size-full items-center justify-center overflow-hidden border-2 p-1 text-center text-[9px] text-slate-800"
      style={{ background: fill, borderColor, borderRadius: ellipse ? '9999px' : undefined }}
    >
      {drawingLabel(entry)}
    </Box>
  );
}

export function PrintPreviewDialog({ open, onClose, projections, headerText, footerText }: PrintPreviewDialogProps) {
  const [pageIndex, setPageIndex] = useState(0);
  useEffect(() => setPageIndex(0), [open, projections]);
  const page = projections[pageIndex];
  const preview = useMemo(() => {
    if (!page) return null;
    const width = Math.max(1, page.contentWidthPx * page.scaleX);
    const height = Math.max(1, page.contentHeightPx * page.scaleY);
    const scale = Math.min(1, 720 / width, 480 / height);
    return { width, height, scale };
  }, [page]);

  return (
    <Dialog open={open} onClose={onClose} title="Print preview" maxWidth="xl">
      <Stack gap="md">
        {!page || !preview ? (
          <StatePanel kind="error" title="Print projection unavailable" description="The workbook did not produce a printable page projection." />
        ) : (
          <Box className="max-h-[58vh] overflow-auto rounded-xl border border-slate-200 bg-slate-100 p-4">
            <Box
              className="relative origin-top-left bg-white shadow-lg"
              style={{ width: preview.width, height: preview.height, transform: `scale(${preview.scale})`, marginBottom: preview.height * (preview.scale - 1) }}
              data-print-projection-page={page.page.pageIndex}
            >
              {headerText ? <Text size="xs" className="absolute left-2 right-2 top-1 truncate text-center text-[9px] text-slate-600">{headerText}</Text> : null}
              {page.cells.map((cell) => (
                <Box
                  key={`${cell.row}:${cell.column}`}
                  className="absolute overflow-hidden border border-slate-300 px-1 text-[10px] leading-tight text-slate-800"
                  style={{
                    left: cell.columnOffsetPx * page.scaleX,
                    top: cell.rowOffsetPx * page.scaleY,
                    width: cell.widthPx * page.scaleX,
                    height: cell.heightPx * page.scaleY,
                    background: cell.style?.background ?? '#ffffff',
                    color: cell.style?.textColor ?? '#1f2937',
                    fontFamily: cell.style?.fontFamily,
                    fontSize: cell.style?.fontSizePx,
                    fontWeight: cell.style?.bold ? 700 : 400,
                    fontStyle: cell.style?.italic ? 'italic' : 'normal',
                  }}
                >
                  {cell.image?.url
                    ? <AssetIcon alt={cell.displayValue} className="!size-full object-contain" size="xl" src={cell.image.url} />
                    : cell.displayValue}
                </Box>
              ))}
              {page.drawings.map((entry) => (
                <Box
                  key={entry.drawing.id}
                  className="absolute overflow-hidden"
                  style={{ left: entry.xPx * page.scaleX, top: entry.yPx * page.scaleY, width: entry.widthPx * page.scaleX, height: entry.heightPx * page.scaleY }}
                >
                  <PrintDrawing entry={entry} />
                </Box>
              ))}
              {footerText ? <Text size="xs" className="absolute bottom-1 left-2 right-2 truncate text-center text-[9px] text-slate-600">{footerText}</Text> : null}
            </Box>
          </Box>
        )}

        <Inline gap="sm" className="items-center justify-between">
          <Button disabled={pageIndex <= 0} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.max(0, index - 1))}>Previous page</Button>
          <Text size="xs" tone="muted">{page ? `Page ${pageIndex + 1} of ${projections.length}` : 'No page'}</Text>
          <Button disabled={pageIndex >= projections.length - 1} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.min(projections.length - 1, index + 1))}>Next page</Button>
        </Inline>
        <Inline gap="sm" className="justify-end"><Button variant="primary" onClick={onClose}>Close</Button></Inline>
      </Stack>
    </Dialog>
  );
}
