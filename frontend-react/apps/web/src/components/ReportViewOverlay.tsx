import { useMemo } from 'react';
import { Box, Text } from '@react-sheets/ui-system';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import { buildReportProjection, type CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import type { CanvasRenderEngine } from '@react-sheets/render-engine';

export interface ReportViewOverlayProps {
  engine: CanvasRenderEngine | null;
  sheet: CanvasSheetSnapshot;
  tables: readonly WorkbookTableModel[];
  sourceSheets?: readonly CanvasSheetSnapshot[];
  scrollTick: number;
}

export function ReportViewOverlay({ engine, sheet, tables, sourceSheets = [], scrollTick }: ReportViewOverlayProps) {
  void scrollTick;
  const projection = useMemo(() => buildReportProjection(sheet, tables, sourceSheets), [sheet, tables, sourceSheets]);
  if (!engine || !sheet.reportSheet) return null;
  if (projection.status === 'error') return <Box role="status" aria-live="polite" className="pointer-events-none absolute left-2 top-2 z-20 max-w-[32rem] rounded border border-red-200 bg-red-50 px-2 py-1 text-[11px] text-red-700 shadow-sm">ReportSheet unavailable: {projection.error}</Box>;
  return (
    <Box className="pointer-events-none absolute inset-0 z-10 overflow-hidden">
      {projection.cells.map((cell, index) => {
        const rect = engine.contentRangeToScreenRects({ startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column })[0];
        if (!rect) return null;
        return <Box key={`${cell.row}:${cell.column}:${index}`} className="absolute overflow-hidden border border-indigo-200/70 bg-indigo-50/80 px-1" style={{ left: rect.x, top: rect.y, width: rect.width, height: rect.height }}><Text size="xs" className="truncate leading-5 text-indigo-900">{cell.value}</Text></Box>;
      })}
      {projection.renderMode === 'paginated' ? <Box className="absolute right-2 top-2 rounded border border-slate-200 bg-white/90 px-2 py-1 shadow-sm"><Text size="xs" tone="muted">Page 1 / {Math.max(1, projection.pageCount)}</Text></Box> : null}
    </Box>
  );
}
