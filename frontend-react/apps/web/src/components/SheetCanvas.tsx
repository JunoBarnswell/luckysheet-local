import { useCallback, useEffect, useMemo, useRef } from 'react';
import { Box, Button, Icon, Inline, Panel, PanelHeader, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import { CanvasRenderSurface, SheetSkeleton, type CanvasRenderEngine, type CellRenderData } from '@react-sheets/render-engine';
import type { SheetView, WorkspacePhase } from '../state/workspace';

export interface SheetCanvasProps {
  onCreateSheet: () => void;
  onMoveCell: (address: string, direction: 'down' | 'left' | 'right' | 'up') => void;
  onRetry: () => void;
  onSelectCell: (address: string) => void;
  phase: WorkspacePhase;
  selectedCell: string;
  sheet: SheetView;
  zoom: number;
}

function parseAddress(address: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function toRenderCells(sheet: SheetView): Map<string, CellRenderData> {
  const cells = new Map<string, CellRenderData>();
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      const position = parseAddress(cell.address);
      if (!position || cell.value === '') continue;
      cells.set(`${position.row}:${position.column}`, {
        value: cell.value,
        displayValue: cell.value,
        style: row.rowNumber === 1
          ? { background: '#eef2ff', textColor: '#172033', font: '600 13px Inter, sans-serif' }
          : cell.tone === 'accent'
            ? { background: '#ecfdf5', textColor: '#047857', font: '600 13px Inter, sans-serif' }
            : { background: '#ffffff', textColor: '#172033', font: '13px Inter, sans-serif' },
      });
    }
  }
  return cells;
}

function cellAddress(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return `${label}${row + 1}`;
}

export function SheetCanvas({ onCreateSheet, onMoveCell, onRetry, onSelectCell, phase, selectedCell, sheet, zoom }: SheetCanvasProps) {
  const engineRef = useRef<CanvasRenderEngine | null>(null);
  const cells = useMemo(() => toRenderCells(sheet), [sheet]);
  const skeleton = useMemo(() => new SheetSkeleton({ rowCount: Math.max(sheet.rows.length, 24), columnCount: sheet.columns.length, defaultRowHeight: 32, defaultColumnWidth: 128 }), [sheet.columns.length, sheet.rows.length]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setSkeleton(skeleton);
    engine.setCells(cells);
    engine.invalidate();
    engine.requestRender();
  }, [cells, skeleton]);

  const handleReady = useCallback((engine: CanvasRenderEngine) => {
    engineRef.current = engine;
    engine.setSkeleton(skeleton);
    engine.setCells(cells);
    engine.render();
  }, [cells, skeleton]);

  const handlePointerDown = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const engine = engineRef.current;
    if (!engine || phase !== 'ready') return;
    const bounds = event.currentTarget.getBoundingClientRect();
    const viewport = engine.viewport.getSnapshot();
    const point = { x: event.clientX - bounds.left + viewport.scrollX, y: event.clientY - bounds.top + viewport.scrollY };
    const cell = engine.skeleton.getCellAtPoint(point);
    if (cell) onSelectCell(cellAddress(cell.row, cell.column));
  }, [onSelectCell, phase]);

  const handleWheel = useCallback((event: React.WheelEvent<HTMLDivElement>) => {
    const engine = engineRef.current;
    if (!engine || phase !== 'ready') return;
    event.preventDefault();
    engine.scrollBy(event.deltaX, event.deltaY);
    engine.requestRender();
  }, [phase]);

  const selected = parseAddress(selectedCell);
  const selectedRect = selected ? skeleton.getCellRect(selected.row, selected.column) : null;
  const isSheetEmpty = phase === 'empty' || sheet.isEmpty;

  return (
    <Panel className="flex h-full min-h-0 flex-col overflow-hidden rounded-none border-0 bg-white shadow-none">
      <PanelHeader className="h-12 shrink-0 px-5">
        <Inline gap="sm">
          <Box className="flex h-7 w-7 items-center justify-center rounded-lg bg-blue-50 text-accent"><Icon name="grid" size="sm" /></Box>
          <Stack gap="none"><PanelTitle as="h2" size="sm">{sheet.name}</PanelTitle><Text size="xs" tone="subtle">Canvas Scene · Layer · Viewport</Text></Stack>
        </Inline>
        <Inline gap="sm"><Text size="xs" tone="muted">{zoom}% view</Text><Button aria-label="Canvas options" disabled={phase !== 'ready'} icon="more-horizontal" iconOnly size="sm" variant="ghost" /></Inline>
      </PanelHeader>

      <Box className="relative min-h-0 flex-1 overflow-hidden bg-canvas/70 p-5" onPointerDown={handlePointerDown} onWheel={handleWheel} role="grid" aria-label={`${sheet.name} spreadsheet`}>
        {phase === 'loading' ? <StatePanel kind="loading" description="Preparing the workbook surface and active sheet." /> : null}
        {phase === 'error' ? <StatePanel actionLabel="Try again" description="The workbook surface could not be prepared." kind="error" onAction={onRetry} /> : null}
        {isSheetEmpty ? <StatePanel actionLabel="Create a sheet" description="There is no active worksheet content yet." kind="empty" onAction={onCreateSheet} /> : null}
        {phase === 'ready' && !isSheetEmpty ? (
          <Box className="relative h-full min-h-[420px] overflow-hidden rounded-xl border border-line bg-white shadow-panel">
            <CanvasRenderSurface ref={engineRef} onReady={handleReady} className="absolute inset-0" />
            {selectedRect ? <Box aria-label={`Selected cell ${selectedCell}`} className="pointer-events-none absolute border-2 border-accent" style={{ left: selectedRect.x, top: selectedRect.y, width: selectedRect.width, height: selectedRect.height }} /> : null}
          </Box>
        ) : null}
      </Box>
      <Inline gap="sm" className="h-9 shrink-0 border-t border-line/80 px-5"><Text size="xs" tone="subtle">RenderPlan driven Canvas</Text><Box className="h-1 w-1 rounded-full bg-slate-300" /><Text size="xs" tone="subtle">Selection and formula bar use Command state</Text></Inline>
    </Panel>
  );
}
