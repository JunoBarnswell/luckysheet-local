import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Box,
  Button,
  ContextMenu,
  Icon,
  Inline,
  Panel,
  PanelHeader,
  PanelTitle,
  Stack,
  StatePanel,
  Text,
  type ContextMenuItem,
} from '@react-sheets/ui-system';
import {
  CanvasRenderSurface,
  SheetSkeleton,
  type CanvasRenderEngine,
  type CellRenderData,
  type Rect,
} from '@react-sheets/render-engine';
import {
  drawChartOnCanvas,
  drawShapeOnCanvas,
  drawSparklineOnCanvas,
} from '@react-sheets/pro-features';
import type { ChartModel, ShapeModel, SparklineModel } from '@react-sheets/core-model';
import type { SheetView, WorkspacePhase } from '../state/workspace';
import { CellEditor } from './CellEditor';

export interface SheetCanvasProps {
  sheet: SheetView;
  sheetId: string;
  selectedCell: string;
  formulaDraft: string;
  phase: WorkspacePhase;
  zoom: number;
  charts?: ChartModel[];
  shapes?: ShapeModel[];
  sparklines?: SparklineModel[];
  onSelectCell: (address: string) => void;
  onCommitCell: (value: string) => void;
  onMoveCell: (address: string, direction: 'down' | 'left' | 'right' | 'up') => void;
  onAction: (action: string, payload?: unknown) => void;
  onRetry: () => void;
  onCreateSheet: () => void;
}

function parseAddress(address: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
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

export function SheetCanvas({
  sheet,
  sheetId,
  selectedCell,
  formulaDraft,
  phase,
  zoom,
  charts = [],
  shapes = [],
  sparklines = [],
  onSelectCell,
  onCommitCell,
  onMoveCell,
  onAction,
  onRetry,
  onCreateSheet,
}: SheetCanvasProps) {
  const engineRef = useRef<CanvasRenderEngine | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [editing, setEditing] = useState(false);
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; open: boolean }>({
    x: 0,
    y: 0,
    open: false,
  });

  const skeleton = useMemo(
    () =>
      new SheetSkeleton({
        rowCount: Math.max(sheet.rows.length, 100),
        columnCount: Math.max(sheet.columns.length, 26),
        defaultRowHeight: 28,
        defaultColumnWidth: 110,
      }),
    [sheet.columns.length, sheet.rows.length],
  );

  const cellsMap = useMemo(() => {
    const map = new Map<string, CellRenderData>();
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const pos = parseAddress(cell.address);
        if (!pos) continue;

        const isHeader = row.rowNumber === 1;
        map.set(`${pos.row}:${pos.column}`, {
          value: cell.value,
          formula: cell.formula,
          displayValue: cell.displayValue ?? cell.value,
          style: isHeader
            ? {
                background: '#f8fafc',
                textColor: '#334155',
                bold: true,
                fontSize: 12,
                fontFamily: 'Inter, sans-serif',
                horizontalAlignment: 'left',
                verticalAlignment: 'middle',
              }
            : cell.tone === 'accent'
              ? {
                  background: '#f0fdf4',
                  textColor: '#15803d',
                  bold: true,
                  fontSize: 12,
                  fontFamily: 'Inter, sans-serif',
                }
              : cell.tone === 'total'
                ? {
                    background: '#fef2f2',
                    textColor: '#b91c1c',
                    bold: true,
                    fontSize: 12,
                    fontFamily: 'Inter, sans-serif',
                  }
                : {
                    background: '#ffffff',
                    textColor: '#1e293b',
                    fontSize: 12,
                    fontFamily: 'Inter, sans-serif',
                  },
        });
      }
    }
    return map;
  }, [sheet]);

  // Update canvas engine on cells/skeleton change
  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setSkeleton(skeleton);
    engine.setCells(cellsMap);
    engine.invalidate();
    engine.requestRender();
  }, [cellsMap, skeleton]);

  // Draw floating items (Charts, Shapes, Sparklines) on overlay/extension canvas layer
  useEffect(() => {
    const engine = engineRef.current;
    if (!engine || phase !== 'ready') return;
    const canvas = engine.getCanvas('overlay');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // Draw charts
    for (const chart of charts) {
      const categories = ['Q1', 'Q2', 'Q3', 'Q4'];
      const series = [
        { name: 'Target', values: [100, 130, 150, 180], color: '#2563eb' },
        { name: 'Actual', values: [95, 140, 145, 190], color: '#10b981' },
      ];
      drawChartOnCanvas({ context: ctx, chart, categories, series });
    }

    // Draw shapes
    for (const shape of shapes) {
      drawShapeOnCanvas({ context: ctx, shape });
    }

    // Draw sparklines
    for (const spark of sparklines) {
      const cellRect = skeleton.getCellRect(spark.anchor.row, spark.anchor.column);
      if (cellRect) {
        drawSparklineOnCanvas({
          context: ctx,
          sparkline: spark,
          values: [10, 14, 8, 16, 22, 19, 25],
          rect: cellRect,
        });
      }
    }
  }, [charts, shapes, sparklines, phase, skeleton]);

  const handleReady = useCallback(
    (engine: CanvasRenderEngine) => {
      engineRef.current = engine;
      engine.setSkeleton(skeleton);
      engine.setCells(cellsMap);
      engine.render();
    },
    [cellsMap, skeleton],
  );

  const handlePointerDown = useCallback(
    (event: React.PointerEvent<HTMLDivElement>) => {
      if (event.button === 2) return; // Right click handled by context menu
      const engine = engineRef.current;
      if (!engine || phase !== 'ready') return;

      const bounds = event.currentTarget.getBoundingClientRect();
      const viewport = engine.viewport.getSnapshot();
      const point = {
        x: event.clientX - bounds.left + viewport.scrollX,
        y: event.clientY - bounds.top + viewport.scrollY,
      };
      const cell = engine.skeleton.getCellAtPoint(point);
      if (cell) {
        onSelectCell(cellAddress(cell.row, cell.column));
        if (editing) setEditing(false);
      }
    },
    [editing, onSelectCell, phase],
  );

  const handleDoubleClick = useCallback(() => {
    if (phase === 'ready') {
      setEditing(true);
    }
  }, [phase]);

  const handleWheel = useCallback(
    (event: React.WheelEvent<HTMLDivElement>) => {
      const engine = engineRef.current;
      if (!engine || phase !== 'ready') return;
      event.preventDefault();
      engine.scrollBy(event.deltaX, event.deltaY);
      engine.requestRender();
    },
    [phase],
  );

  const handleKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (phase !== 'ready' || editing) return;

    if (event.key === 'ArrowUp') {
      event.preventDefault();
      onMoveCell(selectedCell, 'up');
    } else if (event.key === 'ArrowDown') {
      event.preventDefault();
      onMoveCell(selectedCell, 'down');
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      onMoveCell(selectedCell, 'left');
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      onMoveCell(selectedCell, 'right');
    } else if (event.key === 'Tab') {
      event.preventDefault();
      onMoveCell(selectedCell, event.shiftKey ? 'left' : 'right');
    } else if (event.key === 'Enter') {
      event.preventDefault();
      onMoveCell(selectedCell, event.shiftKey ? 'up' : 'down');
    } else if (event.key === 'F2') {
      event.preventDefault();
      setEditing(true);
    } else if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      onAction('clear-range');
    } else if (event.ctrlKey || event.metaKey) {
      if (event.key === 'z') {
        event.preventDefault();
        onAction('undo');
      } else if (event.key === 'y') {
        event.preventDefault();
        onAction('redo');
      } else if (event.key === 'c') {
        event.preventDefault();
        onAction('copy');
      } else if (event.key === 'x') {
        event.preventDefault();
        onAction('cut');
      } else if (event.key === 'v') {
        event.preventDefault();
        onAction('paste');
      } else if (event.key === 'b') {
        event.preventDefault();
        onAction('bold');
      } else if (event.key === 'i') {
        event.preventDefault();
        onAction('italic');
      } else if (event.key === 'u') {
        event.preventDefault();
        onAction('underline');
      }
    } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // Start typing directly in cell
      setEditing(true);
    }
  };

  const handleContextMenu = (event: React.MouseEvent) => {
    event.preventDefault();
    if (phase !== 'ready') return;
    setContextMenu({
      x: event.clientX,
      y: event.clientY,
      open: true,
    });
  };

  const selected = parseAddress(selectedCell);
  const selectedRect = selected ? skeleton.getCellRect(selected.row, selected.column) : null;
  const isSheetEmpty = phase === 'empty' || sheet.isEmpty;

  const contextMenuItems: ContextMenuItem[] = [
    { id: 'cut', label: 'Cut', icon: 'scissors', shortcut: 'Ctrl+X', onSelect: () => onAction('cut') },
    { id: 'copy', label: 'Copy', icon: 'copy', shortcut: 'Ctrl+C', onSelect: () => onAction('copy') },
    { id: 'paste', label: 'Paste', icon: 'clipboard', shortcut: 'Ctrl+V', onSelect: () => onAction('paste') },
    { id: 'sep-1', label: '', separator: true },
    { id: 'insert-row', label: 'Insert Row Above', icon: 'rows', onSelect: () => onAction('insert-row') },
    { id: 'insert-col', label: 'Insert Column Left', icon: 'columns', onSelect: () => onAction('insert-column') },
    { id: 'delete-row', label: 'Delete Row', icon: 'trash', onSelect: () => onAction('delete-row'), danger: true },
    { id: 'delete-col', label: 'Delete Column', icon: 'trash', onSelect: () => onAction('delete-column'), danger: true },
    { id: 'sep-2', label: '', separator: true },
    { id: 'clear', label: 'Clear Contents', icon: 'trash', shortcut: 'Del', onSelect: () => onAction('clear-range') },
    { id: 'conditional-format', label: 'Conditional Format', icon: 'sparkles', onSelect: () => onAction('open-conditional-format') },
    { id: 'data-validation', label: 'Data Validation', icon: 'check-circle', onSelect: () => onAction('open-data-validation') },
  ];

  return (
    <Panel className="flex h-full min-h-0 flex-col overflow-hidden rounded-none border-0 bg-white shadow-none">
      <PanelHeader className="h-10 shrink-0 border-b border-slate-200 px-4">
        <Inline gap="sm">
          <Box className="flex h-6 w-6 items-center justify-center rounded bg-blue-50 text-blue-600">
            <Icon name="grid" size="xs" />
          </Box>
          <PanelTitle as="h2" size="sm">
            {sheet.name}
          </PanelTitle>
          <span className="rounded bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500">
            {sheet.columns.length} × {Math.max(sheet.rows.length, 100)}
          </span>
        </Inline>
        <Inline gap="sm" className="ml-auto">
          <Text size="xs" tone="muted">
            {zoom}%
          </Text>
        </Inline>
      </PanelHeader>

      <div
        ref={containerRef}
        tabIndex={0}
        onKeyDown={handleKeyDown}
        onPointerDown={handlePointerDown}
        onDoubleClick={handleDoubleClick}
        onWheel={handleWheel}
        onContextMenu={handleContextMenu}
        className="relative min-h-0 flex-1 overflow-hidden bg-slate-100/60 p-2 outline-hidden focus:ring-1 focus:ring-blue-400"
        role="grid"
        aria-label={`${sheet.name} spreadsheet surface`}
      >
        {phase === 'loading' ? (
          <StatePanel kind="loading" description="Preparing worksheet canvas..." />
        ) : null}
        {phase === 'error' ? (
          <StatePanel actionLabel="Try again" description="Worksheet failed to render." kind="error" onAction={onRetry} />
        ) : null}
        {isSheetEmpty ? (
          <StatePanel actionLabel="Create a sheet" description="No worksheet data found." kind="empty" onAction={onCreateSheet} />
        ) : null}

        {phase === 'ready' && !isSheetEmpty ? (
          <Box className="relative h-full min-h-[420px] overflow-hidden rounded-lg border border-slate-300 bg-white shadow-xs">
            <CanvasRenderSurface ref={engineRef} onReady={handleReady} className="absolute inset-0" />

            {/* In-Cell Editor Overlay */}
            {editing && selectedRect ? (
              <CellEditor
                initialValue={formulaDraft}
                rect={selectedRect}
                onCommit={(val) => {
                  setEditing(false);
                  onCommitCell(val);
                }}
                onCancel={() => setEditing(false)}
                onNavigate={(dir) => onMoveCell(selectedCell, dir)}
              />
            ) : null}

            {/* Selection Outline with Fill Handle */}
            {!editing && selectedRect ? (
              <div
                aria-label={`Selected cell ${selectedCell}`}
                className="pointer-events-none absolute border-2 border-blue-600 bg-blue-500/10"
                style={{
                  left: selectedRect.x,
                  top: selectedRect.y,
                  width: selectedRect.width,
                  height: selectedRect.height,
                }}
              >
                {/* Fill Handle Corner Box */}
                <div className="absolute -bottom-1.5 -right-1.5 h-2.5 w-2.5 cursor-crosshair rounded-xs border border-white bg-blue-600 pointer-events-auto" />
              </div>
            ) : null}
          </Box>
        ) : null}
      </div>

      {/* Context Menu */}
      <ContextMenu
        x={contextMenu.x}
        y={contextMenu.y}
        open={contextMenu.open}
        items={contextMenuItems}
        onClose={() => setContextMenu((prev) => ({ ...prev, open: false }))}
      />
    </Panel>
  );
}
