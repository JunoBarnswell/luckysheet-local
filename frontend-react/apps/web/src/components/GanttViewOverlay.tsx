import { useMemo } from 'react';
import { Box, Stack, Text } from '@react-sheets/ui-system';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import { buildGanttProjection, type CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import type { CanvasRenderEngine } from '@react-sheets/render-engine';

export interface GanttViewOverlayProps {
  engine: CanvasRenderEngine | null;
  sheet: CanvasSheetSnapshot;
  tables: readonly WorkbookTableModel[];
  scrollTick: number;
}

const DAY_MS = 24 * 60 * 60 * 1000;

function unitMs(unit: 'day' | 'week' | 'month' | 'quarter'): number {
  if (unit === 'day') return DAY_MS;
  if (unit === 'week') return 7 * DAY_MS;
  if (unit === 'month') return 30 * DAY_MS;
  return 90 * DAY_MS;
}

export function GanttViewOverlay({ engine, sheet, tables, scrollTick }: GanttViewOverlayProps) {
  void scrollTick;
  const projection = useMemo(() => buildGanttProjection(sheet, tables), [sheet, tables]);
  const definition = sheet.ganttSheet;
  if (!engine || !definition) return null;
  const timelineColumn = Math.min(Math.max(7, (tables.find((table) => table.id === definition.viewId)?.fields.length ?? 1) + 1), Math.max(0, sheet.columnCount - 1));
  const anchor = engine.contentRangeToScreenRects({ startRow: 0, endRow: 0, startColumn: timelineColumn, endColumn: timelineColumn })[0];
  if (!anchor) return null;
  if (projection.status === 'error') {
    return <Box role="status" aria-live="polite" className="pointer-events-none absolute left-2 top-2 z-20 max-w-[32rem] rounded border border-red-200 bg-red-50 px-2 py-1 text-[11px] text-red-700 shadow-sm">GanttSheet unavailable: {projection.error}</Box>;
  }
  const scale = Math.max(1, projection.timelineEndMs - projection.timelineStartMs);
  const pixelsPerMs = Math.max(0.00001, Math.min(0.08, 720 / scale));
  const rowRects = new Map(projection.tasks.map((task) => [task.id, engine.contentRangeToScreenRects({ startRow: task.row, endRow: task.row, startColumn: timelineColumn, endColumn: timelineColumn })[0]]));
  const barPosition = (task: typeof projection.tasks[number]) => {
    const rowRect = rowRects.get(task.id);
    if (!rowRect) return null;
    const x = anchor.x + (task.startMs - projection.timelineStartMs) * pixelsPerMs;
    const width = Math.max(4, (task.endMs - task.startMs) * pixelsPerMs);
    return { x, y: rowRect.y + 3, width, height: Math.max(8, rowRect.height - 6), rowRect };
  };
  const bars = projection.tasks.map((task) => ({ task, position: barPosition(task) })).filter((item): item is { task: typeof projection.tasks[number]; position: NonNullable<ReturnType<typeof barPosition>> } => Boolean(item.position));
  const byId = new Map(bars.map(({ task, position }) => [task.id, { task, position }]));
  const dependencyLines = bars.flatMap(({ task, position }) => task.dependencies.flatMap((dependency) => {
    const source = byId.get(dependency);
    if (!source) return [];
    const x1 = source.position.x + source.position.width;
    const x2 = position.x;
    const y = position.y + position.height / 2;
    const sourceY = source.position.y + source.position.height / 2;
    return [{ id: `${dependency}:${task.id}`, left: Math.min(x1, x2), top: Math.min(sourceY, y), width: Math.max(1, Math.abs(x2 - x1)), height: Math.max(1, Math.abs(y - sourceY)), horizontalTop: y === sourceY ? 0 : undefined }];
  }));
  const headerLabels = [0, 1, 2, 3, 4].map((index) => {
    const date = new Date(projection.timelineStartMs + index * unitMs(projection.unit));
    return { label: date.toISOString().slice(0, 10), left: anchor.x + index * unitMs(projection.unit) * pixelsPerMs };
  });
  return (
    <Box className="pointer-events-none absolute inset-0 z-10 overflow-hidden">
      <Stack gap="none">
        {headerLabels.map((entry) => <Text key={entry.label} size="xs" tone="muted" className="absolute top-0 whitespace-nowrap" style={{ left: entry.left + 3 }}>{entry.label}</Text>)}
      </Stack>
      {bars.map(({ task, position }) => (
        <Box key={task.id} className="absolute overflow-hidden rounded bg-blue-500/80 shadow-sm" style={{ left: position.x, top: position.y, width: position.width, height: position.height }}>
          <Box className="h-full bg-emerald-400/80" style={{ width: `${task.progress}%` }} />
          <Text size="xs" className="absolute inset-y-0 left-1 truncate leading-5 text-white" style={{ maxWidth: Math.max(0, position.width - 4) }}>{`${'　'.repeat(task.level)}${task.title}`}</Text>
        </Box>
      ))}
      {dependencyLines.map((line) => (
        <Box key={line.id} className="absolute bg-slate-500/70" style={{ left: line.left, top: line.top, width: line.horizontalTop === 0 ? line.width : 1, height: line.horizontalTop === 0 ? Math.max(1, definition.dependencyStyle.width) : line.height, transform: line.horizontalTop === 0 ? undefined : `translateX(${line.width}px)`, transformOrigin: 'top left' }} />
      ))}
    </Box>
  );
}
