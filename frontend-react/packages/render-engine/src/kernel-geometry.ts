import { kernelInvoke } from '@react-sheets/kernel-client';
import type { CellAddress, CellRange, PaneLayout, PaneMap, Point, Rect, RenderPane, ViewportSnapshot } from './types';
import type { SheetSkeleton } from './sheet-skeleton';

interface GeometryResponse { paneMap: { panes: Array<{ id: string; screenRect: Rect; contentOrigin: Point; visibleRange: CellRange | null }> } }
interface HitTestResponse { address: CellAddress | null; pane: string | null }
interface HeaderRectResponse { rect: Rect; axis: 'corner' | 'row' | 'column'; index?: number }

export function geometryRequest(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null): unknown {
  if (!sheetId) throw new Error('Geometry requires a canonical sheetId');
  return {
    sheetId,
    rowCount: skeleton.rowCount,
    columnCount: skeleton.columnCount,
    defaultRowHeightPx: skeleton.defaultRowHeight,
    defaultColumnWidthPx: skeleton.defaultColumnWidth,
    rowHeightsPx: Object.fromEntries(skeleton.getRowHeightOverrides()),
    columnWidthsPx: Object.fromEntries(skeleton.getColumnWidthOverrides()),
    hiddenRows: skeleton.hiddenRows,
    hiddenColumns: skeleton.hiddenColumns,
    zoom: skeleton.zoom,
    viewport,
    pane,
    headerOffset,
  };
}

export function computePaneMapFromKernel(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null): PaneMap {
  const response = kernelInvoke<GeometryResponse>('geometry.computePaneMap', geometryRequest(sheetId, skeleton, viewport, pane, headerOffset));
  const panes: RenderPane[] = response.paneMap.panes.map((entry) => ({ ...entry, id: entry.id as RenderPane['id'] }));
  return createPaneMap(panes);
}

export function hitTestFromKernel(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null, point: Point): HitTestResponse {
  return kernelInvoke<HitTestResponse>('geometry.hitTest', { request: geometryRequest(sheetId, skeleton, viewport, pane, headerOffset), point });
}

export function cellRectFromKernel(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null, address: CellAddress): Rect {
  return kernelInvoke<Rect>('geometry.cellRect', { request: geometryRequest(sheetId, skeleton, viewport, pane, headerOffset), address: { ...address, sheetId } });
}

export function headerRectFromKernel(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null, axis: 'corner' | 'row' | 'column', index?: number): HeaderRectResponse {
  return kernelInvoke<HeaderRectResponse>('geometry.headerRect', { request: geometryRequest(sheetId, skeleton, viewport, pane, headerOffset), axis, ...(index === undefined ? {} : { index }) });
}

export function headerIndexAtKernel(sheetId: string, skeleton: SheetSkeleton, viewport: ViewportSnapshot, pane: PaneLayout | null, headerOffset: Point | null, axis: 'row' | 'column', point: Point): number | null {
  const map = computePaneMapFromKernel(sheetId, skeleton, viewport, pane, headerOffset);
  const candidates = new Set<number>();
  for (const entry of map.panes) {
    const range = entry.visibleRange;
    if (!range) continue;
    const start = axis === 'row' ? range.startRow : range.startColumn;
    const end = axis === 'row' ? range.endRow : range.endColumn;
    for (let index = start; index <= end; index += 1) candidates.add(index);
  }
  for (const index of candidates) {
    const rect = headerRectFromKernel(sheetId, skeleton, viewport, pane, headerOffset, axis, index).rect;
    if (point.x >= rect.x && point.x < rect.x + rect.width && point.y >= rect.y && point.y < rect.y + rect.height) return index;
  }
  return null;
}

function createPaneMap(panes: RenderPane[]): PaneMap {
  return {
    panes,
    paneAtLocalPoint(point) {
      return panes.find((candidate) => point.x >= candidate.screenRect.x && point.x < candidate.screenRect.x + candidate.screenRect.width
        && point.y >= candidate.screenRect.y && point.y < candidate.screenRect.y + candidate.screenRect.height) ?? null;
    },
    paneForCell(cell) {
      return panes.find((candidate) => candidate.visibleRange
        && cell.row >= candidate.visibleRange.startRow && cell.row <= candidate.visibleRange.endRow
        && cell.column >= candidate.visibleRange.startColumn && cell.column <= candidate.visibleRange.endColumn) ?? null;
    },
  };
}
