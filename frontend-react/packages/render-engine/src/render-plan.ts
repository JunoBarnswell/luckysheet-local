import { mergeCellRanges } from './dirty-ranges';
import { intersectRect, mergeRects, translateRect } from './geometry';
import { SheetSkeleton } from './sheet-skeleton';
import {
  DEFAULT_LAYER_DEFINITIONS,
  COL_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  type CellRange,
  type PaneLayout,
  type PaneMap,
  type LayerDefinition,
  type PaneId,
  type Point,
  type Rect,
  type RenderPane,
  type ViewportSnapshot,
} from './types';

export type LayerRenderMode = 'none' | 'full' | 'dirty';

export interface ScrollDeltaPlan {
  delta: Point;
  dx: number;
  dy: number;
  hasDelta: boolean;
  /**
   * Full screen-space rectangles of panes whose content origin changed.
   * Scrolling always clears and redraws these visible panes. Reusing pixels
   * from the same canvas is intentionally unsupported because repeated
   * overlapping copies accumulate stale glyphs during continuous scrolling.
   */
  redrawRects: Rect[];
}

export interface LayerRenderPlan {
  layerId: string;
  mode: LayerRenderMode;
  clearRects: Rect[];
  drawRects: Rect[];
}

export type RenderPlanReason = 'initial' | 'forced' | 'resize' | 'scroll-redraw' | 'mixed' | 'dirty' | 'idle';

export interface RenderPlan {
  viewport: ViewportSnapshot;
  visibleRange: CellRange | null;
  paneMap: PaneMap;
  /** 冻结窗格切分结果(无冻结时为单个 main 窗格) */
  panes: RenderPane[];
  dirtyRanges: CellRange[];
  dirtyRects: Rect[];
  scrollDelta: ScrollDeltaPlan;
  fullRedraw: boolean;
  reason: RenderPlanReason;
  layers: LayerRenderPlan[];
}

export interface RenderPlanInput {
  skeleton: SheetSkeleton;
  viewport: ViewportSnapshot;
  previousViewport?: ViewportSnapshot | null;
  dirtyRanges?: readonly CellRange[];
  forceFull?: boolean;
  /** 仅重绘 chrome 层(选区/hover/浮动选中态) */
  chromeDirty?: boolean;
  layers?: readonly LayerDefinition[];
  /** 文档窗格布局(null = 无窗格) */
  pane?: PaneLayout | null;
  /** 表头条占用偏移(有表头时为 {x:ROW_HEADER_WIDTH,y:COL_HEADER_HEIGHT}) */
  headerOffset?: Point | null;
}

/** 计算文档窗格的屏幕矩形、内容原点与可见范围。 */
export function computePaneMap(
  skeleton: SheetSkeleton,
  viewport: ViewportSnapshot,
  pane: PaneLayout | null,
  headerOffset: Point | null,
): PaneMap {
  const originX = headerOffset?.x ?? 0;
  const originY = headerOffset?.y ?? 0;
  const gridWidth = Math.max(0, viewport.width - originX);
  const gridHeight = Math.max(0, viewport.height - originY);

  if (!pane || pane.kind === 'none' || (pane.xSplit <= 0 && pane.ySplit <= 0)) {
    return createPaneMap([{
      id: 'main',
      screenRect: { x: originX, y: originY, width: gridWidth, height: gridHeight },
      contentOrigin: { x: viewport.scrollX, y: viewport.scrollY },
      visibleRange: skeleton.getVisibleRange({
        x: viewport.scrollX,
        y: viewport.scrollY,
        width: gridWidth,
        height: gridHeight,
      }),
    }]);
  }

  const frozen = pane.kind === 'frozen';
  const xSplit = frozen
    ? Math.max(0, Math.min(Math.trunc(pane.xSplit), skeleton.columnCount))
    : 0;
  const ySplit = frozen
    ? Math.max(0, Math.min(Math.trunc(pane.ySplit), skeleton.rowCount))
    : 0;
  let frozenLeft = 0;
  for (let c = 0; c < xSplit; c++) frozenLeft += skeleton.getColumnWidth(c);
  let frozenTop = 0;
  for (let r = 0; r < ySplit; r++) frozenTop += skeleton.getRowHeight(r);
  if (!frozen) {
    const splitX = pointsToPixels(pane.xSplit / 20);
    const splitY = pointsToPixels(pane.ySplit / 20);
    return createPaneMap(buildPanes(skeleton, viewport, originX, originY, gridWidth, gridHeight, splitX, splitY, 0, 0, 0, 0));
  }

  frozenLeft = Math.min(frozenLeft, gridWidth);
  frozenTop = Math.min(frozenTop, gridHeight);
  // startRow/startColumn describe the initial scroll position saved by Excel;
  // they are seeded into Viewport when the pane is installed, not used as an
  // immutable content origin. This keeps rows between the frozen boundary and
  // the saved position reachable by scrolling back.
  return createPaneMap(buildPanes(skeleton, viewport, originX, originY, gridWidth, gridHeight, frozenLeft, frozenTop, 0, 0, ySplit, xSplit), true);
}

function createPaneMap(panes: RenderPane[], enforceDisjoint = false): PaneMap {
  if (enforceDisjoint) {
    for (let leftIndex = 0; leftIndex < panes.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < panes.length; rightIndex += 1) {
        const left = panes[leftIndex]?.visibleRange;
        const right = panes[rightIndex]?.visibleRange;
        if (left && right && rangesOverlap(left, right)) {
          throw new Error(`Frozen pane visible ranges overlap: ${panes[leftIndex]!.id} and ${panes[rightIndex]!.id}`);
        }
      }
    }
  }
  return {
    panes,
    paneAtLocalPoint(point) {
      return panes.find((pane) => point.x >= pane.screenRect.x && point.x < pane.screenRect.x + pane.screenRect.width
        && point.y >= pane.screenRect.y && point.y < pane.screenRect.y + pane.screenRect.height) ?? null;
    },
    paneForCell(cell) {
      return panes.find((pane) => pane.visibleRange
        && cell.row >= pane.visibleRange.startRow && cell.row <= pane.visibleRange.endRow
        && cell.column >= pane.visibleRange.startColumn && cell.column <= pane.visibleRange.endColumn) ?? null;
    },
  };
}

function rangesOverlap(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow
    && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn
    && right.startColumn <= left.endColumn;
}

function pointsToPixels(points: number): number {
  return Math.max(0, points * (96 / 72));
}

function buildPanes(
  skeleton: SheetSkeleton,
  viewport: ViewportSnapshot,
  originX: number,
  originY: number,
  gridWidth: number,
  gridHeight: number,
  splitWidth: number,
  splitHeight: number,
  startRow: number,
  startColumn: number,
  frozenRows: number,
  frozenColumns: number,
): RenderPane[] {
  const leftWidth = Math.min(gridWidth, splitWidth);
  const topHeight = Math.min(gridHeight, splitHeight);
  const mainStartX = Math.max(0, skeleton.getColumnLeft(Math.max(0, startColumn)));
  const mainStartY = Math.max(0, skeleton.getRowTop(Math.max(0, startRow)));
  const panes: Array<{ id: PaneId; screenRect: Rect; contentOrigin: Point }> = [
    { id: 'topLeft', screenRect: { x: originX, y: originY, width: leftWidth, height: topHeight }, contentOrigin: { x: 0, y: 0 } },
    { id: 'topRight', screenRect: { x: originX + leftWidth, y: originY, width: gridWidth - leftWidth, height: topHeight }, contentOrigin: { x: mainStartX + viewport.scrollX, y: 0 } },
    { id: 'bottomLeft', screenRect: { x: originX, y: originY + topHeight, width: leftWidth, height: gridHeight - topHeight }, contentOrigin: { x: 0, y: mainStartY + viewport.scrollY } },
    { id: 'main', screenRect: { x: originX + leftWidth, y: originY + topHeight, width: gridWidth - leftWidth, height: gridHeight - topHeight }, contentOrigin: { x: mainStartX + viewport.scrollX, y: mainStartY + viewport.scrollY } },
  ];
  return panes.filter((entry) => entry.screenRect.width > 0 && entry.screenRect.height > 0).map((entry) => ({
    ...entry,
    visibleRange: clampFrozenRange(entry.id, skeleton.getVisibleRange({
      x: entry.contentOrigin.x,
      y: entry.contentOrigin.y,
      width: entry.screenRect.width,
      height: entry.screenRect.height,
    }), frozenRows, frozenColumns, skeleton),
  }));
}

function clampFrozenRange(
  paneId: PaneId,
  range: CellRange | null,
  frozenRows: number,
  frozenColumns: number,
  skeleton: SheetSkeleton,
): CellRange | null {
  if (!range || (frozenRows <= 0 && frozenColumns <= 0)) return range;
  const rowStart = paneId === 'topLeft' || paneId === 'topRight' ? 0 : frozenRows;
  const rowEnd = paneId === 'topLeft' || paneId === 'topRight' ? frozenRows - 1 : skeleton.rowCount - 1;
  const columnStart = paneId === 'topLeft' || paneId === 'bottomLeft' ? 0 : frozenColumns;
  const columnEnd = paneId === 'topLeft' || paneId === 'bottomLeft' ? frozenColumns - 1 : skeleton.columnCount - 1;
  const next = {
    startRow: Math.max(range.startRow, rowStart),
    endRow: Math.min(range.endRow, rowEnd),
    startColumn: Math.max(range.startColumn, columnStart),
    endColumn: Math.min(range.endColumn, columnEnd),
  };
  if (next.startRow > next.endRow || next.startColumn > next.endColumn) return null;
  return next;
}

/** 默认表头偏移 */
export function defaultHeaderOffset(): Point {
  return { x: ROW_HEADER_WIDTH, y: COL_HEADER_HEIGHT };
}

function emptyScrollDelta(): ScrollDeltaPlan {
  return {
    delta: { x: 0, y: 0 },
    dx: 0,
    dy: 0,
    hasDelta: false,
    redrawRects: [],
  };
}

export function calculateScrollDelta(
  previousViewport: ViewportSnapshot | null | undefined,
  nextViewport: ViewportSnapshot,
): ScrollDeltaPlan {
  if (!previousViewport) return emptyScrollDelta();
  const dx = nextViewport.scrollX - previousViewport.scrollX;
  const dy = nextViewport.scrollY - previousViewport.scrollY;
  const hasDelta = dx !== 0 || dy !== 0;
  return {
    delta: { x: dx, y: dy },
    dx,
    dy,
    hasDelta,
    redrawRects: [],
  };
}

/**
 * Resolves the viewport delta into disjoint screen-space pane operations.
 * PaneMap is deliberately the only coordinate authority here: this prevents
 * frozen headers and panes from being translated as part of the main grid.
 */
function resolvePaneScrollDelta(base: ScrollDeltaPlan, panes: readonly RenderPane[]): ScrollDeltaPlan {
  if (!base.hasDelta) return base;
  const redrawRects: Rect[] = [];
  for (const pane of panes) {
    const dx = pane.id === 'topLeft' || pane.id === 'bottomLeft' ? 0 : base.dx;
    const dy = pane.id === 'topLeft' || pane.id === 'topRight' ? 0 : base.dy;
    if (dx === 0 && dy === 0) continue;
    redrawRects.push({ ...pane.screenRect });
  }
  return {
    ...base,
    redrawRects,
  };
}

function combineScrollRedrawRects(scrollRects: readonly Rect[], dirtyRects: readonly Rect[]): Rect[] {
  const redrawRects = scrollRects.map((rect) => ({ ...rect }));
  for (const dirtyRect of dirtyRects) {
    const coveredByScrolledPane = scrollRects.some((paneRect) => dirtyRect.x >= paneRect.x
      && dirtyRect.y >= paneRect.y
      && dirtyRect.x + dirtyRect.width <= paneRect.x + paneRect.width
      && dirtyRect.y + dirtyRect.height <= paneRect.y + paneRect.height);
    if (!coveredByScrolledPane) redrawRects.push({ ...dirtyRect });
  }
  return redrawRects;
}

export function rangeToViewportRect(
  range: CellRange,
  skeleton: SheetSkeleton,
  viewport: ViewportSnapshot,
): Rect | null {
  const sheetRect = skeleton.getRangeRect(range);
  if (!sheetRect) return null;
  return intersectRect(
    translateRect(sheetRect, { x: -viewport.scrollX, y: -viewport.scrollY }),
    { x: 0, y: 0, width: viewport.width, height: viewport.height },
  );
}

function viewportChanged(previous: ViewportSnapshot | null | undefined, next: ViewportSnapshot): boolean {
  return !previous
    || previous.width !== next.width
    || previous.height !== next.height
    || previous.devicePixelRatio !== next.devicePixelRatio;
}

function calculateReason(
  hasPrevious: boolean,
  forceFull: boolean,
  resized: boolean,
  redrawScroll: boolean,
  hasDirtyRects: boolean,
): RenderPlanReason {
  if (!hasPrevious) return 'initial';
  if (forceFull) return 'forced';
  if (resized) return 'resize';
  if (redrawScroll && hasDirtyRects) return 'mixed';
  if (redrawScroll) return 'scroll-redraw';
  if (hasDirtyRects) return 'dirty';
  return 'idle';
}

export function calculateRenderPlan(input: RenderPlanInput): RenderPlan {
  const previousViewport = input.previousViewport ?? null;
  const hasPrevious = previousViewport !== null;
  const baseScrollDelta = calculateScrollDelta(previousViewport, input.viewport);
  const paneMap = computePaneMap(input.skeleton, input.viewport, input.pane ?? null, input.headerOffset ?? null);
  const panes = [...paneMap.panes];
  const scrollDelta = resolvePaneScrollDelta(baseScrollDelta, panes);
  const dirtyRanges = mergeCellRanges(input.dirtyRanges ?? []);
  const dirtyRects = mergeRects(
    dirtyRanges
      .map((range) => rangeToViewportRect(range, input.skeleton, input.viewport))
      .filter((rect): rect is Rect => rect !== null),
  );
  const resized = viewportChanged(previousViewport, input.viewport);
  const redrawScroll = scrollDelta.hasDelta;
  // Scrolling redraws only panes whose content origin changed. It must not
  // clear the whole canvas, because that includes headers and frozen panes.
  const fullRedraw = input.forceFull === true || !hasPrevious || resized;
  const reason = calculateReason(
    hasPrevious,
    input.forceFull === true,
    resized,
    redrawScroll,
    dirtyRects.length > 0,
  );

  const layerDefinitions = input.layers?.length ? input.layers : DEFAULT_LAYER_DEFINITIONS;
  const layers = layerDefinitions.map((definition): LayerRenderPlan => {
    if (fullRedraw) return { layerId: definition.id, mode: 'full', clearRects: [], drawRects: [] };
    if (scrollDelta.hasDelta) {
      if (definition.scrollable === false) {
        return { layerId: definition.id, mode: 'full', clearRects: [], drawRects: [] };
      }
      const redrawRects = combineScrollRedrawRects(scrollDelta.redrawRects, dirtyRects);
      return { layerId: definition.id, mode: 'dirty', clearRects: redrawRects, drawRects: redrawRects };
    }
    if (dirtyRects.length > 0) {
      return { layerId: definition.id, mode: 'dirty', clearRects: dirtyRects, drawRects: dirtyRects };
    }
    if (input.chromeDirty && definition.id === 'chrome') {
      return { layerId: definition.id, mode: 'full', clearRects: [], drawRects: [] };
    }
    return { layerId: definition.id, mode: 'none', clearRects: [], drawRects: [] };
  });

  return {
    viewport: { ...input.viewport },
    paneMap,
    visibleRange: panes.at(-1)?.visibleRange ?? null,
    panes,
    dirtyRanges,
    dirtyRects,
    scrollDelta,
    fullRedraw,
    reason,
    layers,
  };
}

export const createRenderPlan = calculateRenderPlan;
