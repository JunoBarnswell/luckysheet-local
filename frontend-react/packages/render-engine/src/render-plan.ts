import { mergeCellRanges } from './dirty-ranges';
import { intersectRect, mergeRects, translateRect } from './geometry';
import { SheetSkeleton } from './sheet-skeleton';
import {
  DEFAULT_LAYER_DEFINITIONS,
  type CellRange,
  type LayerDefinition,
  type Point,
  type Rect,
  type ViewportSnapshot,
} from './types';

export type LayerRenderMode = 'none' | 'full' | 'dirty' | 'scroll';

export interface ScrollDeltaPlan {
  delta: Point;
  dx: number;
  dy: number;
  hasDelta: boolean;
  isSmall: boolean;
  canBlit: boolean;
  source: Rect | null;
  destination: Rect | null;
  copySource: Rect | null;
  copyDestination: Rect | null;
  exposedRects: Rect[];
}

export interface LayerRenderPlan {
  layerId: string;
  mode: LayerRenderMode;
  clearRects: Rect[];
  drawRects: Rect[];
}

export type RenderPlanReason = 'initial' | 'forced' | 'resize' | 'large-scroll' | 'mixed' | 'scroll' | 'dirty' | 'idle';

export interface RenderPlan {
  viewport: ViewportSnapshot;
  visibleRange: CellRange | null;
  dirtyRanges: CellRange[];
  dirtyRects: Rect[];
  scrollDelta: ScrollDeltaPlan;
  scroll: ScrollDeltaPlan;
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
  layers?: readonly LayerDefinition[];
}

function emptyScrollDelta(): ScrollDeltaPlan {
  return {
    delta: { x: 0, y: 0 },
    dx: 0,
    dy: 0,
    hasDelta: false,
    isSmall: false,
    canBlit: false,
    source: null,
    destination: null,
    copySource: null,
    copyDestination: null,
    exposedRects: [],
  };
}

export function getScrollExposedRects(viewport: SizeLike, dx: number, dy: number): Rect[] {
  const exposed: Rect[] = [];
  if (dy > 0) exposed.push({ x: 0, y: viewport.height - dy, width: viewport.width, height: dy });
  if (dy < 0) exposed.push({ x: 0, y: 0, width: viewport.width, height: -dy });
  if (dx > 0) exposed.push({ x: viewport.width - dx, y: 0, width: dx, height: viewport.height });
  if (dx < 0) exposed.push({ x: 0, y: 0, width: -dx, height: viewport.height });
  return exposed.filter((rect) => rect.width > 0 && rect.height > 0);
}

interface SizeLike {
  width: number;
  height: number;
}

export function calculateScrollDelta(
  previousViewport: ViewportSnapshot | null | undefined,
  nextViewport: ViewportSnapshot,
): ScrollDeltaPlan {
  if (!previousViewport) return emptyScrollDelta();
  const dx = nextViewport.scrollX - previousViewport.scrollX;
  const dy = nextViewport.scrollY - previousViewport.scrollY;
  const hasDelta = dx !== 0 || dy !== 0;
  const sameSize = previousViewport.width === nextViewport.width
    && previousViewport.height === nextViewport.height;
  const isSmall = sameSize
    && nextViewport.width > 0
    && nextViewport.height > 0
    && Math.abs(dx) < nextViewport.width
    && Math.abs(dy) < nextViewport.height;
  const canBlit = hasDelta && isSmall;
  if (!canBlit) {
    return {
      delta: { x: dx, y: dy },
      dx,
      dy,
      hasDelta,
      isSmall,
      canBlit: false,
      source: null,
      destination: null,
      copySource: null,
      copyDestination: null,
      exposedRects: [],
    };
  }

  const copyWidth = nextViewport.width - Math.abs(dx);
  const copyHeight = nextViewport.height - Math.abs(dy);
  const source = {
    x: Math.max(dx, 0),
    y: Math.max(dy, 0),
    width: copyWidth,
    height: copyHeight,
  };
  const destination = {
    x: Math.max(-dx, 0),
    y: Math.max(-dy, 0),
    width: copyWidth,
    height: copyHeight,
  };
  const exposedRects = getScrollExposedRects(nextViewport, dx, dy);
  return {
    delta: { x: dx, y: dy },
    dx,
    dy,
    hasDelta,
    isSmall,
    canBlit,
    source,
    destination,
    copySource: source,
    copyDestination: destination,
    exposedRects,
  };
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
  largeScroll: boolean,
  hasScroll: boolean,
  hasDirtyRects: boolean,
): RenderPlanReason {
  if (!hasPrevious) return 'initial';
  if (forceFull) return 'forced';
  if (resized) return 'resize';
  if (largeScroll && hasDirtyRects) return 'mixed';
  if (largeScroll) return 'large-scroll';
  if (hasScroll && hasDirtyRects) return 'mixed';
  if (hasScroll) return 'scroll';
  if (hasDirtyRects) return 'dirty';
  return 'idle';
}

export function calculateRenderPlan(input: RenderPlanInput): RenderPlan {
  const previousViewport = input.previousViewport ?? null;
  const hasPrevious = previousViewport !== null;
  const scrollDelta = calculateScrollDelta(previousViewport, input.viewport);
  const dirtyRanges = mergeCellRanges(input.dirtyRanges ?? []);
  const dirtyRects = mergeRects(
    dirtyRanges
      .map((range) => rangeToViewportRect(range, input.skeleton, input.viewport))
      .filter((rect): rect is Rect => rect !== null),
  );
  const resized = viewportChanged(previousViewport, input.viewport);
  const largeScroll = scrollDelta.hasDelta && !scrollDelta.canBlit;
  const fullRedraw = input.forceFull === true || !hasPrevious || resized || largeScroll;
  const reason = calculateReason(
    hasPrevious,
    input.forceFull === true,
    resized,
    largeScroll,
    scrollDelta.hasDelta,
    dirtyRects.length > 0,
  );

  const layerDefinitions = input.layers?.length ? input.layers : DEFAULT_LAYER_DEFINITIONS;
  const layers = layerDefinitions.map((definition): LayerRenderPlan => {
    if (fullRedraw) return { layerId: definition.id, mode: 'full', clearRects: [], drawRects: [] };
    if (scrollDelta.hasDelta) {
      if (definition.scrollable === false) {
        return { layerId: definition.id, mode: 'full', clearRects: [], drawRects: [] };
      }
      const redrawRects = mergeRects([...scrollDelta.exposedRects, ...dirtyRects]);
      return { layerId: definition.id, mode: 'scroll', clearRects: redrawRects, drawRects: redrawRects };
    }
    if (dirtyRects.length > 0) {
      return { layerId: definition.id, mode: 'dirty', clearRects: dirtyRects, drawRects: dirtyRects };
    }
    return { layerId: definition.id, mode: 'none', clearRects: [], drawRects: [] };
  });

  return {
    viewport: { ...input.viewport },
    visibleRange: input.skeleton.getVisibleRange({
      x: input.viewport.scrollX,
      y: input.viewport.scrollY,
      width: input.viewport.width,
      height: input.viewport.height,
    }),
    dirtyRanges,
    dirtyRects,
    scrollDelta,
    scroll: scrollDelta,
    fullRedraw,
    reason,
    layers,
  };
}

export const createRenderPlan = calculateRenderPlan;
export const calculateSmallScrollDelta = calculateScrollDelta;
