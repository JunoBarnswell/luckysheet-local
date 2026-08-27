import { drawCellLayer, drawExtensionsLayer, drawGridLayer, type AssetUrlResolver } from "./cell-renderer";
import { DirtyRangeSet } from "./dirty-ranges";
import { calculateRenderPlan, computePaneMap, defaultHeaderOffset, type RenderPlan } from "./render-plan";
import { Scene } from "./scene";
import { SheetSkeleton } from "./sheet-skeleton";
import { Viewport } from "./viewport";
import { drawChromeLayer } from "./chrome-renderer";
import {
  COL_HEADER_HEIGHT,
  RESIZE_HIT_TOLERANCE_PX,
  ROW_HEADER_WIDTH,
  createEmptyChromeState,
  type CellAddress,
  type CellProvider,
  type CellRange,
  type CellRenderData,
  type ChromeState,
  type FloatingDrawable,
  type FloatingHandle,
  type FloatingHit,
  type HeaderHit,
  type LayerDefinition,
  type Point,
  type Rect,
  type RenderPane,
  type RenderTheme,
  type ViewportSnapshot,
  type PaneLayout,
  DEFAULT_LAYER_DEFINITIONS as DEFAULT_LAYERS_SOURCE,
  DEFAULT_RENDER_THEME,
} from "./types";

export interface CanvasRenderEngineOptions {
  skeleton?: SheetSkeleton;
  viewport?: Partial<ViewportSnapshot>;
  cellProvider?: CellProvider;
  cells?: ReadonlyMap<string, CellRenderData>;
  theme?: Partial<RenderTheme>;
  layers?: readonly LayerDefinition[];
  resolveAssetUrl?: AssetUrlResolver;
  assetUrlCache?: Map<string, string>;
  assetUrlPending?: Set<string>;
  assetUrlErrors?: Map<string, string>;
}

function cellMapKey(row: number, column: number): string {
  return row + ":" + column;
}

function createMapProvider(cells: ReadonlyMap<string, CellRenderData>): CellProvider {
  return ({ row, column }) => cells.get(cellMapKey(row, column));
}

function mergeTheme(theme: Partial<RenderTheme> | undefined): RenderTheme {
  return { ...DEFAULT_RENDER_THEME, ...theme };
}



export class CanvasRenderEngine {
  readonly viewport: Viewport;

  private skeletonModel: SheetSkeleton;
  private cellProvider: CellProvider;
  private theme: RenderTheme;
  private readonly layerDefinitions: LayerDefinition[];
  private scene: Scene;
  private readonly dirtyRanges = new DirtyRangeSet();
  private previousViewport: ViewportSnapshot | null = null;
  private lastPlan: RenderPlan | null = null;
  private forceFullRedraw = true;
  private chromeDirty = false;
  private frameHandle: number | null = null;
  private frameUsesAnimationFrame = false;
  private disposed = false;

  private paneLayout: PaneLayout | null = null;
  private readonly headerOrigin: Point = defaultHeaderOffset();
  private chrome: ChromeState = createEmptyChromeState();
  private floatables: readonly FloatingDrawable[] = [];
  private readonly viewportListeners = new Set<() => void>();
  private readonly resolveAssetUrl?: AssetUrlResolver;
  private readonly assetUrlCache: Map<string, string>;
  private readonly assetUrlPending: Set<string>;
  private readonly assetUrlErrors: Map<string, string>;

  constructor(options: CanvasRenderEngineOptions = {}) {
    this.skeletonModel = options.skeleton ?? new SheetSkeleton({ rowCount: 1000, columnCount: 26 });
    this.viewport = new Viewport(options.viewport);
    this.theme = mergeTheme(options.theme);
    this.layerDefinitions = (options.layers ?? DEFAULT_LAYERS_SOURCE).map((definition) => ({ ...definition }));
    this.scene = new Scene(this.layerDefinitions);
    this.resolveAssetUrl = options.resolveAssetUrl;
    this.assetUrlCache = options.assetUrlCache ?? new Map<string, string>();
    this.assetUrlPending = options.assetUrlPending ?? new Set<string>();
    this.assetUrlErrors = options.assetUrlErrors ?? new Map<string, string>();
    this.cellProvider = options.cellProvider
      ?? (options.cells ? createMapProvider(options.cells) : () => undefined);
    this.viewport.clampTo(this.skeletonModel.contentSize);
  }

  get skeleton(): SheetSkeleton {
    return this.skeletonModel;
  }

  get sceneGraph(): Scene {
    return this.scene;
  }

  get mounted(): boolean {
    return this.scene.mounted;
  }

  get lastRenderPlan(): RenderPlan | null {
    return this.lastPlan;
  }

  get headerOffset(): Point {
    return { ...this.headerOrigin };
  }

  mount(host: HTMLElement): void {
    this.reviveIfDisposed();
    this.scene.mount(host);
    const width = this.viewport.width || host.clientWidth;
    const height = this.viewport.height || host.clientHeight;
    const devicePixelRatio = this.viewport.devicePixelRatio;
    this.viewport.setSize(width, height, devicePixelRatio);
    this.scene.resize({ width, height }, devicePixelRatio);
    this.previousViewport = null;
    this.forceFullRedraw = true;
    this.render();
  }

  unmount(): void {
    this.cancelScheduledRender();
    this.scene.unmount();
    this.previousViewport = null;
    this.forceFullRedraw = true;
  }

  dispose(): void {
    if (this.disposed) return;
    this.cancelScheduledRender();
    this.scene.dispose();
    this.disposed = true;
    this.previousViewport = null;
    this.lastPlan = null;
    this.viewportListeners.clear();
  }

  destroy(): void {
    this.dispose();
  }

  resize(width: number, height: number, devicePixelRatio = this.viewport.devicePixelRatio): void {
    this.assertActive();
    this.viewport.setSize(width, height, devicePixelRatio);
    if (this.scene.mounted) this.scene.resize({ width, height }, devicePixelRatio);
    this.forceFullRedraw = true;
    this.requestRender();
  }

  resizeFromHost(): void {
    this.assertActive();
    const host = this.scene.currentHost;
    if (!host) return;
    this.resize(host.clientWidth, host.clientHeight, this.viewport.devicePixelRatio);
  }

  setViewport(snapshot: Partial<ViewportSnapshot>): void {
    this.assertActive();
    const before = this.viewport.getSnapshot();
    const after = this.viewport.setSnapshot(snapshot);
    this.viewport.clampTo(this.skeletonModel.contentSize);
    if (before.width !== after.width || before.height !== after.height || before.devicePixelRatio !== after.devicePixelRatio) {
      this.scene.resize({ width: after.width, height: after.height }, after.devicePixelRatio);
      this.forceFullRedraw = true;
    }
    this.requestRender();
  }

  scrollTo(x: number, y: number): void {
    this.assertActive();
    this.viewport.scrollTo(x, y, this.skeletonModel.contentSize);
    this.requestRender();
  }

  scrollBy(deltaX: number, deltaY: number): void {
    this.assertActive();
    this.viewport.scrollBy(deltaX, deltaY, this.skeletonModel.contentSize);
    this.requestRender();
  }

  setSkeleton(skeleton: SheetSkeleton): void {
    this.assertActive();
    this.skeletonModel = skeleton;
    this.viewport.clampTo(skeleton.contentSize);
    this.forceFullRedraw = true;
    this.requestRender();
  }

  setCellProvider(cellProvider: CellProvider): void {
    this.assertActive();
    this.cellProvider = cellProvider;
    this.forceFullRedraw = true;
    this.requestRender();
  }

  setCells(cells: ReadonlyMap<string, CellRenderData>): void {
    this.setCellProvider(createMapProvider(cells));
  }

  setTheme(theme: Partial<RenderTheme>): void {
    this.assertActive();
    this.theme = mergeTheme(theme);
    this.forceFullRedraw = true;
    this.requestRender();
  }

  invalidate(ranges?: readonly CellRange[]): void {
    this.assertActive();
    if (ranges === undefined) {
      this.forceFullRedraw = true;
      this.requestRender();
      return;
    }
    this.dirtyRanges.addMany(ranges);
    this.requestRender();
  }

  markDirty(ranges?: readonly CellRange[]): void {
    this.invalidate(ranges);
  }

  // ---------- 冻结 / Chrome / 浮动对象 ----------

  setPane(pane: PaneLayout | null): void {
    this.assertActive();
    if (!this.paneLayout && (pane?.kind === 'frozen' || pane?.kind === 'split')) {
      this.viewport.setSnapshot({
        scrollX: this.skeletonModel.getColumnLeft(Math.max(0, Math.trunc(pane.startColumn))),
        scrollY: this.skeletonModel.getRowTop(Math.max(0, Math.trunc(pane.startRow))),
      });
      this.viewport.clampTo(this.skeletonModel.contentSize);
    }
    this.paneLayout = pane && pane.kind !== 'none' && (pane.xSplit > 0 || pane.ySplit > 0) ? structuredClone(pane) : null;
    this.forceFullRedraw = true;
    this.requestRender();
  }

  invalidateChrome(): void {
    this.assertActive();
    this.chromeDirty = true;
    this.requestRender();
  }

  setChrome(state: ChromeState): void {
    this.assertActive();
    this.chrome = state;
    this.invalidateChrome();
  }

  setFloating(objects: readonly FloatingDrawable[], selectedId: string | null = null): void {
    this.assertActive();
    const objectsChanged = objects !== this.floatables;
    this.floatables = objects;
    this.chrome.selectedFloatingId = selectedId;
    if (objectsChanged) {
      // Floating drawables close over canonical sheet snapshots. When a source
      // range changes (including a cross-sheet Camera source), the extensions
      // layer must be repainted; chrome-only invalidation would leave a stale
      // camera bitmap on screen.
      this.forceFullRedraw = true;
      this.requestRender();
    } else {
      this.invalidateChrome();
    }
  }

  onViewportChanged(listener: () => void): () => void {
    this.viewportListeners.add(listener);
    return () => this.viewportListeners.delete(listener);
  }

  /** 保证单元格进入主窗格可视区(考虑冻结条占用) */
  ensureVisible(cell: CellAddress): void {
    this.assertActive();
    const rect = this.skeletonModel.getCellRect(cell.row, cell.column);
    if (!rect) return;
    const origin = this.headerOrigin;
    let frozenLeft = 0;
    let frozenTop = 0;
    if (this.paneLayout?.kind === 'frozen') {
      for (let c = 0; c < this.paneLayout.xSplit; c++) frozenLeft += this.skeletonModel.getColumnWidth(c);
      for (let r = 0; r < this.paneLayout.ySplit; r++) frozenTop += this.skeletonModel.getRowHeight(r);
    }
    const viewLeft = origin.x + frozenLeft;
    const viewTop = origin.y + frozenTop;
    const viewWidth = Math.max(1, this.viewport.width - viewLeft);
    const viewHeight = Math.max(1, this.viewport.height - viewTop);

    // 单元格在冻结区内的行列无需滚动
    const inFrozenColumn = this.paneLayout?.kind === 'frozen' && cell.column < this.paneLayout.xSplit;
    const inFrozenRow = this.paneLayout?.kind === 'frozen' && cell.row < this.paneLayout.ySplit;

    let nextScrollX = this.viewport.scrollX;
    let nextScrollY = this.viewport.scrollY;
    if (!inFrozenColumn) {
      const left = rect.x - viewLeft;
      const right = left + rect.width;
      if (left < nextScrollX) nextScrollX = left;
      else if (right > nextScrollX + viewWidth) nextScrollX = right - viewWidth;
    }
    if (!inFrozenRow) {
      const top = rect.y - viewTop;
      const bottom = top + rect.height;
      if (top < nextScrollY) nextScrollY = top;
      else if (bottom > nextScrollY + viewHeight) nextScrollY = bottom - viewHeight;
    }
    if (nextScrollX !== this.viewport.scrollX || nextScrollY !== this.viewport.scrollY) {
      this.scrollTo(nextScrollX, nextScrollY);
    }
  }

  /** 屏幕本地坐标 → 所在窗格 */
  paneAtLocalPoint(local: Point): RenderPane | null {
    return this.currentPaneMap().paneAtLocalPoint(local);
  }

  /** 屏幕本地坐标 → 内容坐标(按所在窗格的偏移换算) */
  localToContent(local: Point): Point {
    const pane = this.paneAtLocalPoint(local);
    if (!pane) return { x: local.x, y: local.y };
    return { x: local.x - pane.screenRect.x + pane.contentOrigin.x, y: local.y - pane.screenRect.y + pane.contentOrigin.y };
  }

  /** 屏幕本地坐标 → 模型单元格(null = 表头区或越界) */
  cellAtLocalPoint(local: Point): CellAddress | null {
    const origin = this.headerOrigin;
    if (local.x < origin.x || local.y < origin.y) return null;
    const content = this.localToContent(local);
    return this.skeletonModel.getCellAtPoint(content);
  }

  /** 表头命中:角块/行头/列头 + 调整热区 */
  headerHitAtLocal(local: Point): HeaderHit | null {
    const origin = this.headerOrigin;
    if (local.x < origin.x && local.y < origin.y) return { kind: "corner", index: 0 };
    if (local.y < origin.y && local.x >= origin.x) {
      const pane = this.paneAtLocalPoint({ x: local.x, y: origin.y + 1 });
      if (!pane) return null;
      const contentX = local.x - pane.screenRect.x + pane.contentOrigin.x;
      const column = this.skeletonModel.findColumnAt(contentX);
      if (column < 0) return null;
      const boundary = this.skeletonModel.findNearestColumnBoundary(contentX, RESIZE_HIT_TOLERANCE_PX);
      return boundary
        ? { kind: "col", index: boundary.index, resizeBoundaryPx: boundary.deltaPx }
        : { kind: "col", index: column };
    }
    if (local.x < origin.x && local.y >= origin.y) {
      const pane = this.paneAtLocalPoint({ x: origin.x + 1, y: local.y });
      if (!pane) return null;
      const contentY = local.y - pane.screenRect.y + pane.contentOrigin.y;
      const row = this.skeletonModel.findRowAt(contentY);
      if (row < 0) return null;
      const boundary = this.skeletonModel.getRowTop(row) + this.skeletonModel.getRowHeight(row) - contentY;
      const resizeBoundaryPx = Math.abs(boundary) <= RESIZE_HIT_TOLERANCE_PX ? boundary : undefined;
      return { kind: "row", index: row, resizeBoundaryPx };
    }
    return null;
  }

  /**
   * 内容坐标 → PaneMap 决定的屏幕坐标。
   *
   * 单元格编辑器必须传入 cell；浮动对象使用内容点定位。
   */
  contentToScreen(content: Point, cell?: CellAddress): Point {
    const paneMap = this.currentPaneMap();
    const panes = paneMap.panes;
    const main = panes.find((pane) => pane.id === "main") ?? panes.at(-1)!;
    const target = cell ? paneMap.paneForCell(cell) : undefined;
    const pane = target ?? panes.find((candidate) => content.x >= candidate.contentOrigin.x
      && content.x <= candidate.contentOrigin.x + candidate.screenRect.width
      && content.y >= candidate.contentOrigin.y
      && content.y <= candidate.contentOrigin.y + candidate.screenRect.height) ?? main;
    return { x: content.x - pane.contentOrigin.x + pane.screenRect.x, y: content.y - pane.contentOrigin.y + pane.screenRect.y };
  }

  /**
   * Convert a content range into one screen rect per visible freeze pane.
   * DOM overlays must use this instead of assuming the main pane is the only
   * coordinate space.
   */
  contentRangeToScreenRects(range: CellRange): Rect[] {
    const content = this.skeletonModel.getRangeRect(range);
    if (!content) return [];
    const panes = this.currentPaneMap().panes;
    const rects: Rect[] = [];
    for (const pane of panes) {
      const left = Math.max(content.x, pane.contentOrigin.x);
      const top = Math.max(content.y, pane.contentOrigin.y);
      const right = Math.min(content.x + content.width, pane.contentOrigin.x + pane.screenRect.width);
      const bottom = Math.min(content.y + content.height, pane.contentOrigin.y + pane.screenRect.height);
      if (right <= left || bottom <= top) continue;
      rects.push({
        x: left - pane.contentOrigin.x + pane.screenRect.x,
        y: top - pane.contentOrigin.y + pane.screenRect.y,
        width: right - left,
        height: bottom - top,
      });
    }
    return rects;
  }

  hitTestFloating(local: Point): FloatingHit | null {
    const selectedId = this.chrome.selectedFloatingId;
    // 已选中对象的缩放手柄优先
    if (selectedId) {
      const handleHit = this.hitFloatingHandles(local, selectedId);
      if (handleHit) return handleHit;
    }
    for (let i = this.floatables.length - 1; i >= 0; i--) {
      const drawable = this.floatables[i]!;
      const screen = this.contentToScreen(drawable.bounds);
      if (local.x >= screen.x && local.x <= screen.x + drawable.bounds.width
        && local.y >= screen.y && local.y <= screen.y + drawable.bounds.height) {
        const control = drawable.hitTest?.({ x: local.x - screen.x, y: local.y - screen.y }) ?? undefined;
        return control ? { kind: drawable.kind, id: drawable.id, control } : { kind: drawable.kind, id: drawable.id };
      }
    }
    return null;
  }

  hitFloatingHandles(local: Point, id: string): FloatingHit | null {
    const drawable = this.floatables.find((item) => item.id === id);
    if (!drawable) return null;
    const screen = this.contentToScreen(drawable.bounds);
    const tolerance = 5;
    const handles: Array<[FloatingHandle, Point]> = [
      ["nw", { x: screen.x, y: screen.y }],
      ["n", { x: screen.x + drawable.bounds.width / 2, y: screen.y }],
      ["ne", { x: screen.x + drawable.bounds.width, y: screen.y }],
      ["e", { x: screen.x + drawable.bounds.width, y: screen.y + drawable.bounds.height / 2 }],
      ["se", { x: screen.x + drawable.bounds.width, y: screen.y + drawable.bounds.height }],
      ["s", { x: screen.x + drawable.bounds.width / 2, y: screen.y + drawable.bounds.height }],
      ["sw", { x: screen.x, y: screen.y + drawable.bounds.height }],
      ["w", { x: screen.x, y: screen.y + drawable.bounds.height / 2 }],
    ];
    for (const [handle, point] of handles) {
      if (Math.abs(local.x - point.x) <= tolerance && Math.abs(local.y - point.y) <= tolerance) {
        return { kind: drawable.kind, id, handle };
      }
    }
    return null;
  }

  get floatingSelectionId(): string | null {
    return this.chrome.selectedFloatingId;
  }

  // ---------- 渲染 ----------

  render(): RenderPlan {
    this.assertActive();
    const viewport = this.viewport.getSnapshot();
    const plan = calculateRenderPlan({
      skeleton: this.skeletonModel,
      viewport,
      previousViewport: this.previousViewport,
      dirtyRanges: this.dirtyRanges.toArray(),
      forceFull: this.forceFullRedraw,
      chromeDirty: this.chromeDirty,
      layers: this.layerDefinitions,
      pane: this.paneLayout,
      headerOffset: this.headerOrigin,
    });
    if (this.scene.mounted) this.applyPlan(plan);

    const before = this.previousViewport;
    this.previousViewport = viewport;
    this.lastPlan = plan;
    this.forceFullRedraw = false;
    this.chromeDirty = false;
    this.dirtyRanges.clear();

    if (before
      && (before.scrollX !== viewport.scrollX || before.scrollY !== viewport.scrollY
        || before.width !== viewport.width || before.height !== viewport.height)) {
      for (const listener of this.viewportListeners) listener();
    }
    return plan;
  }

  requestRender(): void {
    this.assertActive();
    if (!this.scene.mounted || this.frameHandle !== null) return;
    if (typeof window !== "undefined" && typeof window.requestAnimationFrame === "function") {
      this.frameUsesAnimationFrame = true;
      this.frameHandle = window.requestAnimationFrame(() => {
        this.frameHandle = null;
        this.render();
      });
      return;
    }
    this.frameUsesAnimationFrame = false;
    this.frameHandle = setTimeout(() => {
      this.frameHandle = null;
      this.render();
    }, 0) as unknown as number;
  }

  getCanvas(layerId: string): HTMLCanvasElement | null {
    return this.scene.getLayer(layerId)?.element ?? null;
  }

  getLayer(layerId: string) {
    return this.scene.getLayer(layerId);
}

  private applyPlan(plan: RenderPlan): void {
    for (const layerPlan of plan.layers) {
      const layer = this.scene.getLayer(layerPlan.layerId);
      if (!layer || layerPlan.mode === "none") continue;
      if (layerPlan.mode === "full") {
        layer.clear();
        this.drawLayer(layerPlan.layerId, plan);
        continue;
      }
      if (layerPlan.mode === "scroll") layer.blit(plan.scrollDelta);
      layer.clear(layerPlan.clearRects);
      this.drawLayer(layerPlan.layerId, plan, layerPlan.drawRects);
    }
  }

  private drawLayer(layerId: string, plan: RenderPlan, drawRects?: readonly Rect[]): void {
    const layer = this.scene.getLayer(layerId);
    if (!layer) return;
    layer.withLogicalContext((context) => {
      if (layerId === "chrome") {
        drawChromeLayer({
          context,
          skeleton: this.skeletonModel,
          plan,
          chrome: this.chrome,
          theme: this.theme,
        });
        return;
      }
      for (const pane of plan.panes) {
        context.save();
        context.beginPath();
        context.rect(pane.screenRect.x, pane.screenRect.y, pane.screenRect.width, pane.screenRect.height);
        context.clip();
        context.translate(pane.screenRect.x - pane.contentOrigin.x, pane.screenRect.y - pane.contentOrigin.y);
        const options = {
          context,
          skeleton: this.skeletonModel,
          pane,
          visibleRange: pane.visibleRange,
          cellProvider: this.cellProvider,
          theme: this.theme,
          drawRects: convertDrawRectsForPane(drawRects, pane),
          resolveAssetUrl: this.resolveAssetUrl,
          assetUrlCache: this.assetUrlCache,
          assetUrlPending: this.assetUrlPending,
          assetUrlErrors: this.assetUrlErrors,
          requestRender: () => this.requestRender(),
        };
        if (layerId === "grid") drawGridLayer(options);
        else if (layerId === "content") drawCellLayer(options);
        else if (layerId === "extensions") {
          drawExtensionsLayer({ ...options, floatables: this.floatables });
        }
        context.restore();
      }
    });
  }

  private cancelScheduledRender(): void {
    if (this.frameHandle === null) return;
    if (this.frameUsesAnimationFrame && typeof window !== "undefined") {
      window.cancelAnimationFrame(this.frameHandle);
    } else {
      clearTimeout(this.frameHandle as unknown as ReturnType<typeof setTimeout>);
    }
    this.frameHandle = null;
  }

  private assertActive(): void {
    if (this.disposed) throw new Error("CanvasRenderEngine has been disposed");
  }

  private currentPaneMap() {
    return this.lastPlan?.paneMap
      ?? computePaneMap(this.skeletonModel, this.viewport.getSnapshot(), this.paneLayout, this.headerOrigin);
  }

  private reviveIfDisposed(): void {
    if (!this.disposed) return;
    this.scene = new Scene(this.layerDefinitions);
    this.disposed = false;
    this.lastPlan = null;
  }
}

/** 屏幕坐标矩形 → 该窗格内容坐标矩形 */
function convertDrawRectsForPane(rects: readonly Rect[] | undefined, pane: RenderPane): Rect[] | undefined {
  if (!rects) return undefined;
  const converted: Rect[] = [];
  for (const rect of rects) {
    const left = Math.max(rect.x, pane.screenRect.x);
    const top = Math.max(rect.y, pane.screenRect.y);
    const right = Math.min(rect.x + rect.width, pane.screenRect.x + pane.screenRect.width);
    const bottom = Math.min(rect.y + rect.height, pane.screenRect.y + pane.screenRect.height);
    if (right <= left || bottom <= top) continue;
    converted.push({
      x: left - pane.screenRect.x + pane.contentOrigin.x,
      y: top - pane.screenRect.y + pane.contentOrigin.y,
      width: right - left,
      height: bottom - top,
    });
  }
  return converted;
}
