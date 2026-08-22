import { drawCellLayer, drawGridLayer } from './cell-renderer';
import { DirtyRangeSet } from './dirty-ranges';
import { calculateRenderPlan, type RenderPlan } from './render-plan';
import { Scene } from './scene';
import { SheetSkeleton } from './sheet-skeleton';
import { Viewport } from './viewport';
import {
  DEFAULT_LAYER_DEFINITIONS,
  DEFAULT_RENDER_THEME,
  type CellProvider,
  type CellRange,
  type CellRenderData,
  type LayerDefinition,
  type Rect,
  type RenderTheme,
  type ViewportSnapshot,
} from './types';

export interface CanvasRenderEngineOptions {
  skeleton?: SheetSkeleton;
  viewport?: Partial<ViewportSnapshot>;
  cellProvider?: CellProvider;
  cells?: ReadonlyMap<string, CellRenderData>;
  theme?: Partial<RenderTheme>;
  layers?: readonly LayerDefinition[];
}

function cellMapKey(row: number, column: number): string {
  return `${row}:${column}`;
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
  private frameHandle: number | null = null;
  private frameUsesAnimationFrame = false;
  private disposed = false;

  constructor(options: CanvasRenderEngineOptions = {}) {
    this.skeletonModel = options.skeleton ?? new SheetSkeleton({ rowCount: 1000, columnCount: 26 });
    this.viewport = new Viewport(options.viewport);
    this.theme = mergeTheme(options.theme);
    this.layerDefinitions = (options.layers ?? DEFAULT_LAYER_DEFINITIONS).map((definition) => ({ ...definition }));
    this.scene = new Scene(this.layerDefinitions);
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
  }

  destroy(): void {
    this.dispose();
  }

  resize(width: number, height: number, devicePixelRatio = this.viewport.devicePixelRatio): void {
    this.assertActive();
    this.viewport.setSize(width, height, devicePixelRatio);
    if (this.scene.mounted) this.scene.resize({ width, height }, devicePixelRatio);
    this.forceFullRedraw = true;
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
  }

  scrollTo(x: number, y: number): void {
    this.assertActive();
    this.viewport.scrollTo(x, y, this.skeletonModel.contentSize);
  }

  scrollBy(deltaX: number, deltaY: number): void {
    this.assertActive();
    this.viewport.scrollBy(deltaX, deltaY, this.skeletonModel.contentSize);
  }

  setSkeleton(skeleton: SheetSkeleton): void {
    this.assertActive();
    this.skeletonModel = skeleton;
    this.viewport.clampTo(skeleton.contentSize);
    this.forceFullRedraw = true;
  }

  setCellProvider(cellProvider: CellProvider): void {
    this.assertActive();
    this.cellProvider = cellProvider;
    this.forceFullRedraw = true;
  }

  setCells(cells: ReadonlyMap<string, CellRenderData>): void {
    this.setCellProvider(createMapProvider(cells));
  }

  setTheme(theme: Partial<RenderTheme>): void {
    this.assertActive();
    this.theme = mergeTheme({ ...this.theme, ...theme });
    this.forceFullRedraw = true;
  }

  invalidate(ranges?: readonly CellRange[]): void {
    this.assertActive();
    if (ranges === undefined) {
      this.forceFullRedraw = true;
      return;
    }
    this.dirtyRanges.addMany(ranges);
  }

  markDirty(ranges?: readonly CellRange[]): void {
    this.invalidate(ranges);
  }

  render(): RenderPlan {
    this.assertActive();
    const viewport = this.viewport.getSnapshot();
    const plan = calculateRenderPlan({
      skeleton: this.skeletonModel,
      viewport,
      previousViewport: this.previousViewport,
      dirtyRanges: this.dirtyRanges.toArray(),
      forceFull: this.forceFullRedraw,
      layers: this.layerDefinitions,
    });
    if (this.scene.mounted) this.applyPlan(plan);
    this.previousViewport = viewport;
    this.lastPlan = plan;
    this.forceFullRedraw = false;
    this.dirtyRanges.clear();
    return plan;
  }

  requestRender(): void {
    this.assertActive();
    if (!this.scene.mounted || this.frameHandle !== null) return;
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
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
      if (!layer || layerPlan.mode === 'none') continue;
      if (layerPlan.mode === 'full') {
        layer.clear();
        this.drawLayer(layerPlan.layerId, plan);
        continue;
      }
      if (layerPlan.mode === 'scroll') layer.blit(plan.scrollDelta);
      layer.clear(layerPlan.clearRects);
      this.drawLayer(layerPlan.layerId, plan, layerPlan.drawRects);
    }
  }

  private drawLayer(layerId: string, plan: RenderPlan, drawRects?: readonly Rect[]): void {
    const layer = this.scene.getLayer(layerId);
    if (!layer) return;
    layer.withLogicalContext((context) => {
      const drawOptions = {
        context,
        skeleton: this.skeletonModel,
        viewport: plan.viewport,
        visibleRange: plan.visibleRange,
        cellProvider: this.cellProvider,
        theme: this.theme,
        drawRects,
      };
      if (layerId === 'grid') drawGridLayer(drawOptions);
      else if (layerId === 'content') drawCellLayer(drawOptions);
    });
  }

  private cancelScheduledRender(): void {
    if (this.frameHandle === null) return;
    if (this.frameUsesAnimationFrame && typeof window !== 'undefined') {
      window.cancelAnimationFrame(this.frameHandle);
    } else {
      clearTimeout(this.frameHandle as unknown as ReturnType<typeof setTimeout>);
    }
    this.frameHandle = null;
  }

  private assertActive(): void {
    if (this.disposed) throw new Error('CanvasRenderEngine has been disposed');
  }

  private reviveIfDisposed(): void {
    if (!this.disposed) return;
    this.scene = new Scene(this.layerDefinitions);
    this.disposed = false;
    this.lastPlan = null;
  }
}
