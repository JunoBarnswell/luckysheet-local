import type { ScrollDeltaPlan } from './render-plan';
import type { LayerDefinition, Rect, Size } from './types';

export class Layer {
  readonly id: string;
  readonly definition: LayerDefinition;

  private canvas: HTMLCanvasElement | null = null;
  private context: CanvasRenderingContext2D | null = null;
  private host: HTMLElement | null = null;
  private logicalSize: Size = { width: 0, height: 0 };
  private pixelRatio = 1;

  constructor(definition: LayerDefinition) {
    this.id = definition.id;
    this.definition = { ...definition };
  }

  get mounted(): boolean {
    return this.canvas !== null && this.host !== null;
  }

  get element(): HTMLCanvasElement | null {
    return this.canvas;
  }

  get renderingContext(): CanvasRenderingContext2D | null {
    return this.context;
  }

  get size(): Size {
    return { ...this.logicalSize };
  }

  get devicePixelRatio(): number {
    return this.pixelRatio;
  }

  mount(nextHost: HTMLElement): HTMLCanvasElement {
    if (this.host && this.host !== nextHost) this.unmount();
    this.host = nextHost;
    if (!this.canvas) {
      this.canvas = nextHost.ownerDocument.createElement('canvas');
      this.canvas.dataset.renderLayer = this.id;
      this.context = this.canvas.getContext('2d');
    }
    if (this.canvas.parentElement !== nextHost) nextHost.appendChild(this.canvas);
    this.canvas.style.position = 'absolute';
    this.canvas.style.inset = '0';
    this.canvas.style.display = 'block';
    this.canvas.style.zIndex = String(this.definition.zIndex);
    this.canvas.style.pointerEvents = this.definition.pointerEvents ?? 'none';
    if (this.definition.opacity !== undefined) this.canvas.style.opacity = String(this.definition.opacity);
    this.resize(this.logicalSize, this.pixelRatio);
    return this.canvas;
  }

  resize(size: Size, devicePixelRatio: number): void {
    this.logicalSize = {
      width: Math.max(0, Number.isFinite(size.width) ? size.width : 0),
      height: Math.max(0, Number.isFinite(size.height) ? size.height : 0),
    };
    this.pixelRatio = Number.isFinite(devicePixelRatio) && devicePixelRatio > 0 ? devicePixelRatio : 1;
    if (!this.canvas) return;
    this.canvas.width = Math.ceil(this.logicalSize.width * this.pixelRatio);
    this.canvas.height = Math.ceil(this.logicalSize.height * this.pixelRatio);
    this.canvas.style.width = `${this.logicalSize.width}px`;
    this.canvas.style.height = `${this.logicalSize.height}px`;
    this.context = this.canvas.getContext('2d');
    this.context?.setTransform(this.pixelRatio, 0, 0, this.pixelRatio, 0, 0);
  }

  withLogicalContext(callback: (context: CanvasRenderingContext2D) => void): void {
    const context = this.context;
    if (!context) return;
    context.save();
    context.setTransform(this.pixelRatio, 0, 0, this.pixelRatio, 0, 0);
    try {
      callback(context);
    } finally {
      context.restore();
    }
  }

  clear(rects?: readonly Rect[]): void {
    const context = this.context;
    if (!context) return;
    context.save();
    context.setTransform(1, 0, 0, 1, 0, 0);
    if (!rects || rects.length === 0) {
      context.clearRect(0, 0, this.canvas?.width ?? 0, this.canvas?.height ?? 0);
    } else {
      for (const rect of rects) {
        context.clearRect(
          rect.x * this.pixelRatio,
          rect.y * this.pixelRatio,
          rect.width * this.pixelRatio,
          rect.height * this.pixelRatio,
        );
      }
    }
    context.restore();
  }

  blit(scrollDelta: ScrollDeltaPlan): void {
    const context = this.context;
    const source = scrollDelta.source;
    const destination = scrollDelta.destination;
    if (!context || !this.canvas || !scrollDelta.canBlit || !source || !destination) return;
    context.save();
    context.setTransform(1, 0, 0, 1, 0, 0);
    context.drawImage(
      this.canvas,
      source.x * this.pixelRatio,
      source.y * this.pixelRatio,
      source.width * this.pixelRatio,
      source.height * this.pixelRatio,
      destination.x * this.pixelRatio,
      destination.y * this.pixelRatio,
      destination.width * this.pixelRatio,
      destination.height * this.pixelRatio,
    );
    context.restore();
  }

  unmount(): void {
    if (this.canvas?.parentElement) this.canvas.parentElement.removeChild(this.canvas);
    this.host = null;
  }

  dispose(): void {
    this.unmount();
    this.context = null;
    this.canvas = null;
  }
}
