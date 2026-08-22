import { Layer } from './layer';
import { DEFAULT_LAYER_DEFINITIONS, type LayerDefinition, type Size } from './types';

export class Scene {
  private readonly definitions: LayerDefinition[];
  private readonly layerMap = new Map<string, Layer>();
  private host: HTMLElement | null = null;
  private previousHostStyle: { position: string; overflow: string } | null = null;
  private currentSize: Size = { width: 0, height: 0 };
  private currentPixelRatio = 1;

  constructor(definitions: readonly LayerDefinition[] = DEFAULT_LAYER_DEFINITIONS) {
    this.definitions = definitions.map((definition) => ({ ...definition }));
    for (const definition of this.definitions) {
      if (this.layerMap.has(definition.id)) throw new Error(`Duplicate canvas layer: ${definition.id}`);
      this.layerMap.set(definition.id, new Layer(definition));
    }
  }

  get mounted(): boolean {
    return this.host !== null;
  }

  get layers(): Layer[] {
    return this.definitions
      .map((definition) => this.layerMap.get(definition.id))
      .filter((layer): layer is Layer => layer !== undefined);
  }

  get layerDefinitions(): LayerDefinition[] {
    return this.definitions.map((definition) => ({ ...definition }));
  }

  get currentHost(): HTMLElement | null {
    return this.host;
  }

  mount(nextHost: HTMLElement): void {
    if (this.host === nextHost && this.mounted) return;
    if (this.host && this.host !== nextHost) this.unmount();
    this.host = nextHost;
    this.previousHostStyle = {
      position: nextHost.style.position,
      overflow: nextHost.style.overflow,
    };
    nextHost.style.position = nextHost.style.position || 'relative';
    nextHost.style.overflow = nextHost.style.overflow || 'hidden';
    for (const layer of this.layers) layer.mount(nextHost);
    this.resize(this.currentSize, this.currentPixelRatio);
  }

  resize(size: Size, devicePixelRatio: number): void {
    this.currentSize = { ...size };
    this.currentPixelRatio = devicePixelRatio;
    for (const layer of this.layers) layer.resize(size, devicePixelRatio);
  }

  getLayer(id: string): Layer | undefined {
    return this.layerMap.get(id);
  }

  unmount(): void {
    for (const layer of this.layers) layer.unmount();
    if (this.host && this.previousHostStyle) {
      this.host.style.position = this.previousHostStyle.position;
      this.host.style.overflow = this.previousHostStyle.overflow;
    }
    this.previousHostStyle = null;
    this.host = null;
  }

  dispose(): void {
    this.unmount();
    for (const layer of this.layers) layer.dispose();
    this.layerMap.clear();
  }
}
