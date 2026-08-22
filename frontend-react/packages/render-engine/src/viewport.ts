import type { Point, Rect, Size, ViewportSnapshot } from './types';

function normalizeLength(value: number): number {
  return Number.isFinite(value) ? Math.max(0, value) : 0;
}

function normalizeDpr(value: number): number {
  return Number.isFinite(value) && value > 0 ? value : 1;
}

function defaultDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : normalizeDpr(window.devicePixelRatio);
}

export class Viewport {
  private current: ViewportSnapshot;

  constructor(initial: Partial<ViewportSnapshot> = {}) {
    this.current = {
      width: normalizeLength(initial.width ?? 0),
      height: normalizeLength(initial.height ?? 0),
      scrollX: normalizeLength(initial.scrollX ?? 0),
      scrollY: normalizeLength(initial.scrollY ?? 0),
      devicePixelRatio: normalizeDpr(initial.devicePixelRatio ?? defaultDevicePixelRatio()),
    };
  }

  get width(): number {
    return this.current.width;
  }

  get height(): number {
    return this.current.height;
  }

  get scrollX(): number {
    return this.current.scrollX;
  }

  get scrollY(): number {
    return this.current.scrollY;
  }

  get devicePixelRatio(): number {
    return this.current.devicePixelRatio;
  }

  getSnapshot(): ViewportSnapshot {
    return { ...this.current };
  }

  setSize(width: number, height: number, devicePixelRatio = this.devicePixelRatio): ViewportSnapshot {
    this.current = {
      ...this.current,
      width: normalizeLength(width),
      height: normalizeLength(height),
      devicePixelRatio: normalizeDpr(devicePixelRatio),
    };
    return this.getSnapshot();
  }

  setSnapshot(snapshot: Partial<ViewportSnapshot>): ViewportSnapshot {
    this.current = {
      width: normalizeLength(snapshot.width ?? this.width),
      height: normalizeLength(snapshot.height ?? this.height),
      scrollX: normalizeLength(snapshot.scrollX ?? this.scrollX),
      scrollY: normalizeLength(snapshot.scrollY ?? this.scrollY),
      devicePixelRatio: normalizeDpr(snapshot.devicePixelRatio ?? this.devicePixelRatio),
    };
    return this.getSnapshot();
  }

  scrollTo(x: number, y: number, contentSize?: Size): ViewportSnapshot {
    const maxX = contentSize ? Math.max(0, contentSize.width - this.width) : Number.POSITIVE_INFINITY;
    const maxY = contentSize ? Math.max(0, contentSize.height - this.height) : Number.POSITIVE_INFINITY;
    this.current = {
      ...this.current,
      scrollX: Math.min(maxX, normalizeLength(x)),
      scrollY: Math.min(maxY, normalizeLength(y)),
    };
    return this.getSnapshot();
  }

  setScroll(x: number, y: number, contentSize?: Size): ViewportSnapshot {
    return this.scrollTo(x, y, contentSize);
  }

  scrollBy(deltaX: number, deltaY: number, contentSize?: Size): ViewportSnapshot {
    return this.scrollTo(this.scrollX + deltaX, this.scrollY + deltaY, contentSize);
  }

  clampTo(contentSize: Size): ViewportSnapshot {
    return this.scrollTo(this.scrollX, this.scrollY, contentSize);
  }

  toSheetRect(): Rect {
    return { x: this.scrollX, y: this.scrollY, width: this.width, height: this.height };
  }

  toViewportPoint(point: Point): Point {
    return { x: point.x - this.scrollX, y: point.y - this.scrollY };
  }

  toSheetPoint(point: Point): Point {
    return { x: point.x + this.scrollX, y: point.y + this.scrollY };
  }
}
