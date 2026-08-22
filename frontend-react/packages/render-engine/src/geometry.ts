import type { Rect } from './types';

export function rectRight(rect: Rect): number {
  return rect.x + rect.width;
}

export function rectBottom(rect: Rect): number {
  return rect.y + rect.height;
}

export function isRectUsable(rect: Rect): boolean {
  return Number.isFinite(rect.x)
    && Number.isFinite(rect.y)
    && Number.isFinite(rect.width)
    && Number.isFinite(rect.height)
    && rect.width > 0
    && rect.height > 0;
}

export function intersectRect(first: Rect, second: Rect): Rect | null {
  const x = Math.max(first.x, second.x);
  const y = Math.max(first.y, second.y);
  const right = Math.min(rectRight(first), rectRight(second));
  const bottom = Math.min(rectBottom(first), rectBottom(second));
  if (right <= x || bottom <= y) return null;
  return { x, y, width: right - x, height: bottom - y };
}

export function translateRect(rect: Rect, offset: { x: number; y: number }): Rect {
  return { ...rect, x: rect.x + offset.x, y: rect.y + offset.y };
}

export function rectsTouch(first: Rect, second: Rect): boolean {
  return rectRight(first) >= second.x
    && rectRight(second) >= first.x
    && rectBottom(first) >= second.y
    && rectBottom(second) >= first.y;
}

export function mergeRects(rects: readonly Rect[]): Rect[] {
  const merged: Rect[] = [];
  for (const rect of rects) {
    if (!isRectUsable(rect)) continue;
    let candidate = { ...rect };
    let mergedAnother = true;
    while (mergedAnother) {
      mergedAnother = false;
      for (let index = merged.length - 1; index >= 0; index -= 1) {
        const existing = merged[index];
        if (!existing || !rectsTouch(existing, candidate)) continue;
        candidate = {
          x: Math.min(candidate.x, existing.x),
          y: Math.min(candidate.y, existing.y),
          width: Math.max(rectRight(candidate), rectRight(existing)) - Math.min(candidate.x, existing.x),
          height: Math.max(rectBottom(candidate), rectBottom(existing)) - Math.min(candidate.y, existing.y),
        };
        merged.splice(index, 1);
        mergedAnother = true;
      }
    }
    merged.push(candidate);
  }
  return merged.sort((first, second) => first.y - second.y || first.x - second.x);
}
