import type { DrawingObject, DrawingTransform, WorksheetModel } from '@react-sheets/core-model';

export type DrawingSelectionMode = 'replace' | 'add' | 'toggle';

export interface DrawingHitTestResult {
  drawing: DrawingObject;
  handle?: 'move' | 'nw' | 'ne' | 'sw' | 'se' | 'n' | 's' | 'e' | 'w';
}

export class DrawingRuntime {
  private readonly selectedBySheet = new Map<string, Set<string>>();

  select(sheetId: string, drawingIds: readonly string[], mode: DrawingSelectionMode = 'replace'): string[] {
    const current = this.selectedBySheet.get(sheetId) ?? new Set<string>();
    const next = mode === 'replace' ? new Set<string>() : new Set(current);
    for (const id of drawingIds) {
      if (mode === 'toggle' && next.has(id)) next.delete(id);
      else next.add(id);
    }
    this.selectedBySheet.set(sheetId, next);
    return [...next];
  }

  deselect(sheetId: string, drawingIds?: readonly string[]): string[] {
    if (!drawingIds || drawingIds.length === 0) {
      this.selectedBySheet.delete(sheetId);
      return [];
    }
    const current = this.selectedBySheet.get(sheetId);
    if (!current) return [];
    for (const id of drawingIds) current.delete(id);
    if (current.size === 0) this.selectedBySheet.delete(sheetId);
    return current.size ? [...current] : [];
  }

  getSelection(sheetId: string): readonly string[] {
    return [...(this.selectedBySheet.get(sheetId) ?? [])];
  }

  hitTest(sheet: WorksheetModel, point: { x: number; y: number }): DrawingHitTestResult | undefined {
    const sorted = [...sheet.drawings].sort((left, right) => right.zIndex - left.zIndex);
    for (const drawing of sorted) {
      const { x, y, width, height } = drawing.transform;
      if (point.x < x || point.y < y || point.x > x + width || point.y > y + height) continue;
      const handle = this.resolveHandle(drawing.transform, point);
      return { drawing, handle };
    }
    return undefined;
  }

  private resolveHandle(anchor: DrawingTransform, point: { x: number; y: number }): DrawingHitTestResult['handle'] {
    const threshold = 8;
    const { x, y, width, height } = anchor;
    const near = (left: number, top: number) => Math.abs(point.x - left) <= threshold && Math.abs(point.y - top) <= threshold;
    if (near(x, y)) return 'nw';
    if (near(x + width, y)) return 'ne';
    if (near(x, y + height)) return 'sw';
    if (near(x + width, y + height)) return 'se';
    if (Math.abs(point.y - y) <= threshold && point.x >= x && point.x <= x + width) return 'n';
    if (Math.abs(point.y - (y + height)) <= threshold && point.x >= x && point.x <= x + width) return 's';
    if (Math.abs(point.x - x) <= threshold && point.y >= y && point.y <= y + height) return 'w';
    if (Math.abs(point.x - (x + width)) <= threshold && point.y >= y && point.y <= y + height) return 'e';
    return 'move';
  }
}

export function nextZIndex(sheet: WorksheetModel): number {
  if (sheet.drawings.length === 0) return 1;
  return Math.max(...sheet.drawings.map((drawing) => drawing.zIndex)) + 1;
}

export function reorderDrawing(sheet: WorksheetModel, drawingId: string, direction: 'forward' | 'backward' | 'front' | 'back'): void {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  if (!drawing) return;
  const ordered = [...sheet.drawings].sort((left, right) => left.zIndex - right.zIndex);
  const index = ordered.findIndex((entry) => entry.id === drawingId);
  if (index < 0) return;
  if (direction === 'front') {
    const max = Math.max(...ordered.map((entry) => entry.zIndex));
    drawing.zIndex = max + 1;
    return;
  }
  if (direction === 'back') {
    const min = Math.min(...ordered.map((entry) => entry.zIndex));
    drawing.zIndex = min - 1;
    return;
  }
  const swapIndex = direction === 'forward' ? index + 1 : index - 1;
  const swap = ordered[swapIndex];
  if (!swap) return;
  const current = drawing.zIndex;
  drawing.zIndex = swap.zIndex;
  swap.zIndex = current;
}
