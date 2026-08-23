import type { DrawingObject, DrawingTransform, WorksheetModel } from '@react-sheets/core-model';

export type DrawingSelectionMode = 'replace' | 'add' | 'toggle';

export interface DrawingHitTestResult {
  drawing: DrawingObject;
  handle?: 'move' | 'nw' | 'ne' | 'sw' | 'se' | 'n' | 's' | 'e' | 'w';
}

export interface DrawingPointerTransaction {
  id: string;
  drawingId: string;
  before: DrawingTransform;
  preview: DrawingTransform;
}

let pointerTransactionSequence = 0;

export class DrawingRuntime {
  private readonly selectedBySheet = new Map<string, Set<string>>();
  private readonly pointerTransactions = new Map<string, DrawingPointerTransaction>();

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

  /** Begin a transient pointer gesture without changing the worksheet. */
  beginPointerTransform(sheet: WorksheetModel, drawingId: string): DrawingPointerTransaction {
    const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
    if (!drawing) throw new Error(`Unknown drawing: ${drawingId}`);
    const transaction: DrawingPointerTransaction = {
      id: `drawing-pointer-${++pointerTransactionSequence}`,
      drawingId,
      before: structuredClone(drawing.transform),
      preview: structuredClone(drawing.transform),
    };
    this.pointerTransactions.set(transaction.id, transaction);
    return structuredClone(transaction);
  }

  /** Update only the transient preview; callers commit it through drawing.transform.commit. */
  previewPointerTransform(transactionId: string, transform: DrawingTransform, grid = 1): DrawingTransform {
    const transaction = this.pointerTransactions.get(transactionId);
    if (!transaction) throw new Error(`Unknown drawing pointer transaction: ${transactionId}`);
    transaction.preview = snapTransform(normalizeTransform(transform), grid);
    return structuredClone(transaction.preview);
  }

  finishPointerTransform(transactionId: string): DrawingTransformCommit {
    const transaction = this.pointerTransactions.get(transactionId);
    if (!transaction) throw new Error(`Unknown drawing pointer transaction: ${transactionId}`);
    this.pointerTransactions.delete(transactionId);
    return {
      drawingId: transaction.drawingId,
      before: structuredClone(transaction.before),
      after: structuredClone(transaction.preview),
    };
  }

  cancelPointerTransform(transactionId: string): void {
    this.pointerTransactions.delete(transactionId);
  }

  hitTest(sheet: WorksheetModel, point: { x: number; y: number }): DrawingHitTestResult | undefined {
    const sorted = [...sheet.drawings].sort((left, right) => right.zIndex - left.zIndex);
    for (const drawing of sorted) {
      const localPoint = inverseRotatePoint(point, drawing.transform);
      const { x, y, width, height } = drawing.transform;
      if (localPoint.x < x || localPoint.y < y || localPoint.x > x + width || localPoint.y > y + height) continue;
      const handle = this.resolveHandle(drawing.transform, localPoint);
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

export interface DrawingTransformCommit {
  drawingId: string;
  before: DrawingTransform;
  after: DrawingTransform;
}

export function normalizeTransform(transform: DrawingTransform): DrawingTransform {
  const width = Math.max(0, transform.width);
  const height = Math.max(0, transform.height);
  const rotation = transform.rotation ?? 0;
  return {
    x: Number.isFinite(transform.x) ? transform.x : 0,
    y: Number.isFinite(transform.y) ? transform.y : 0,
    width,
    height,
    rotation: Number.isFinite(rotation) ? rotation : 0,
  };
}

export function snapTransform(transform: DrawingTransform, gridSize = 1): DrawingTransform {
  if (!Number.isFinite(gridSize) || gridSize <= 0) return structuredClone(transform);
  const snap = (value: number): number => Math.round(value / gridSize) * gridSize;
  return {
    ...transform,
    x: snap(transform.x),
    y: snap(transform.y),
    width: Math.max(gridSize, snap(transform.width)),
    height: Math.max(gridSize, snap(transform.height)),
  };
}

function inverseRotatePoint(point: { x: number; y: number }, transform: DrawingTransform): { x: number; y: number } {
  const rotation = (transform.rotation ?? 0) * Math.PI / 180;
  if (rotation === 0) return point;
  const centerX = transform.x + transform.width / 2;
  const centerY = transform.y + transform.height / 2;
  const dx = point.x - centerX;
  const dy = point.y - centerY;
  const cos = Math.cos(-rotation);
  const sin = Math.sin(-rotation);
  return {
    x: centerX + dx * cos - dy * sin,
    y: centerY + dx * sin + dy * cos,
  };
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
