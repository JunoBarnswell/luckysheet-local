import type { DrawingObject, DrawingTransform, WorksheetModel } from '@react-sheets/core-model';
import { isWorksheetSnapSettings } from '@react-sheets/core-model';

export type DrawingSelectionMode = 'replace' | 'add' | 'toggle';

export interface DrawingHitTestResult {
  drawing: DrawingObject;
  handle?: 'move' | 'nw' | 'ne' | 'sw' | 'se' | 'n' | 's' | 'e' | 'w';
}

export interface DrawingPointerTransaction {
  id: string;
  sheetId: string;
  drawingId: string;
  before: DrawingTransform;
  preview: DrawingTransform;
}

export interface DrawingReconciliation {
  /** Selection that remains after removing IDs absent from the worksheet. */
  selection: string[];
  /** Pointer gestures cancelled because their drawing no longer exists. */
  cancelledPointerTransactionIds: string[];
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

  /**
   * Reconcile transient drawing state against the canonical worksheet drawing
   * collection. The workbook model is the existence authority; this runtime
   * never keeps a selected ID or pointer gesture for an absent drawing.
   */
  reconcile(sheetId: string, validDrawingIds: Iterable<string>): DrawingReconciliation {
    const valid = new Set(validDrawingIds);
    const current = this.selectedBySheet.get(sheetId);
    if (current) {
      for (const id of current) if (!valid.has(id)) current.delete(id);
      if (current.size === 0) this.selectedBySheet.delete(sheetId);
    }

    const cancelledPointerTransactionIds: string[] = [];
    for (const [transactionId, transaction] of this.pointerTransactions) {
      if (transaction.sheetId !== sheetId || valid.has(transaction.drawingId)) continue;
      this.pointerTransactions.delete(transactionId);
      cancelledPointerTransactionIds.push(transactionId);
    }

    return {
      selection: [...(this.selectedBySheet.get(sheetId) ?? [])],
      cancelledPointerTransactionIds,
    };
  }

  /** Clear transient state belonging to sheets removed from the workbook. */
  clearMissingSheets(liveSheetIds: Iterable<string>): string[] {
    const live = new Set(liveSheetIds);
    const clearedSheetIds: string[] = [];
    for (const sheetId of this.selectedBySheet.keys()) {
      if (live.has(sheetId)) continue;
      this.selectedBySheet.delete(sheetId);
      clearedSheetIds.push(sheetId);
    }
    for (const [transactionId, transaction] of this.pointerTransactions) {
      if (live.has(transaction.sheetId)) continue;
      this.pointerTransactions.delete(transactionId);
    }
    return clearedSheetIds;
  }

  /** Begin a transient pointer gesture without changing the worksheet. */
  beginPointerTransform(sheet: WorksheetModel, drawingId: string): DrawingPointerTransaction {
    const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
    if (!drawing) throw new Error(`Unknown drawing: ${drawingId}`);
    const transaction: DrawingPointerTransaction = {
      id: `drawing-pointer-${++pointerTransactionSequence}`,
      sheetId: sheet.id,
      drawingId,
      before: structuredClone(drawing.transform),
      preview: structuredClone(drawing.transform),
    };
    this.pointerTransactions.set(transaction.id, transaction);
    return structuredClone(transaction);
  }

  /** Update only the transient preview; callers commit it through drawing.transform.commit. */
  previewPointerTransform(sheet: WorksheetModel, transactionId: string, transform: DrawingTransform): DrawingTransform {
    const transaction = this.pointerTransactions.get(transactionId);
    if (!transaction) throw new Error(`Unknown drawing pointer transaction: ${transactionId}`);
    if (transaction.sheetId !== sheet.id) throw new Error(`Drawing pointer worksheet mismatch: ${transactionId}`);
    if (!isWorksheetSnapSettings(sheet.snapSettings)) throw new Error(`Worksheet snap settings are invalid: ${sheet.id}`);
    const settings = sheet.snapSettings;
    const gridSize = settings.enabled && settings.snapToGrid ? settings.gridSize : undefined;
    transaction.preview = snapTransform(normalizeTransform(transform), gridSize);
    if (settings.enabled && settings.snapToShape) transaction.preview = snapToShape(sheet, transaction.drawingId, transaction.preview);
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

export function snapTransform(transform: DrawingTransform, gridSize?: number): DrawingTransform {
  const resolvedGridSize = gridSize;
  if (resolvedGridSize === undefined || !Number.isFinite(resolvedGridSize) || resolvedGridSize <= 0) return structuredClone(transform);
  const snap = (value: number): number => Math.round(value / resolvedGridSize) * resolvedGridSize;
  return {
    ...transform,
    x: snap(transform.x),
    y: snap(transform.y),
    width: Math.max(resolvedGridSize, snap(transform.width)),
    height: Math.max(resolvedGridSize, snap(transform.height)),
  };
}

/** Snap moving/resizing bounds to nearby shape edges and centers in worksheet coordinates. */
function snapToShape(sheet: WorksheetModel, drawingId: string, transform: DrawingTransform): DrawingTransform {
  const threshold = 8;
  const others = sheet.drawings.filter((drawing) => drawing.id !== drawingId);
  const xCandidates = others.flatMap((drawing) => [drawing.transform.x, drawing.transform.x + drawing.transform.width / 2, drawing.transform.x + drawing.transform.width]);
  const yCandidates = others.flatMap((drawing) => [drawing.transform.y, drawing.transform.y + drawing.transform.height / 2, drawing.transform.y + drawing.transform.height]);
  const nearest = (value: number, candidates: readonly number[]) => {
    let result = value;
    let distance = threshold + 1;
    for (const candidate of candidates) {
      const next = Math.abs(candidate - value);
      if (next < distance) { result = candidate; distance = next; }
    }
    return result;
  };
  const left = nearest(transform.x, xCandidates);
  const top = nearest(transform.y, yCandidates);
  const right = nearest(transform.x + transform.width, xCandidates);
  const bottom = nearest(transform.y + transform.height, yCandidates);
  const width = right !== transform.x + transform.width ? Math.max(0, right - left) : transform.width;
  const height = bottom !== transform.y + transform.height ? Math.max(0, bottom - top) : transform.height;
  return { ...transform, x: left, y: top, width, height };
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
