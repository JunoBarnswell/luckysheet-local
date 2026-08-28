import type { RangeRef } from '@react-sheets/core-model';

export type RangeDragMode = 'move-replace' | 'copy-replace' | 'move-insert' | 'copy-insert';

export interface RangeDragModifiers {
  copy: boolean;
  insert: boolean;
}

export interface RangeDragPlan {
  mode: RangeDragMode;
  sourceRange: RangeRef;
  targetRange: RangeRef;
  targetOrigin: { row: number; column: number };
}

export interface RangeDragBounds {
  rowCount: number;
  columnCount: number;
}

export function rangeDragMode(modifiers: RangeDragModifiers): RangeDragMode {
  return `${modifiers.copy ? 'copy' : 'move'}-${modifiers.insert ? 'insert' : 'replace'}` as RangeDragMode;
}

export function planRangeDrag(sourceRange: RangeRef, targetOrigin: { row: number; column: number }, mode: RangeDragMode, bounds: RangeDragBounds): RangeDragPlan {
  if (!Number.isSafeInteger(targetOrigin.row) || !Number.isSafeInteger(targetOrigin.column) || targetOrigin.row < 0 || targetOrigin.column < 0) throw new Error('RANGE_DRAG_INVALID_TARGET: target origin is invalid');
  const rowCount = sourceRange.endRow - sourceRange.startRow + 1;
  const columnCount = sourceRange.endColumn - sourceRange.startColumn + 1;
  const targetRange: RangeRef = {
    sheetId: sourceRange.sheetId,
    startRow: targetOrigin.row,
    endRow: targetOrigin.row + rowCount - 1,
    startColumn: targetOrigin.column,
    endColumn: targetOrigin.column + columnCount - 1,
  };
  if (targetRange.endRow >= bounds.rowCount || targetRange.endColumn >= bounds.columnCount) throw new Error('RANGE_DRAG_OUT_OF_BOUNDS: target range exceeds worksheet bounds');
  if (sourceRange.startRow === targetRange.startRow && sourceRange.startColumn === targetRange.startColumn) throw new Error('RANGE_DRAG_NOOP: source and target ranges are identical');
  return { mode, sourceRange: structuredClone(sourceRange), targetRange, targetOrigin: { ...targetOrigin } };
}

export function isRangeBorderPoint(point: { x: number; y: number }, rect: { x: number; y: number; width: number; height: number }, tolerance = 6): boolean {
  if (point.x < rect.x - tolerance || point.x > rect.x + rect.width + tolerance || point.y < rect.y - tolerance || point.y > rect.y + rect.height + tolerance) return false;
  const nearHorizontal = Math.abs(point.y - rect.y) <= tolerance || Math.abs(point.y - (rect.y + rect.height)) <= tolerance;
  const nearVertical = Math.abs(point.x - rect.x) <= tolerance || Math.abs(point.x - (rect.x + rect.width)) <= tolerance;
  return nearHorizontal || nearVertical;
}
