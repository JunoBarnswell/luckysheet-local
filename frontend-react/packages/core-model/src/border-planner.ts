import type { CellBorderSide, CellBorders, RangeRef } from './index';

/** Border placements are topology, not a pre-expanded style object. */
export type BorderPlacement =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'none'
  | 'all'
  | 'outside'
  | 'thick-outside'
  | 'inside-horizontal'
  | 'inside-vertical';

export interface BorderLine {
  style: CellBorderSide['style'];
  color: string;
}

export interface BorderPlanCell {
  row: number;
  column: number;
  /** Only the sides owned by this cell are returned. */
  sides: CellBorders;
}

export interface BorderPlan {
  range: RangeRef;
  placement: BorderPlacement;
  line?: BorderLine;
  cells: readonly BorderPlanCell[];
}

export interface BorderPlanBounds {
  rowCount: number;
  columnCount: number;
}

const PLACEMENTS: readonly BorderPlacement[] = [
  'bottom', 'top', 'left', 'right', 'none', 'all', 'outside', 'thick-outside', 'inside-horizontal', 'inside-vertical',
];

function isPlacement(value: unknown): value is BorderPlacement {
  return typeof value === 'string' && (PLACEMENTS as readonly string[]).includes(value);
}

function isLineStyle(value: unknown): value is CellBorderSide['style'] {
  return value === 'hair' || value === 'thin' || value === 'medium' || value === 'thick' || value === 'dotted' || value === 'dashed' || value === 'dashDot' || value === 'dashDotDot' || value === 'double';
}

function validateLine(line: BorderLine | undefined, placement: BorderPlacement): BorderLine | undefined {
  if (placement === 'none') return undefined;
  if (!line || !isLineStyle(line.style) || typeof line.color !== 'string' || line.color.trim().length === 0 || line.color.length > 64) {
    throw new Error('Border line must contain a supported style and non-empty color');
  }
  return { style: line.style, color: line.color };
}

function validateBounds(bounds: BorderPlanBounds): void {
  if (!Number.isSafeInteger(bounds.rowCount) || bounds.rowCount <= 0 || !Number.isSafeInteger(bounds.columnCount) || bounds.columnCount <= 0) {
    throw new Error('Border planner requires positive worksheet dimensions');
  }
}

function validateRange(range: RangeRef, bounds: BorderPlanBounds): RangeRef {
  if (!range || typeof range.sheetId !== 'string' || range.sheetId.length === 0
    || !Number.isSafeInteger(range.startRow) || !Number.isSafeInteger(range.endRow)
    || !Number.isSafeInteger(range.startColumn) || !Number.isSafeInteger(range.endColumn)
    || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow >= bounds.rowCount || range.endColumn >= bounds.columnCount) {
    throw new Error('Border range is outside worksheet bounds');
  }
  return {
    sheetId: range.sheetId,
    startRow: range.startRow,
    endRow: range.endRow,
    startColumn: range.startColumn,
    endColumn: range.endColumn,
  };
}

function sidesFor(
  placement: BorderPlacement,
  row: number,
  column: number,
  range: RangeRef,
  line: BorderLine | undefined,
): CellBorders {
  if (placement === 'none') return {};
  const borderLine = line as BorderLine;
  if (placement === 'all') return { top: borderLine, right: borderLine, bottom: borderLine, left: borderLine };
  if (placement === 'bottom') return { bottom: borderLine };
  if (placement === 'top') return { top: borderLine };
  if (placement === 'left') return { left: borderLine };
  if (placement === 'right') return { right: borderLine };
  if (placement === 'outside' || placement === 'thick-outside') {
    const perimeterLine = placement === 'thick-outside' ? { ...borderLine, style: 'thick' as const } : borderLine;
    return {
      ...(row === range.startRow ? { top: perimeterLine } : {}),
      ...(row === range.endRow ? { bottom: perimeterLine } : {}),
      ...(column === range.startColumn ? { left: perimeterLine } : {}),
      ...(column === range.endColumn ? { right: perimeterLine } : {}),
    };
  }
  if (placement === 'inside-horizontal') return row > range.startRow ? { top: borderLine } : {};
  if (placement === 'inside-vertical') return column > range.startColumn ? { left: borderLine } : {};
  throw new Error(`Unsupported border placement: ${placement}`);
}

/**
 * Pure border topology planner. It does not read or mutate cells; callers can
 * merge each returned side patch with the existing cell style atomically.
 * Internal edges have one canonical owner (top/left of the lower/right cell),
 * so All and Outside never accidentally create duplicate interior semantics.
 */
export function planBorderChange(
  range: RangeRef,
  placement: BorderPlacement,
  line: BorderLine | undefined,
  bounds: BorderPlanBounds,
): BorderPlan {
  validateBounds(bounds);
  if (!isPlacement(placement)) throw new Error('Unsupported border placement');
  const normalizedRange = validateRange(range, bounds);
  const normalizedLine = validateLine(line, placement);
  const cells: BorderPlanCell[] = [];
  for (let row = normalizedRange.startRow; row <= normalizedRange.endRow; row += 1) {
    for (let column = normalizedRange.startColumn; column <= normalizedRange.endColumn; column += 1) {
      cells.push({ row, column, sides: sidesFor(placement, row, column, normalizedRange, normalizedLine) });
    }
  }
  return { range: normalizedRange, placement, line: normalizedLine, cells };
}

export function isBorderPlacement(value: unknown): value is BorderPlacement {
  return isPlacement(value);
}

export function isBorderLine(value: unknown): value is BorderLine {
  if (!value || typeof value !== 'object') return false;
  const line = value as Record<string, unknown>;
  return isLineStyle(line.style) && typeof line.color === 'string' && line.color.trim().length > 0 && line.color.length <= 64;
}
