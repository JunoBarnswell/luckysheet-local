import type { MergeSpan, RangeRef } from '@react-sheets/core-model';
import { normalizeRangeRef } from '@react-sheets/sheet-features';
import type { SelectionCell, SelectionKind } from './selection-service';

export interface SelectionTargetSurface {
  rowCount: number;
  columnCount: number;
  merges: readonly MergeSpan[];
  hiddenRows?: readonly number[];
  hiddenColumns?: readonly number[];
}

export interface ResolvedSelectionTarget {
  cell: SelectionCell;
  range: RangeRef;
}

/**
 * Resolves worksheet coordinates before they enter the selection state. The
 * render engine already resolves pane/frozen geometry; this boundary owns the
 * worksheet semantics that are independent of screen coordinates.
 */
export function resolveSelectionTarget(
  surface: SelectionTargetSurface,
  cell: SelectionCell,
  kind: SelectionKind = 'cells',
  sheetId = '',
): ResolvedSelectionTarget {
  const bounded = {
    row: clamp(cell.row, surface.rowCount),
    column: clamp(cell.column, surface.columnCount),
  };
  if (kind === 'rows') {
    return { cell: { row: bounded.row, column: 0 }, range: fullRowRange(sheetId, bounded.row, surface.columnCount) };
  }
  if (kind === 'columns') {
    return { cell: { row: 0, column: bounded.column }, range: fullColumnRange(sheetId, bounded.column, surface.rowCount) };
  }
  if (kind === 'sheet') {
    return { cell: { row: 0, column: 0 }, range: fullSheetRange(sheetId, surface.rowCount, surface.columnCount) };
  }
  const merge = surface.merges.find((candidate) => containsCell(candidate.range, bounded));
  return {
    cell: merge ? { ...merge.anchor } : bounded,
    range: normalizeRangeRef(merge?.range ?? { sheetId, startRow: bounded.row, endRow: bounded.row, startColumn: bounded.column, endColumn: bounded.column }),
  };
}

export function expandSelectionRangeForMerges(surface: Pick<SelectionTargetSurface, 'merges'>, range: RangeRef): RangeRef {
  let expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of surface.merges) {
      if (!intersectsRange(expanded, merge.range)) continue;
      const next = {
        startRow: Math.min(expanded.startRow, merge.range.startRow),
        endRow: Math.max(expanded.endRow, merge.range.endRow),
        startColumn: Math.min(expanded.startColumn, merge.range.startColumn),
        endColumn: Math.max(expanded.endColumn, merge.range.endColumn),
      };
      changed = next.startRow !== expanded.startRow || next.endRow !== expanded.endRow
        || next.startColumn !== expanded.startColumn || next.endColumn !== expanded.endColumn;
      expanded = { ...expanded, ...next };
    }
  }
  return expanded;
}

export function intersectsRange(
  first: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>,
  second: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>,
): boolean {
  return first.startRow <= second.endRow && second.startRow <= first.endRow
    && first.startColumn <= second.endColumn && second.startColumn <= first.endColumn;
}

export function containsRange(
  outer: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>,
  inner: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>,
): boolean {
  return outer.startRow <= inner.startRow && outer.endRow >= inner.endRow
    && outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn;
}

export function nextVisibleCell(surface: SelectionTargetSurface, cell: SelectionCell, rowDelta: number, columnDelta: number): SelectionCell {
  const hiddenRows = new Set(surface.hiddenRows ?? []);
  const hiddenColumns = new Set(surface.hiddenColumns ?? []);
  let row = clamp(cell.row, surface.rowCount);
  let column = clamp(cell.column, surface.columnCount);
  do {
    row += rowDelta;
    column += columnDelta;
  } while (row >= 0 && row < surface.rowCount && column >= 0 && column < surface.columnCount
    && (hiddenRows.has(row) || hiddenColumns.has(column)));
  return {
    row: clamp(row, surface.rowCount),
    column: clamp(column, surface.columnCount),
  };
}

function containsCell(range: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>, cell: SelectionCell): boolean {
  return cell.row >= range.startRow && cell.row <= range.endRow
    && cell.column >= range.startColumn && cell.column <= range.endColumn;
}

function fullRowRange(sheetId: string, row: number, columnCount: number): RangeRef {
  return normalizeRangeRef({ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: Math.max(0, columnCount - 1) });
}

function fullColumnRange(sheetId: string, column: number, rowCount: number): RangeRef {
  return normalizeRangeRef({ sheetId, startRow: 0, endRow: Math.max(0, rowCount - 1), startColumn: column, endColumn: column });
}

function fullSheetRange(sheetId: string, rowCount: number, columnCount: number): RangeRef {
  return normalizeRangeRef({ sheetId, startRow: 0, endRow: Math.max(0, rowCount - 1), startColumn: 0, endColumn: Math.max(0, columnCount - 1) });
}

function clamp(value: number, count: number): number {
  return Math.max(0, Math.min(Math.max(0, count - 1), Math.trunc(Number.isFinite(value) ? value : 0)));
}
