import { normalizeRangeRef } from '@react-sheets/sheet-features';
import type { RangeRef } from '@react-sheets/core-model';
import type { SelectionArea, SelectionCell, SelectionKind, SelectionState } from './selection-service';

export interface SelectionBounds {
  rowCount: number;
  columnCount: number;
  hiddenRows?: readonly number[];
  hiddenColumns?: readonly number[];
}

export interface SelectionGesture {
  origin: SelectionCell;
  target: SelectionCell;
  pointerId?: number;
  kind?: SelectionKind;
  additive?: boolean;
  expandedRange?: RangeRef;
  bounds?: SelectionBounds;
}

export type SelectionInteractionEvent =
  | { type: 'pointer.commit'; gesture: SelectionGesture; sheetId: string }
  | { type: 'keyboard.move'; rowDelta: number; columnDelta: number; extend: boolean; bounds: SelectionBounds }
  | { type: 'context.target'; target: SelectionCell; sheetId: string }
  | { type: 'formula-reference.enter' }
  | { type: 'formula-reference.exit' };

/** Pure reducer for every selection gesture. Canvas owns only pointer input. */
export function reduceSelectionInteraction(state: SelectionState, event: SelectionInteractionEvent): SelectionState {
  switch (event.type) {
    case 'pointer.commit':
      return selectionFromGesture(state, event.gesture, event.sheetId);
    case 'keyboard.move':
      return moveSelection(state, event.rowDelta, event.columnDelta, event.extend, event.bounds);
    case 'context.target':
      return selectionFromGesture(state, { origin: event.target, target: event.target, kind: 'cells' }, event.sheetId);
    case 'formula-reference.enter':
      return { ...state, mode: 'formulaReference' };
    case 'formula-reference.exit':
      return { ...state, mode: 'normal' };
  }
}

export function selectionFromGesture(state: SelectionState, gesture: SelectionGesture, sheetId: string): SelectionState {
  const kind = gesture.kind ?? 'cells';
  const range = gesture.expandedRange ?? rangeFromCells(gesture.origin, gesture.target, kind, sheetId, gesture.bounds);
  return {
    ...state,
    ranges: gesture.additive ? [...state.ranges, range] : [range],
    primaryRangeIndex: gesture.additive ? state.ranges.length : 0,
    activeCell: { ...gesture.target },
    anchorCell: { ...gesture.origin },
    selectionKind: kind,
    mode: state.mode ?? 'normal',
  };
}

export function moveSelection(state: SelectionState, rowDelta: number, columnDelta: number, extend: boolean, bounds: SelectionBounds): SelectionState {
  const target = {
    row: skipHidden(clamp(state.activeCell.row + rowDelta, bounds.rowCount), Math.sign(rowDelta), bounds.hiddenRows, bounds.rowCount),
    column: skipHidden(clamp(state.activeCell.column + columnDelta, bounds.columnCount), Math.sign(columnDelta), bounds.hiddenColumns, bounds.columnCount),
  };
  if (!extend) {
    const sheetId = state.ranges[state.primaryRangeIndex]?.sheetId ?? '';
    return { ...state, ranges: [cellRange(sheetId, target)], primaryRangeIndex: 0, activeCell: target, anchorCell: target, selectionKind: 'cells', mode: state.mode ?? 'normal' };
  }
  const ranges = state.ranges.map((range) => ({ ...range }));
  const primary = ranges[state.primaryRangeIndex];
  if (!primary) return state;
  ranges[state.primaryRangeIndex] = rangeFromCells(state.anchorCell, target, 'cells', primary.sheetId);
  return { ...state, ranges, activeCell: target, selectionKind: 'cells', mode: state.mode ?? 'normal' };
}

export function selectionArea(state: SelectionState): SelectionArea[] {
  return state.ranges.map((range) => ({ kind: state.selectionKind ?? 'cells', range: { ...range } }));
}

function rangeFromCells(origin: SelectionCell, target: SelectionCell, kind: SelectionKind, sheetId: string, bounds?: SelectionBounds): RangeRef {
  return normalizeRangeRef({
    sheetId,
    startRow: kind === 'columns' ? 0 : Math.min(origin.row, target.row),
    endRow: kind === 'columns' ? Math.max(0, (bounds?.rowCount ?? Math.max(origin.row, target.row) + 1) - 1) : Math.max(origin.row, target.row),
    startColumn: kind === 'rows' ? 0 : Math.min(origin.column, target.column),
    endColumn: kind === 'rows' ? Math.max(0, (bounds?.columnCount ?? Math.max(origin.column, target.column) + 1) - 1) : Math.max(origin.column, target.column),
  });
}

function cellRange(sheetId: string, cell: SelectionCell): RangeRef {
  return { sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column };
}

function clamp(value: number, count: number): number {
  return Math.max(0, Math.min(Math.max(0, count - 1), Math.trunc(value)));
}

function skipHidden(value: number, direction: number, hidden: readonly number[] | undefined, count: number): number {
  if (!hidden || direction === 0) return value;
  const hiddenSet = new Set(hidden);
  let next = value;
  while (hiddenSet.has(next)) {
    const candidate = next + direction;
    if (candidate < 0 || candidate >= count) break;
    next = candidate;
  }
  return next;
}
