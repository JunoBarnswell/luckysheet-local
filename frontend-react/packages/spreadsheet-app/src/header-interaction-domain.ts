import type { RangeRef } from '@react-sheets/core-model';
import type { SelectionState } from './selection-service';
import { selectionFromGesture } from './selection-interaction-machine';

export type HeaderTarget =
  | { kind: 'column'; index: number }
  | { kind: 'row'; index: number }
  | { kind: 'corner' };

export type HeaderIntent =
  | { type: 'select'; target: HeaderTarget; additive: boolean; extend: boolean }
  | { type: 'resize.begin'; target: Exclude<HeaderTarget, { kind: 'corner' }>; pointer: number }
  | { type: 'resize.update'; pointer: number }
  | { type: 'resize.commit' }
  | { type: 'autofit'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'hide'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'unhide'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'insert'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'delete'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'clear'; target: Exclude<HeaderTarget, { kind: 'corner' }> }
  | { type: 'move'; target: Exclude<HeaderTarget, { kind: 'corner' }>; destination: number; copy: boolean; insert: boolean };

export interface HeaderBounds {
  rowCount: number;
  columnCount: number;
}

export type HeaderContextAction = 'cut' | 'copy' | 'paste' | 'paste-special' | 'insert' | 'delete' | 'clear' | 'format' | 'size' | 'hide' | 'unhide';

export interface HeaderContextMenuDescriptor {
  id: string;
  action: HeaderContextAction;
  label: string;
  danger?: boolean;
  separator?: boolean;
}

export function headerRange(target: HeaderTarget, sheetId: string, bounds: HeaderBounds): RangeRef | null {
  if (target.kind === 'corner') return { sheetId, startRow: 0, endRow: Math.max(0, bounds.rowCount - 1), startColumn: 0, endColumn: Math.max(0, bounds.columnCount - 1) };
  if (target.kind === 'column') {
    if (!Number.isInteger(target.index) || target.index < 0 || target.index >= bounds.columnCount) return null;
    return { sheetId, startRow: 0, endRow: Math.max(0, bounds.rowCount - 1), startColumn: target.index, endColumn: target.index };
  }
  if (!Number.isInteger(target.index) || target.index < 0 || target.index >= bounds.rowCount) return null;
  return { sheetId, startRow: target.index, endRow: target.index, startColumn: 0, endColumn: Math.max(0, bounds.columnCount - 1) };
}

export function applyHeaderSelection(state: SelectionState, target: HeaderTarget, sheetId: string, bounds: HeaderBounds, modifiers: { additive: boolean; extend: boolean }): SelectionState {
  const range = headerRange(target, sheetId, bounds);
  if (!range || target.kind === 'corner') return range ? { ...state, ranges: [range], primaryRangeIndex: 0, activeCell: { row: 0, column: 0 }, anchorCell: { row: 0, column: 0 }, selectionKind: 'cells' } : state;
  const anchor = modifiers.extend
    ? target.kind === 'column' ? { row: 0, column: state.anchorCell.column } : { row: state.anchorCell.row, column: 0 }
    : target.kind === 'column' ? { row: 0, column: target.index } : { row: target.index, column: 0 };
  const gesture = target.kind === 'column'
    ? { origin: anchor, target: { row: 0, column: target.index }, kind: 'columns' as const, additive: modifiers.additive, expandedRange: modifiers.extend ? { ...range, startColumn: Math.min(anchor.column, target.index), endColumn: Math.max(anchor.column, target.index) } : range }
    : { origin: anchor, target: { row: target.index, column: 0 }, kind: 'rows' as const, additive: modifiers.additive, expandedRange: modifiers.extend ? { ...range, startRow: Math.min(anchor.row, target.index), endRow: Math.max(anchor.row, target.index) } : range };
  return selectionFromGesture(state, gesture, sheetId);
}

export function headerTargetSelected(state: SelectionState, target: Exclude<HeaderTarget, { kind: 'corner' }>, bounds: HeaderBounds): boolean {
  return state.ranges.some((range) => target.kind === 'column'
    ? range.startRow === 0 && range.endRow >= bounds.rowCount - 1 && range.startColumn <= target.index && range.endColumn >= target.index
    : range.startColumn === 0 && range.endColumn >= bounds.columnCount - 1 && range.startRow <= target.index && range.endRow >= target.index);
}

export interface DimensionSelectionOptions {
  /**
   * Ribbon dimension commands apply to the columns/rows covered by an ordinary
   * cell selection. Header context menus leave this disabled so a cell range
   * never turns into an implicit whole-dimension command.
   */
  includeOrdinaryCellRanges?: boolean;
}

export function selectedHeaderIndices(
  state: SelectionState,
  kind: 'column' | 'row',
  bounds: HeaderBounds,
  options: DimensionSelectionOptions = {},
): number[] {
  const result = new Set<number>();
  for (const range of state.ranges) {
    const full = kind === 'column'
      ? range.startRow === 0 && range.endRow >= bounds.rowCount - 1
      : range.startColumn === 0 && range.endColumn >= bounds.columnCount - 1;
    if (!full && !options.includeOrdinaryCellRanges) continue;
    const start = kind === 'column' ? range.startColumn : range.startRow;
    const end = kind === 'column' ? range.endColumn : range.endRow;
    const max = kind === 'column' ? bounds.columnCount - 1 : bounds.rowCount - 1;
    for (let index = Math.max(0, start); index <= Math.min(max, end); index += 1) result.add(index);
  }
  if (!result.size) result.add(kind === 'column' ? state.activeCell.column : state.activeCell.row);
  return [...result].sort((left, right) => left - right);
}

export function headerContextMenuCatalog(kind: 'column' | 'row'): readonly HeaderContextMenuDescriptor[] {
  const dimension = kind === 'column' ? 'Column' : 'Row';
  const dimensions = kind === 'column' ? 'Columns' : 'Rows';
  return [
    { id: `${kind}-cut`, action: 'cut', label: 'Cut' },
    { id: `${kind}-copy`, action: 'copy', label: 'Copy' },
    { id: `${kind}-paste`, action: 'paste', label: 'Paste' },
    { id: `${kind}-paste-special`, action: 'paste-special', label: 'Paste Special…' },
    { id: `${kind}-separator-1`, action: 'clear', label: '', separator: true },
    { id: `${kind}-insert`, action: 'insert', label: `Insert ${dimensions}` },
    { id: `${kind}-delete`, action: 'delete', label: `Delete ${dimensions}`, danger: true },
    { id: `${kind}-clear`, action: 'clear', label: 'Clear Contents' },
    { id: `${kind}-format`, action: 'format', label: 'Format Cells…' },
    { id: `${kind}-size`, action: 'size', label: `${dimension} ${kind === 'column' ? 'Width' : 'Height'}…` },
    { id: `${kind}-separator-2`, action: 'clear', label: '', separator: true },
    { id: `${kind}-hide`, action: 'hide', label: `Hide ${dimensions}` },
    { id: `${kind}-unhide`, action: 'unhide', label: `Unhide ${dimensions}` },
  ];
}
