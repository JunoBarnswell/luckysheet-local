import type { RangeRef, SheetId, UnitId } from '@react-sheets/core-model';
import { normalizeRangeRef } from '@react-sheets/sheet-features';
import { cellAddress } from './address';
import { moveSelection, reduceSelectionInteraction } from './selection-interaction-machine';

export interface SelectionCell {
  row: number;
  column: number;
}

export type SelectionKind = 'cells' | 'rows' | 'columns' | 'sheet';
export type SelectionMode = 'normal' | 'formulaReference';
export interface SelectionArea {
  kind: SelectionKind;
  range: RangeRef;
}

/**
 * The only UI selection state shape. primaryRange is derived from
 * ranges[primaryRangeIndex] so the state cannot contain two competing range
 * values.
 */
export interface SelectionState {
  ranges: RangeRef[];
  primaryRangeIndex: number;
  activeCell: SelectionCell;
  anchorCell: SelectionCell;
  selectionKind?: SelectionKind;
  mode?: SelectionMode;
}

export interface SelectionSnapshot {
  unitId: UnitId;
  sheetId: SheetId;
  ranges: RangeRef[];
  primaryRangeIndex: number;
  activeCell: SelectionCell;
  anchorCell: SelectionCell;
  selectionKind?: SelectionKind;
  mode?: SelectionMode;
}

export function createInitialSelection(sheetId: SheetId): SelectionState {
  const range = normalizeRangeRef({ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  return {
    ranges: [range],
    primaryRangeIndex: 0,
    activeCell: { row: 0, column: 0 },
    anchorCell: { row: 0, column: 0 },
    selectionKind: 'cells',
    mode: 'normal',
  };
}

export class SelectionService {
  private state: SelectionState;

  constructor(
    private readonly unitId: UnitId,
    private readonly getActiveSheetId: () => SheetId,
    private readonly getSheetBounds: () => { rowCount: number; columnCount: number; hiddenRows?: readonly number[]; hiddenColumns?: readonly number[] },
    initial?: SelectionState,
  ) {
    this.state = this.normalizeState(initial ?? createInitialSelection(getActiveSheetId()));
  }

  getSnapshot(): SelectionSnapshot {
    const state = this.getState();
    return {
      unitId: this.unitId,
      sheetId: this.getActiveSheetId(),
      ranges: state.ranges.map((range) => ({ ...range })),
      primaryRangeIndex: state.primaryRangeIndex,
      activeCell: { ...state.activeCell },
      anchorCell: { ...state.anchorCell },
      selectionKind: state.selectionKind,
      mode: state.mode,
    };
  }

  getState(): SelectionState {
    return {
      ranges: this.state.ranges.map((range) => ({ ...range })),
      primaryRangeIndex: this.state.primaryRangeIndex,
      activeCell: { ...this.state.activeCell },
      anchorCell: { ...this.state.anchorCell },
      selectionKind: this.state.selectionKind,
      mode: this.state.mode,
    };
  }

  get activeCell(): string {
    return cellAddress(this.state.activeCell.row, this.state.activeCell.column);
  }

  resetForSheet(sheetId: SheetId): void {
    this.state = this.normalizeState(createInitialSelection(sheetId));
  }

  selectCell(address: string, opts?: { insertRef?: (ref: string) => void; editing?: boolean }): boolean {
    if (opts?.editing && opts.insertRef) {
      this.setInteractionMode('formulaReference');
      const parsed = parseAddressForSelection(address);
      if (parsed) opts.insertRef(`${columnLabel(parsed.column)}${parsed.row + 1}`);
      return false;
    }
    const parsed = parseAddressForSelection(address);
    if (!parsed) return false;
    this.selectCellAt(parsed.row, parsed.column);
    return true;
  }

  selectCellAt(row: number, column: number): void {
    const nextRow = this.clampRow(row);
    const nextColumn = this.clampColumn(column);
    const sheetId = this.getActiveSheetId();
    this.applyState({
      ranges: [normalizeRangeRef({ sheetId, startRow: nextRow, endRow: nextRow, startColumn: nextColumn, endColumn: nextColumn })],
      primaryRangeIndex: 0,
      activeCell: { row: nextRow, column: nextColumn },
      anchorCell: { row: nextRow, column: nextColumn },
      selectionKind: 'cells',
      mode: this.state.mode ?? 'normal',
    });
  }

  selectRange(
    range: { startRow: number; startColumn: number; endRow: number; endColumn: number },
    mode: 'replace' | 'add' | 'extend' = 'replace',
    release?: SelectionCell,
  ): void {
    const normalized = this.normalizeRange(range);
    if (mode === 'extend') {
      const extended = this.rangeFromCells(this.state.anchorCell, {
        row: normalized.endRow,
        column: normalized.endColumn,
      });
      this.applyState({
        ranges: [extended],
        primaryRangeIndex: 0,
        activeCell: { row: normalized.endRow, column: normalized.endColumn },
        anchorCell: { ...this.state.anchorCell },
        selectionKind: 'cells',
        mode: this.state.mode ?? 'normal',
      });
      return;
    }
    if (mode === 'add' && this.state.ranges.length > 0) {
      this.applyState({
        ranges: [...this.state.ranges, normalized],
        primaryRangeIndex: this.state.ranges.length,
        activeCell: release ? { ...release } : { row: normalized.startRow, column: normalized.startColumn },
        anchorCell: release ? { ...this.state.anchorCell } : { row: normalized.startRow, column: normalized.startColumn },
        selectionKind: 'cells',
        mode: this.state.mode ?? 'normal',
      });
      return;
    }
    this.applyState({
      ranges: [normalized],
      primaryRangeIndex: 0,
      activeCell: { row: normalized.startRow, column: normalized.startColumn },
      anchorCell: release ? { ...release } : { row: normalized.startRow, column: normalized.startColumn },
      selectionKind: 'cells',
      mode: this.state.mode ?? 'normal',
    });
  }

  selectRow(row: number, columnCount: number): void {
    const targetRow = this.clampRow(row);
    this.applyState({
      ranges: [this.normalizeRange({ startRow: targetRow, endRow: targetRow, startColumn: 0, endColumn: columnCount - 1 })],
      primaryRangeIndex: 0,
      activeCell: { row: targetRow, column: 0 },
      anchorCell: { row: targetRow, column: 0 },
      selectionKind: 'rows',
      mode: this.state.mode ?? 'normal',
    });
  }

  selectColumn(column: number, rowCount: number): void {
    const targetColumn = this.clampColumn(column);
    this.applyState({
      ranges: [this.normalizeRange({ startRow: 0, endRow: rowCount - 1, startColumn: targetColumn, endColumn: targetColumn })],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: targetColumn },
      anchorCell: { row: 0, column: targetColumn },
      selectionKind: 'columns',
      mode: this.state.mode ?? 'normal',
    });
  }

  /** Apply the complete state in one operation. Empty input is ignored. */
  applyState(selection: SelectionState): void {
    if (selection.ranges.length === 0) return;
    this.state = this.normalizeState(selection);
  }

  setInteractionMode(mode: SelectionMode): void {
    this.applyState(reduceSelectionInteraction(this.state, { type: mode === 'formulaReference' ? 'formula-reference.enter' : 'formula-reference.exit' }));
  }

  movePrimary(rowDelta: number, columnDelta: number, opts?: { extend?: boolean }): void {
    const bounds = this.getSheetBounds();
    this.applyState(moveSelection(this.state, rowDelta, columnDelta, Boolean(opts?.extend), bounds));
  }

  selectAll(rowCount: number, columnCount: number): void {
    const endRow = Math.max(0, rowCount - 1);
    const endColumn = Math.max(0, columnCount - 1);
    this.applyState({
      ranges: [this.normalizeRange({ startRow: 0, endRow, startColumn: 0, endColumn })],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
      selectionKind: 'sheet',
      mode: this.state.mode ?? 'normal',
    });
  }

  primaryRangeOrDefault(): RangeRef {
    const primary = this.state.ranges[this.state.primaryRangeIndex];
    if (primary) return { ...primary };
    const { row, column } = this.state.activeCell;
    return this.normalizeRange({ startRow: row, endRow: row, startColumn: column, endColumn: column });
  }

  private normalizeState(selection: SelectionState): SelectionState {
    const ranges = selection.ranges.map((range) => this.normalizeRange(range));
    const primaryRangeIndex = Math.max(0, Math.min(ranges.length - 1, Math.trunc(selection.primaryRangeIndex)));
    const primaryRange = ranges[primaryRangeIndex]!;
    const activeCell = this.clampCellToRange(selection.activeCell, primaryRange);
    const anchorCell = this.clampCellToRange(selection.anchorCell, primaryRange);
    return {
      ranges,
      primaryRangeIndex,
      activeCell,
      anchorCell,
      selectionKind: selection.selectionKind ?? this.state.selectionKind ?? 'cells',
      mode: selection.mode ?? this.state.mode ?? 'normal',
    };
  }

  private normalizeRange(range: { sheetId?: SheetId; startRow: number; startColumn: number; endRow: number; endColumn: number }): RangeRef {
    const sheetId = range.sheetId ?? this.getActiveSheetId();
    return normalizeRangeRef({
      sheetId,
      startRow: this.clampRow(Math.min(range.startRow, range.endRow)),
      endRow: this.clampRow(Math.max(range.startRow, range.endRow)),
      startColumn: this.clampColumn(Math.min(range.startColumn, range.endColumn)),
      endColumn: this.clampColumn(Math.max(range.startColumn, range.endColumn)),
    });
  }

  private rangeFromCells(first: SelectionCell, second: SelectionCell): RangeRef {
    return this.normalizeRange({
      startRow: Math.min(first.row, second.row),
      endRow: Math.max(first.row, second.row),
      startColumn: Math.min(first.column, second.column),
      endColumn: Math.max(first.column, second.column),
    });
  }

  private clampCellToRange(cell: SelectionCell, range: RangeRef): SelectionCell {
    return {
      row: Math.max(range.startRow, Math.min(range.endRow, this.clampRow(cell.row))),
      column: Math.max(range.startColumn, Math.min(range.endColumn, this.clampColumn(cell.column))),
    };
  }

  private clampRow(row: number): number {
    const { rowCount } = this.getSheetBounds();
    return Math.max(0, Math.min(Math.max(0, rowCount - 1), Math.trunc(Number.isFinite(row) ? row : 0)));
  }

  private clampColumn(column: number): number {
    const { columnCount } = this.getSheetBounds();
    return Math.max(0, Math.min(Math.max(0, columnCount - 1), Math.trunc(Number.isFinite(column) ? column : 0)));
  }
}

export function parseRangeReference(input: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } | undefined {
  const parseCell = (value: string) => {
    const match = /^([A-Z]+)(\d+)$/.exec(value.trim().toUpperCase());
    if (!match?.[1] || !match[2]) return undefined;
    const column = [...match[1]].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    return Number.isInteger(row) && row >= 0 && column >= 0 ? { row, column } : undefined;
  };
  const [startText, endText = startText] = input.split(':');
  const start = parseCell(startText ?? '');
  const end = parseCell(endText ?? '');
  if (!start || !end) return undefined;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  };
}

function parseAddressForSelection(address: string): SelectionCell | undefined {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  return { column: column - 1, row: Number(match[2]) - 1 };
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}
