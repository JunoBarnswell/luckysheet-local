import type { RangeRef, SheetId, UnitId } from '@react-sheets/core-model';
import { normalizeRangeRef } from '@react-sheets/sheet-features';
import { cellAddress } from './address';

/** UI-facing selection state (canvas + formula bar) */
export interface SelectionState {
  ranges: RangeRef[];
  primaryRowIndex: number;
  primaryColumnIndex: number;
  primaryRangeIndex: number;
  anchorRowIndex: number;
  anchorColumnIndex: number;
}

type SelectionStateCore = Omit<SelectionState, 'anchorRowIndex' | 'anchorColumnIndex'>;

export interface SelectionSnapshot {
  unitId: UnitId;
  sheetId: SheetId;
  ranges: RangeRef[];
  primaryRangeIndex: number;
  primaryRowIndex: number;
  primaryColumnIndex: number;
  anchorCell: { row: number; column: number };
}

export function createInitialSelection(sheetId: SheetId): SelectionStateCore {
  return {
    ranges: [normalizeRangeRef({ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })],
    primaryRowIndex: 0,
    primaryColumnIndex: 0,
    primaryRangeIndex: 0,
  };
}

export class SelectionService {
  private state: SelectionStateCore;
  private anchorCell: { row: number; column: number };

  constructor(
    private readonly unitId: UnitId,
    private readonly getActiveSheetId: () => SheetId,
    private readonly getSheetBounds: () => { rowCount: number; columnCount: number },
    initial?: SelectionStateCore,
  ) {
    this.state = initial ?? createInitialSelection(getActiveSheetId());
    this.anchorCell = { row: this.state.primaryRowIndex, column: this.state.primaryColumnIndex };
  }

  getSnapshot(): SelectionSnapshot {
    return {
      unitId: this.unitId,
      sheetId: this.getActiveSheetId(),
      ranges: this.state.ranges.map((range) => ({ ...range })),
      primaryRangeIndex: this.state.primaryRangeIndex,
      primaryRowIndex: this.state.primaryRowIndex,
      primaryColumnIndex: this.state.primaryColumnIndex,
      anchorCell: { ...this.anchorCell },
    };
  }

  getState(): SelectionState {
    return {
      ...this.state,
      anchorRowIndex: this.anchorCell.row,
      anchorColumnIndex: this.anchorCell.column,
    };
  }

  get activeCell(): string {
    return cellAddress(this.state.primaryRowIndex, this.state.primaryColumnIndex);
  }

  resetForSheet(sheetId: SheetId): void {
    this.state = {
      ranges: [],
      primaryRowIndex: 0,
      primaryColumnIndex: 0,
      primaryRangeIndex: 0,
    };
    this.anchorCell = { row: 0, column: 0 };
    if (sheetId) {
      this.setPrimary(0, 0, sheetId);
    }
  }

  selectCell(address: string, opts?: { insertRef?: (ref: string) => void; editing?: boolean }): boolean {
    if (opts?.editing && opts.insertRef) {
      const parsed = parseAddressForSelection(address);
      if (parsed) opts.insertRef(`${columnLabel(parsed.column)}${parsed.row + 1}`);
      return false;
    }
    const parsed = parseAddressForSelection(address);
    if (!parsed) return false;
    this.setPrimary(this.clampRow(parsed.row), this.clampColumn(parsed.column));
    return true;
  }

  selectRange(
    range: { startRow: number; startColumn: number; endRow: number; endColumn: number },
    mode: 'replace' | 'add' | 'extend' = 'replace',
  ): void {
    const sheetId = this.getActiveSheetId();
    const startRow = this.clampRow(Math.min(range.startRow, range.endRow));
    const endRow = this.clampRow(Math.max(range.startRow, range.endRow));
    const startColumn = this.clampColumn(Math.min(range.startColumn, range.endColumn));
    const endColumn = this.clampColumn(Math.max(range.startColumn, range.endColumn));
    const normalized = normalizeRangeRef({ sheetId, startRow, endRow, startColumn, endColumn });
    if (mode === 'extend') {
      const anchorRow = this.anchorCell.row;
      const anchorColumn = this.anchorCell.column;
      const extended = normalizeRangeRef({
        sheetId,
        startRow: Math.min(anchorRow, startRow, endRow),
        endRow: Math.max(anchorRow, startRow, endRow),
        startColumn: Math.min(anchorColumn, startColumn, endColumn),
        endColumn: Math.max(anchorColumn, startColumn, endColumn),
      });
      this.state = {
        ranges: [extended],
        primaryRowIndex: endRow,
        primaryColumnIndex: endColumn,
        primaryRangeIndex: 0,
      };
      return;
    }
    if (mode === 'add' && this.state.ranges.length > 0) {
      this.state = {
        ...this.state,
        ranges: [...this.state.ranges, normalized],
        primaryRangeIndex: this.state.ranges.length,
        primaryRowIndex: normalized.startRow,
        primaryColumnIndex: normalized.startColumn,
      };
    } else {
      this.state = {
        ranges: [normalized],
        primaryRowIndex: normalized.startRow,
        primaryColumnIndex: normalized.startColumn,
        primaryRangeIndex: 0,
      };
      this.anchorCell = { row: normalized.startRow, column: normalized.startColumn };
    }
  }

  applyFromRanges(ranges: RangeRef[]): void {
    if (ranges.length === 0) return;
    const sheetId = this.getActiveSheetId();
    const normalized = ranges.map((range) =>
      normalizeRangeRef({
        sheetId: range.sheetId ?? sheetId,
        startRow: this.clampRow(Math.min(range.startRow, range.endRow)),
        endRow: this.clampRow(Math.max(range.startRow, range.endRow)),
        startColumn: this.clampColumn(Math.min(range.startColumn, range.endColumn)),
        endColumn: this.clampColumn(Math.max(range.startColumn, range.endColumn)),
      }),
    );
    const primary = normalized[0]!;
    this.state = {
      ranges: normalized,
      primaryRowIndex: primary.startRow,
      primaryColumnIndex: primary.startColumn,
      primaryRangeIndex: 0,
    };
    this.anchorCell = { row: primary.startRow, column: primary.startColumn };
  }

  /**
   * Apply the complete transient canvas selection without collapsing its
   * active cell to the top-left of a dragged range.  This is essential for
   * edit placement: a range may start at A2 while its active cell is C9.
   */
  applyState(selection: SelectionState): void {
    if (selection.ranges.length === 0) return;
    const sheetId = this.getActiveSheetId();
    const ranges = selection.ranges.map((range) => normalizeRangeRef({
      sheetId: range.sheetId ?? sheetId,
      startRow: this.clampRow(Math.min(range.startRow, range.endRow)),
      endRow: this.clampRow(Math.max(range.startRow, range.endRow)),
      startColumn: this.clampColumn(Math.min(range.startColumn, range.endColumn)),
      endColumn: this.clampColumn(Math.max(range.startColumn, range.endColumn)),
    }));
    this.state = {
      ranges,
      primaryRangeIndex: Math.max(0, Math.min(ranges.length - 1, selection.primaryRangeIndex)),
      primaryRowIndex: this.clampRow(selection.primaryRowIndex),
      primaryColumnIndex: this.clampColumn(selection.primaryColumnIndex),
    };
    this.anchorCell = {
      row: this.clampRow(selection.anchorRowIndex),
      column: this.clampColumn(selection.anchorColumnIndex),
    };
  }

  movePrimary(rowDelta: number, columnDelta: number, opts?: { extend?: boolean }): void {
    const sheetId = this.getActiveSheetId();
    const targetRow = this.clampRow(this.state.primaryRowIndex + rowDelta);
    const targetColumn = this.clampColumn(this.state.primaryColumnIndex + columnDelta);
    if (opts?.extend && this.state.ranges.length > 0) {
      const anchorRow = this.anchorCell.row;
      const anchorColumn = this.anchorCell.column;
      const next = normalizeRangeRef({
        sheetId,
        startRow: Math.min(anchorRow, targetRow),
        endRow: Math.max(anchorRow, targetRow),
        startColumn: Math.min(anchorColumn, targetColumn),
        endColumn: Math.max(anchorColumn, targetColumn),
      });
      const ranges = [...this.state.ranges];
      ranges[this.state.primaryRangeIndex] = next;
      this.state = { ...this.state, ranges, primaryRowIndex: targetRow, primaryColumnIndex: targetColumn };
      return;
    }
    this.setPrimary(targetRow, targetColumn, sheetId);
  }

  setPrimary(row: number, column: number, sheetId?: SheetId, opts?: { preserveAnchor?: boolean }): void {
    const sid = sheetId ?? this.getActiveSheetId();
    this.state = {
      ranges: [normalizeRangeRef({ sheetId: sid, startRow: row, endRow: row, startColumn: column, endColumn: column })],
      primaryRowIndex: row,
      primaryColumnIndex: column,
      primaryRangeIndex: 0,
    };
    if (!opts?.preserveAnchor) {
      this.anchorCell = { row, column };
    }
  }

  setPrimaryCell(row: number, column: number): void {
    this.state = {
      ...this.state,
      primaryRowIndex: this.clampRow(row),
      primaryColumnIndex: this.clampColumn(column),
    };
  }

  setAnchor(row: number, column: number): void {
    this.anchorCell = { row: this.clampRow(row), column: this.clampColumn(column) };
  }

  selectAll(rowCount: number, columnCount: number): void {
    this.selectRange({ startRow: 0, startColumn: 0, endRow: rowCount - 1, endColumn: columnCount - 1 }, 'replace');
  }

  primaryRangeOrDefault(): RangeRef {
    const sheetId = this.getActiveSheetId();
    return (
      this.state.ranges[this.state.primaryRangeIndex] ??
      normalizeRangeRef({
        sheetId,
        startRow: this.state.primaryRowIndex,
        endRow: this.state.primaryRowIndex,
        startColumn: this.state.primaryColumnIndex,
        endColumn: this.state.primaryColumnIndex,
      })
    );
  }

  private clampRow(row: number): number {
    const { rowCount } = this.getSheetBounds();
    return Math.max(0, Math.min(rowCount - 1, row));
  }

  private clampColumn(column: number): number {
    const { columnCount } = this.getSheetBounds();
    return Math.max(0, Math.min(columnCount - 1, column));
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

function parseAddressForSelection(address: string): { column: number; row: number } | undefined {
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
