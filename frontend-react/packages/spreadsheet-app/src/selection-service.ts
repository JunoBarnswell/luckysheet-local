import type { RangeRef, SheetId, UnitId } from '@react-sheets/core-model';
import { normalizeRangeRef } from '@react-sheets/sheet-features';
import { cellAddress } from './address';

/** UI-facing selection state (canvas + formula bar) */
export interface SelectionState {
  ranges: RangeRef[];
  primaryRowIndex: number;
  primaryColumnIndex: number;
  primaryRangeIndex: number;
}

export interface SelectionSnapshot {
  unitId: UnitId;
  sheetId: SheetId;
  ranges: RangeRef[];
  primaryRangeIndex: number;
  primaryRowIndex: number;
  primaryColumnIndex: number;
  anchorCell: { row: number; column: number };
}

export function createInitialSelection(sheetId: SheetId): SelectionState {
  return {
    ranges: [normalizeRangeRef({ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })],
    primaryRowIndex: 0,
    primaryColumnIndex: 0,
    primaryRangeIndex: 0,
  };
}

export class SelectionService {
  private state: SelectionState;
  private anchorCell: { row: number; column: number };

  constructor(
    private readonly unitId: UnitId,
    private readonly getActiveSheetId: () => SheetId,
    private readonly getSheetBounds: () => { rowCount: number; columnCount: number },
    initial?: SelectionState,
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
    return this.state;
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
    mode: 'replace' | 'add' = 'replace',
  ): void {
    const sheetId = this.getActiveSheetId();
    const normalized = normalizeRangeRef({
      sheetId,
      startRow: this.clampRow(Math.min(range.startRow, range.endRow)),
      endRow: this.clampRow(Math.max(range.startRow, range.endRow)),
      startColumn: this.clampColumn(Math.min(range.startColumn, range.endColumn)),
      endColumn: this.clampColumn(Math.max(range.startColumn, range.endColumn)),
    });
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

  movePrimary(rowDelta: number, columnDelta: number, opts?: { extend?: boolean }): void {
    const sheetId = this.getActiveSheetId();
    const targetRow = this.clampRow(this.state.primaryRowIndex + rowDelta);
    const targetColumn = this.clampColumn(this.state.primaryColumnIndex + columnDelta);
    if (opts?.extend && this.state.ranges.length > 0) {
      const range = this.state.ranges[this.state.primaryRangeIndex] ?? this.state.ranges[0]!;
      const anchorRow = this.state.primaryRowIndex <= (range.startRow + range.endRow) / 2 ? range.endRow : range.startRow;
      const anchorColumn = this.state.primaryColumnIndex <= (range.startColumn + range.endColumn) / 2 ? range.endColumn : range.startColumn;
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

  setPrimary(row: number, column: number, sheetId?: SheetId): void {
    const sid = sheetId ?? this.getActiveSheetId();
    this.state = {
      ranges: [normalizeRangeRef({ sheetId: sid, startRow: row, endRow: row, startColumn: column, endColumn: column })],
      primaryRowIndex: row,
      primaryColumnIndex: column,
      primaryRangeIndex: 0,
    };
    this.anchorCell = { row, column };
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
