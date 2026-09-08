import type { SpillRange, SpillState } from './spill';
/** Rust-produced spill projection; the host never solves collisions or clips it. */
export interface ResolvedSpill {
  sheetId: string;
  anchor: { row: number; column: number };
  range: SpillRange['range'];
  values: SpillRange['values'];
  state: SpillState;
  blocker?: { row: number; column: number };
}
export function isSpillChild(spill: SpillRange, row: number, column: number): boolean {
  return !(row === spill.anchor.row && column === spill.anchor.column)
    && row >= spill.range.startRow && row <= spill.range.endRow
    && column >= spill.range.startColumn && column <= spill.range.endColumn;
}
