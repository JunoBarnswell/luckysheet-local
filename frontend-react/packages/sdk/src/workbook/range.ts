import { cellAddress } from '@react-sheets/core-model';
import type { Cell } from './cell';
import type { Worksheet } from './worksheet';

export class Range {
  readonly address: string;
  readonly #sheet: Worksheet;
  readonly #start: { row: number; column: number };
  readonly #end: { row: number; column: number };
  constructor(sheet: Worksheet, start: { row: number; column: number }, end: { row: number; column: number }) { this.#sheet = sheet; this.#start = start; this.#end = end; this.address = `${cellAddress(start.row, start.column)}:${cellAddress(end.row, end.column)}`; Object.freeze(this); }
  *cells(): IterableIterator<Cell> {
    for (let row = this.#start.row; row <= this.#end.row; row++) for (let column = this.#start.column; column <= this.#end.column; column++) yield this.#sheet.cells.get(cellAddress(row, column));
  }
}
