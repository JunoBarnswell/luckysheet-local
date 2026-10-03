import { cellAddress } from '@react-sheets/core-model';
import type { Workbook } from './workbook';
import { Cell } from './cell';
import { parseCellAddress } from './cell-address';
import { domainFor } from './object-domain';

export class CellCollection {
  readonly #workbook: Workbook;
  readonly #sheetId: string;
  readonly #cells = new Map<string, Cell>();
  constructor(workbook: Workbook, sheetId: string) { this.#workbook = workbook; this.#sheetId = sheetId; Object.freeze(this); }
  get(input: string): Cell {
    try {
      const parsed = parseCellAddress(input), key = cellAddress(parsed.row, parsed.column);
      domainFor(this.#workbook).sheets();
      let cell = this.#cells.get(key);
      if (!cell) { cell = new Cell(this.#workbook, this.#sheetId, parsed.row, parsed.column); this.#cells.set(key, cell); }
      return cell;
    } catch (cause) { return domainFor(this.#workbook).invalid('cells.get', cause, { sheetId: this.#sheetId, address: input }); }
  }
}
