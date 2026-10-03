import type { Workbook } from './workbook';
import type { Worksheet } from './worksheet';
import { Range } from './range';
import { parseCellAddress } from './cell-address';
import { domainFor } from './object-domain';

export class RangeCollection {
  readonly #workbook: Workbook;
  readonly #sheet: Worksheet;
  constructor(workbook: Workbook, sheet: Worksheet) { this.#workbook = workbook; this.#sheet = sheet; Object.freeze(this); }
  get(input: string): Range {
    try {
      if (typeof input !== 'string') throw new Error('Range address must be text');
      const parts = input.split(':');
      if (parts.length < 1 || parts.length > 2) throw new Error('Range must be one rectangular A1 address');
      const start = parseCellAddress(parts[0]!), end = parseCellAddress(parts.at(-1)!);
      if (end.row < start.row || end.column < start.column) throw new Error('Range endpoints must be ordered');
      domainFor(this.#workbook).sheets();
      this.#sheet.name;
      return new Range(this.#workbook, this.#sheet, start, end);
    } catch (cause) { return domainFor(this.#workbook).invalid('ranges.get', cause, { sheetId: this.#sheet.id, address: input }); }
  }
}
