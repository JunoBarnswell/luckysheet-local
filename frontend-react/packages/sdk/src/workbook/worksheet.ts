import type { Workbook } from './workbook';
import { CellCollection } from './cell-collection';
import { RangeCollection } from './range-collection';
import { domainFor } from './object-domain';

export class Worksheet {
  readonly id: string;
  readonly cells: CellCollection;
  readonly ranges: RangeCollection;
  readonly #workbook: Workbook;
  constructor(workbook: Workbook, id: string) { this.#workbook = workbook; this.id = id; this.cells = new CellCollection(workbook, id); this.ranges = new RangeCollection(workbook, this); Object.freeze(this); }
  get name(): string {
    const sheet = domainFor(this.#workbook).sheets().find(sheet => sheet.id === this.id);
    if (!sheet) return domainFor(this.#workbook).invalid('worksheet.name', new Error(`Worksheet no longer exists: ${this.id}`), { sheetId: this.id });
    return sheet.name;
  }
}
