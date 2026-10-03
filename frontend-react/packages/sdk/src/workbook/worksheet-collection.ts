import type { Workbook } from './workbook';
import { Worksheet } from './worksheet';
import { domainFor } from './object-domain';

export class WorksheetCollection {
  readonly #workbook: Workbook;
  readonly #objects = new Map<string, Worksheet>();
  constructor(workbook: Workbook) { this.#workbook = workbook; Object.freeze(this); }
  list(): readonly Worksheet[] { return Object.freeze(domainFor(this.#workbook).sheets().map(sheet => this.byId(sheet.id))); }
  byId(id: string): Worksheet {
    if (!domainFor(this.#workbook).sheets().some(sheet => sheet.id === id)) return domainFor(this.#workbook).invalid('worksheets.byId', new Error(`Unknown sheet ID: ${id}`));
    let object = this.#objects.get(id);
    if (!object) { object = new Worksheet(this.#workbook, id); this.#objects.set(id, object); }
    return object;
  }
  byName(name: string): Worksheet {
    const found = typeof name === 'string' ? domainFor(this.#workbook).sheets().find(sheet => sheet.name.toUpperCase() === name.toUpperCase()) : undefined;
    if (!found) return domainFor(this.#workbook).invalid('worksheets.byName', new Error(`Unknown sheet name: ${String(name)}`));
    return this.byId(found.id);
  }
  at(index: number): Worksheet {
    const sheet = Number.isSafeInteger(index) && index >= 0 ? domainFor(this.#workbook).sheets()[index] : undefined;
    if (!sheet) return domainFor(this.#workbook).invalid('worksheets.at', new Error(`Unknown sheet index: ${index}`));
    return this.byId(sheet.id);
  }
}
