import type { Workbook } from './workbook';
import type { Worksheet } from './worksheet';
import { domainFor } from './object-domain';

/** Indices are zero-based, explicit, and never inferred from selection. */
export class WorksheetAxis {
  readonly #workbook: Workbook;
  readonly #sheet: Worksheet;
  readonly #axis: 'rows' | 'columns';
  constructor(workbook: Workbook, sheet: Worksheet, axis: 'rows' | 'columns') { this.#workbook = workbook; this.#sheet = sheet; this.#axis = axis; Object.freeze(this); }
  async insert(at: number, count = 1): Promise<void> { await this.#structure('insert', at, count); }
  async delete(at: number, count = 1): Promise<void> { await this.#structure('delete', at, count); }
  async #structure(action: 'insert' | 'delete', at: number, count: number): Promise<void> {
    const extent = this.#extent();
    if (!Number.isSafeInteger(at) || at < 0 || !Number.isSafeInteger(count) || count <= 0 || at > extent || action === 'delete' && (at + count > extent || count >= extent)) this.#invalid('structure', 'Axis operation exceeds the canonical worksheet extent.');
    await domainFor(this.#workbook).command(`worksheet.${this.#axis}.${action}`, { commandId: `sheet.${this.#axis}.${action}`, params: { sheetId: this.#sheet.id, at, count } });
  }
  async setPixels(indices: readonly number[], pixels: number): Promise<void> {
    this.#validateIndices(indices);
    if (!Number.isFinite(pixels) || pixels <= 0) this.#invalid('setPixels', 'Dimension must be a positive finite number of pixels.');
    const entries = indices.map(index => this.#axis === 'rows' ? { row: index, heightPx: pixels } : { column: index, widthPx: pixels });
    await domainFor(this.#workbook).command(`worksheet.${this.#axis}.setPixels`, { commandId: 'sheet.dimensions.apply', params: { sheetId: this.#sheet.id, [this.#axis]: entries } });
  }
  async setHidden(indices: readonly number[], hidden: boolean): Promise<void> {
    this.#validateIndices(indices);
    if (typeof hidden !== 'boolean') this.#invalid('setHidden', 'Visibility must be a boolean.');
    await domainFor(this.#workbook).command(`worksheet.${this.#axis}.setHidden`, { commandId: `sheet.${this.#axis}.visibility.set`, params: { sheetId: this.#sheet.id, [this.#axis]: [...indices], hidden } });
  }
  #extent(): number { const snapshot = this.#sheet.snapshot(); return this.#axis === 'rows' ? snapshot.rowCount : snapshot.columnCount; }
  #validateIndices(indices: readonly number[]): void {
    const extent = this.#extent();
    if (!Array.isArray(indices) || indices.length === 0 || indices.length > 10_000 || new Set(indices).size !== indices.length || !indices.every(index => Number.isSafeInteger(index) && index >= 0 && index < extent)) this.#invalid('indices', 'Indices must be unique, nonempty and within the canonical worksheet extent (maximum 10000).');
  }
  #invalid(operation: string, message: string): never { return domainFor(this.#workbook).invalid(`worksheet.${this.#axis}.${operation}`, new Error(message), { sheetId: this.#sheet.id }); }
}
