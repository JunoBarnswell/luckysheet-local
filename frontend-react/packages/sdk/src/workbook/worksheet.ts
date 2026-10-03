import type { WorksheetPane } from '@react-sheets/core-model';
import type { Workbook } from './workbook';
import type { WorksheetSnapshot } from './contract';
import { CellCollection } from './cell-collection';
import { RangeCollection } from './range-collection';
import { WorksheetAxis } from './worksheet-axis';
import { domainFor } from './object-domain';
import { immutableSnapshot } from './value';

export class Worksheet {
  readonly id: string;
  readonly cells: CellCollection;
  readonly ranges: RangeCollection;
  readonly rows: WorksheetAxis;
  readonly columns: WorksheetAxis;
  readonly #workbook: Workbook;
  constructor(workbook: Workbook, id: string) {
    this.#workbook = workbook; this.id = id;
    this.cells = new CellCollection(workbook, id); this.ranges = new RangeCollection(workbook, this);
    this.rows = new WorksheetAxis(workbook, this, 'rows'); this.columns = new WorksheetAxis(workbook, this, 'columns');
    Object.freeze(this);
  }
  get name(): string {
    const sheet = domainFor(this.#workbook).sheets().find(sheet => sheet.id === this.id);
    if (!sheet) return domainFor(this.#workbook).invalid('worksheet.name', new Error(`Worksheet no longer exists: ${this.id}`), { sheetId: this.id });
    return sheet.name;
  }
  snapshot(): WorksheetSnapshot {
    this.name;
    return immutableSnapshot(domainFor(this.#workbook).sheet(this.id));
  }

  async rename(name: string): Promise<void> { await this.#command('rename', 'sheet.rename', { name }); }
  async remove(): Promise<void> { this.name; await domainFor(this.#workbook).command('worksheet.remove', { commandId: 'sheet.remove', params: { id: this.id } }); }
  async reorder(toIndex: number): Promise<void> {
    if (!Number.isSafeInteger(toIndex) || toIndex < 0 || toIndex >= domainFor(this.#workbook).sheets().length) domainFor(this.#workbook).invalid('worksheet.reorder', new Error('Worksheet index is outside the workbook.'));
    await this.#command('reorder', 'sheet.reorder', { toIndex });
  }
  async duplicate(name: string): Promise<Worksheet> {
    this.name;
    const newId = globalThis.crypto.randomUUID();
    await domainFor(this.#workbook).command('worksheet.duplicate', { commandId: 'sheet.duplicate', params: { sourceSheetId: this.id, newId, newName: name } });
    return this.#workbook.worksheets.byId(newId);
  }
  async setHidden(hidden: boolean): Promise<void> {
    if (typeof hidden !== 'boolean') domainFor(this.#workbook).invalid('worksheet.setHidden', new Error('Visibility must be a boolean.'));
    await this.#command('setHidden', hidden ? 'sheet.hide' : 'sheet.unhide', {});
  }
  async setPane(pane: WorksheetPane): Promise<void> { await this.#command('setPane', 'sheet.freeze.set', { pane }); }
  async growExtent(rowCount: number, columnCount: number): Promise<void> { await this.#command('growExtent', 'sheet.extent.grow', { rowCount, columnCount }); }
  async #command(operation: string, commandId: string, params: object): Promise<void> {
    this.name;
    await domainFor(this.#workbook).command(`worksheet.${operation}`, { commandId, params: { ...params, sheetId: this.id } });
  }
}
