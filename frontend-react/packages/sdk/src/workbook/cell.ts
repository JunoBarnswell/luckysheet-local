import { cellAddress, type CellValue, type RichTextRun, type CellStyle, type BorderPlacement, type BorderLine } from '@react-sheets/core-model';
import type { Workbook } from './workbook';
import type { CellSnapshot, RangeStyleOptions } from './contract';
import { domainFor } from './object-domain';
import { isCellValue } from './value';


export class Cell {
  readonly address: string;
  readonly #workbook: Workbook;
  readonly #sheetId: string;
  readonly #row: number;
  readonly #column: number;
  constructor(workbook: Workbook, sheetId: string, row: number, column: number) { this.#workbook = workbook; this.#sheetId = sheetId; this.#row = row; this.#column = column; this.address = cellAddress(row, column); Object.freeze(this); }
  read(): Promise<CellSnapshot> { return domainFor(this.#workbook).read(this.#sheetId, this.#row, this.#column); }
  async setRichText(text: string, runs: readonly RichTextRun[]): Promise<void> { await this.#workbook.worksheets.byId(this.#sheetId).ranges.get(this.address).setRichText(text, runs); }
  async setStyle(style: Partial<CellStyle>, options: RangeStyleOptions = {}): Promise<void> { await this.#workbook.worksheets.byId(this.#sheetId).ranges.get(this.address).setStyle(style, options); }
  async setNumberFormat(numberFormat: string): Promise<void> { await this.#workbook.worksheets.byId(this.#sheetId).ranges.get(this.address).setNumberFormat(numberFormat); }
  async setBorders(placement: BorderPlacement, line?: BorderLine): Promise<void> { await this.#workbook.worksheets.byId(this.#sheetId).ranges.get(this.address).setBorders(placement, line); }
  async setValue(value: CellValue): Promise<void> {
    if (!isCellValue(value)) return domainFor(this.#workbook).invalid('cell.setValue', new Error('Value must be a finite Excel scalar or null'), { sheetId: this.#sheetId, address: this.address });
    await domainFor(this.#workbook).write(this.#sheetId, this.#row, this.#column, { kind: 'value', value });
  }
  async setFormula(formula: string): Promise<void> {
    if (typeof formula !== 'string' || !formula.startsWith('=') || formula.length < 2) return domainFor(this.#workbook).invalid('cell.setFormula', new Error('Formula must start with ='), { sheetId: this.#sheetId, address: this.address });
    await domainFor(this.#workbook).write(this.#sheetId, this.#row, this.#column, { kind: 'formula', formula });
  }
}
