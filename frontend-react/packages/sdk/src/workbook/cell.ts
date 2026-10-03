import { cellAddress, type CellValue } from '@react-sheets/core-model';
import type { Workbook } from './workbook';
import type { CellSnapshot } from './contract';
import { domainFor } from './object-domain';

function valueIsValid(value: unknown): value is CellValue {
  return value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value);
}


export class Cell {
  readonly address: string;
  readonly #workbook: Workbook;
  readonly #sheetId: string;
  readonly #row: number;
  readonly #column: number;
  constructor(workbook: Workbook, sheetId: string, row: number, column: number) { this.#workbook = workbook; this.#sheetId = sheetId; this.#row = row; this.#column = column; this.address = cellAddress(row, column); Object.freeze(this); }
  read(): Promise<CellSnapshot> { return domainFor(this.#workbook).read(this.#sheetId, this.#row, this.#column); }
  async setValue(value: CellValue): Promise<void> {
    if (!valueIsValid(value)) return domainFor(this.#workbook).invalid('cell.setValue', new Error('Value must be a finite Excel scalar or null'), { sheetId: this.#sheetId, address: this.address });
    await domainFor(this.#workbook).write(this.#sheetId, this.#row, this.#column, { kind: 'value', value });
  }
  async setFormula(formula: string): Promise<void> {
    if (typeof formula !== 'string' || !formula.startsWith('=') || formula.length < 2) return domainFor(this.#workbook).invalid('cell.setFormula', new Error('Formula must start with ='), { sheetId: this.#sheetId, address: this.address });
    await domainFor(this.#workbook).write(this.#sheetId, this.#row, this.#column, { kind: 'formula', formula });
  }
}
