import { cellAddress, type CellValue, type FormulaValue, type CellStyle, type BorderPlacement, type BorderLine, type RangeRef, type RichTextRun } from '@react-sheets/core-model';
import type { ClearFamily, FillDirection, FillMode, FillSeriesOptions } from '@react-sheets/sheet-features';
import type { Cell } from './cell';
import type { Workbook } from './workbook';
import type { Worksheet } from './worksheet';
import type { CellInput, CellSnapshot, RangeStyleOptions } from './contract';
import { domainFor } from './object-domain';
import { immutableSnapshot, isCellValue, MAX_OBJECT_RANGE_CELLS } from './value';
import { SdkError } from '../error';

export class Range {
  readonly address: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly #sheet: Worksheet;
  readonly #workbook: Workbook;
  readonly #range: RangeRef;
  constructor(workbook: Workbook, sheet: Worksheet, start: { row: number; column: number }, end: { row: number; column: number }) {
    this.#workbook = workbook; this.#sheet = sheet;
    this.#range = Object.freeze({ sheetId: sheet.id, startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column });
    this.rowCount = end.row - start.row + 1; this.columnCount = end.column - start.column + 1;
    this.address = `${cellAddress(start.row, start.column)}:${cellAddress(end.row, end.column)}`;
    Object.freeze(this);
  }
  #assertBounded(operation: string): void {
    this.#sheet.name;
    if (this.rowCount * this.columnCount > MAX_OBJECT_RANGE_CELLS) throw new SdkError('UNSUPPORTED_FEATURE', operation, `Range exceeds the ${MAX_OBJECT_RANGE_CELLS} cell operation budget.`, 'Submit smaller explicit ranges.', { object: { workbookId: this.#workbook.id, sheetId: this.#sheet.id, address: this.address } });
  }
  *cells(): IterableIterator<Cell> {
    this.#sheet.name;
    for (let row = this.#range.startRow; row <= this.#range.endRow; row++) for (let column = this.#range.startColumn; column <= this.#range.endColumn; column++) yield this.#sheet.cells.get(cellAddress(row, column));
  }
  async read(): Promise<readonly (readonly CellSnapshot[])[]> { this.#assertBounded('range.read'); return domainFor(this.#workbook).readRange(this.#range); }
  async readValues(): Promise<readonly (readonly FormulaValue[])[]> {
    return immutableSnapshot((await this.read()).map(row => row.map(cell => cell.calculatedValue)));
  }
  async setValues(values: readonly (readonly CellValue[])[]): Promise<void> {
    this.#validateMatrix(values, 'range.setValues', isCellValue);
    await domainFor(this.#workbook).writeRange(this.#range, values.map(row => row.map(value => ({ kind: 'value', value }))));
  }
  async setFormulas(formulas: readonly (readonly string[])[]): Promise<void> {
    this.#validateMatrix(formulas, 'range.setFormulas', value => typeof value === 'string' && value.startsWith('=') && value.length > 1);
    await domainFor(this.#workbook).writeRange(this.#range, formulas.map(row => row.map(formula => ({ kind: 'formula', formula }))));
  }
  async setInputs(inputs: readonly (readonly CellInput[])[]): Promise<void> {
    this.#validateMatrix(inputs, 'range.setInputs', input => {
      if (!input || typeof input !== 'object' || !('kind' in input)) return false;
      if (input.kind === 'value') return 'value' in input && isCellValue(input.value);
      return input.kind === 'formula' && 'formula' in input && typeof input.formula === 'string' && input.formula.startsWith('=') && input.formula.length > 1;
    });
    await domainFor(this.#workbook).writeRange(this.#range, inputs);
  }
  #validateMatrix(matrix: unknown, operation: string, validate: (value: unknown) => boolean): void {
    this.#assertBounded(operation);
    if (!Array.isArray(matrix) || matrix.length !== this.rowCount || !matrix.every(row => Array.isArray(row) && row.length === this.columnCount && row.every(validate))) {
      domainFor(this.#workbook).invalid(operation, new Error('Input must exactly match the range dimensions and canonical value contract.'), { sheetId: this.#sheet.id, address: this.address });
    }
  }
  async setRichText(text: string, runs: readonly RichTextRun[]): Promise<void> {
    this.#assertBounded('range.setRichText');
    if (typeof text !== 'string' || text.length > 32_767 || !Array.isArray(runs)
      || runs.some(run => !run || typeof run.text !== 'string') || runs.map(run => run.text).join('') !== text) {
      return domainFor(this.#workbook).invalid('range.setRichText', new Error('Canonical rich-text runs must reproduce the complete plain text.'), { sheetId: this.#sheet.id, address: this.address });
    }
    await domainFor(this.#workbook).writeRichText(this.#range, text, runs);
  }
  async clear(family: ClearFamily = 'contents'): Promise<void> { await this.#command('range.clear', 'sheet.range.clear', { family }); }
  async setStyle(style: Partial<CellStyle>, options: RangeStyleOptions = {}): Promise<void> { await this.#command('range.setStyle', 'sheet.style.set', { style, ...options }); }
  async setNumberFormat(numberFormat: string): Promise<void> { await this.setStyle({}, { numberFormat }); }
  async setBorders(placement: BorderPlacement, line?: BorderLine): Promise<void> { await this.#command('range.setBorders', 'sheet.borders.set', { placement, ...(line === undefined ? {} : { line }) }); }
  async merge(options: { center?: boolean; across?: boolean; confirmDataLoss?: boolean } = {}): Promise<void> {
    await this.#command('range.merge', options.across ? 'sheet.merge.across' : options.center ? 'sheet.merge.center' : 'sheet.merge.cells', { confirmDataLoss: options.confirmDataLoss === true });
  }
  async unmerge(): Promise<void> { await this.#command('range.unmerge', 'sheet.merge.unmerge', {}); }
  async fillFrom(source: Range, direction: FillDirection, mode: FillMode = 'copy', series?: FillSeriesOptions): Promise<void> {
    this.#assertBounded('range.fillFrom'); source.#assertBounded('range.fillFrom');
    if (source.#workbook !== this.#workbook || source.#sheet.id !== this.#sheet.id) domainFor(this.#workbook).invalid('range.fillFrom', new Error('Fill requires source and target on the same worksheet.'));
    await domainFor(this.#workbook).command('range.fillFrom', { commandId: 'sheet.range.fill', params: { sheetId: this.#sheet.id, sourceRange: source.#range, targetRange: this.#range, direction, mode, ...(series ? { series } : {}) } });
  }
  /** Cut and replace an equal-sized range on this worksheet in one canonical transaction. */
  async moveTo(destination: Range): Promise<void> {
    this.#assertBounded('range.moveTo'); destination.#assertBounded('range.moveTo');
    if (this.#workbook !== destination.#workbook || this.#sheet.id !== destination.#sheet.id) {
      throw new SdkError('UNSUPPORTED_FEATURE', 'range.moveTo', 'Range moves require the same workbook and worksheet.', 'Use a same-worksheet destination; cross-worksheet and cross-workbook moves require a shared transaction.', { object: { workbookId: this.#workbook.id, sheetId: this.#sheet.id, address: this.address } });
    }
    if (this.rowCount !== destination.rowCount || this.columnCount !== destination.columnCount) {
      domainFor(this.#workbook).invalid('range.moveTo', new Error('Move source and destination must have equal dimensions.'), { sheetId: this.#sheet.id, address: this.address });
    }
    await domainFor(this.#workbook).command('range.moveTo', { commandId: 'sheet.range.move', params: {
      sheetId: this.#sheet.id, sourceRange: this.#range,
      targetOrigin: { row: destination.#range.startRow, column: destination.#range.startColumn },
    } });
  }
  async copyValuesTo(target: Range): Promise<void> {
    this.#assertBounded('range.copyValuesTo'); target.#assertBounded('range.copyValuesTo');
    if (domainFor(this.#workbook).scope !== domainFor(target.#workbook).scope || this.rowCount !== target.rowCount || this.columnCount !== target.columnCount) domainFor(this.#workbook).invalid('range.copyValuesTo', new Error('Copy requires equal dimensions and workbooks in the same SDK context.'));
    const values = (await this.read()).map(row => row.map(cell => cell.calculatedValue));
    if (!values.every(row => row.every(isCellValue))) throw new SdkError('UNSUPPORTED_FEATURE', 'range.copyValuesTo', 'Calculated error values have no authored CellValue representation.', 'Correct the source error before copying values.', { object: { workbookId: this.#workbook.id, sheetId: this.#sheet.id, address: this.address } });
    this.#sheet.name;
    await target.setValues(values as CellValue[][]);
  }
  async #command(operation: string, commandId: string, params: object): Promise<void> {
    this.#assertBounded(operation);
    await domainFor(this.#workbook).command(operation, { commandId, params: { ...params, sheetId: this.#sheet.id, range: this.#range } });
  }
}
