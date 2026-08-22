import type { CellAddress, FormulaAst } from './ast';
import { cellAddressKey, compareCellAddresses, parseCellAddress } from './address';
import { collectFormulaDependencies } from './dependencies';
import { evaluateFormula } from './evaluator';
import { formatFormula } from './ast-format';
import { remapAst } from './ast-rewrite';
import { FormulaLexError, FormulaReferenceError, FormulaSyntaxError } from './errors';
import { parseFormula as parseFormulaSource } from './parser';
import { RangeIndex, type FormulaDependency, type RangeDependency } from './range-index';
import { createFormulaError, type FormulaError, type FormulaValue, type ScalarValue } from './values';

export type CellAddressInput = CellAddress | string;

export type CellInput = { readonly value: ScalarValue } | { readonly formula: string };

export interface FormulaResult {
  readonly value: FormulaValue;
  readonly formula?: string;
  readonly ast?: FormulaAst;
  readonly dependencies: readonly FormulaDependency[];
}

export interface RecalculationReport {
  readonly recalculated: readonly CellAddress[];
  readonly results: ReadonlyMap<string, FormulaResult>;
}

export interface FormulaEngineOptions {
  readonly defaultSheetId?: string;
}

interface StoredCell {
  readonly address: CellAddress;
  formula?: string;
  ast?: FormulaAst;
  parseError?: FormulaError;
  result: FormulaResult;
}

export class FormulaEngine {
  readonly defaultSheetId: string;
  readonly dependencies: RangeIndex;
  private definedNames: Record<string, string> = {};

  private readonly cells = new Map<string, StoredCell>();

  constructor(options: FormulaEngineOptions = {}) {
    this.defaultSheetId = options.defaultSheetId ?? 'Sheet1';
    if (!this.defaultSheetId) throw new Error('FormulaEngine requires a default worksheet id');
    this.dependencies = new RangeIndex();
  }

  setCell(addressInput: CellAddressInput, input: CellInput): FormulaResult {
    return 'formula' in input ? this.setFormula(addressInput, input.formula) : this.setValue(addressInput, input.value);
  }

  setValue(addressInput: CellAddressInput, value: ScalarValue): FormulaResult {
    const address = this.resolveAddress(addressInput);
    const key = cellAddressKey(address);
    this.dependencies.remove(address);
    const result: FormulaResult = { value, dependencies: [] };
    this.cells.set(key, { address, result });
    this.recalculate(address);
    return this.getCellResult(address) ?? result;
  }

  setFormula(addressInput: CellAddressInput, formula: string): FormulaResult {
    const address = this.resolveAddress(addressInput);
    const key = cellAddressKey(address);
    let ast: FormulaAst | undefined;
    let formulaDependencies: readonly FormulaDependency[] = [];
    let parseError: FormulaError | undefined;

    try {
      const parsed = this.parseFormula(formula);
      const extractedDependencies = collectFormulaDependencies(parsed, address);
      this.dependencies.set(address, extractedDependencies);
      ast = parsed;
      formulaDependencies = extractedDependencies;
    } catch (error) {
      parseError = formulaErrorFrom(error);
      this.dependencies.set(address, []);
    }

    const result: FormulaResult = parseError
      ? { value: parseError, formula, dependencies: [] }
      : { value: null, formula, ast, dependencies: formulaDependencies };
    this.cells.set(key, { address, formula, ast, parseError, result });
    this.recalculate(address);
    return this.getCellResult(address) ?? result;
  }

  clearCell(addressInput: CellAddressInput): RecalculationReport {
    const address = this.resolveAddress(addressInput);
    this.dependencies.remove(address);
    this.cells.delete(cellAddressKey(address));
    return this.recalculate(address);
  }

  getCellResult(addressInput: CellAddressInput): FormulaResult | undefined {
    const address = this.resolveAddress(addressInput);
    return this.cells.get(cellAddressKey(address))?.result;
  }

  getCellValue(addressInput: CellAddressInput): FormulaValue {
    return this.getCellResult(addressInput)?.value ?? null;
  }

  getDependencies(addressInput: CellAddressInput): readonly FormulaDependency[] {
    return this.dependencies.getDependencies(this.resolveAddress(addressInput));
  }

  getDependents(addressInput: CellAddressInput): readonly CellAddress[] {
    return this.dependencies.getDependents(this.resolveAddress(addressInput));
  }

  /** 批量设置定义名称(值为公式文本或引用文本,不带 =) */
  setDefinedNames(names: Record<string, string>): void {
    this.definedNames = { ...names };
  }

  getDefinedNames(): Record<string, string> {
    return { ...this.definedNames };
  }

  /** 清空全部公式与缓存(结构操作后整体重建前调用) */
  reset(): void {
    this.cells.clear();
    this.dependencies.clear?.();
  }

  /**
   * 结构变更(插入/删除行列)后重映射所有存活公式的引用,
   * 并用序列化器回写公式文本。
   */
  remapStructure(
    sheetId: string,
    shift: { axis: 'row' | 'column'; at: number; count: number; op: 'insert' | 'delete' },
  ): RecalculationReport {
    const updates: Array<{ address: CellAddress; formula: string }> = [];
    for (const [, cell] of this.cells) {
      if (cell.formula === undefined || !cell.ast) continue;
      const relevantSheet =
        cell.address.sheetId === sheetId
        || collectFormulaDependencies(cell.ast, cell.address).some((dependency) =>
          dependency.kind === 'cell'
            ? dependency.address.sheetId === sheetId
            : dependency.start.sheetId === sheetId,
        );
      if (!relevantSheet) continue;
      const remapped = remapAst(cell.ast, shift);
      updates.push({ address: { ...cell.address }, formula: formatFormula(remapped) });
    }
    for (const update of updates) {
      this.cells.delete(cellAddressKey(update.address));
    }
    let report: RecalculationReport = { recalculated: [], results: new Map() };
    for (const update of updates) {
      this.setFormula(update.address, update.formula);
    }
    report = this.recalculate();
    return report;
  }

  recalculate(addressInput?: CellAddressInput): RecalculationReport {
    const affected = addressInput === undefined
      ? this.allFormulaAddresses()
      : this.collectAffected(this.resolveAddress(addressInput));
    const evaluationCache = new Map<string, FormulaValue>();
    const visiting = new Set<string>();
    const recalculated: CellAddress[] = [];
    const results = new Map<string, FormulaResult>();

    const affectedEntries = [...affected.entries()]
      .map(([key, address]) => ({ key, address }))
      .sort((left, right) => compareCellAddresses(left.address, right.address));
    for (const { key, address } of affectedEntries) {
      const cell = this.cells.get(key);
      if (cell?.formula === undefined) continue;
      this.evaluateCell(address, evaluationCache, visiting);
      recalculated.push({ ...address });
      const result = this.cells.get(key)?.result;
      if (result) results.set(key, result);
    }

    return { recalculated, results };
  }

  private parseFormula(formula: string): FormulaAst {
    // Kept as a method so callers of the engine have one parse boundary and no executable formula path.
    return this.parseFormulaSource(formula);
  }

  private parseFormulaSource(formula: string): FormulaAst {
    // The parser builds data-only AST nodes; evaluation is performed by the dedicated evaluator.
    return parseFormulaSource(formula);
  }

  private resolveAddress(input: CellAddressInput): CellAddress {
    return typeof input === 'string' ? parseCellAddress(input, this.defaultSheetId) : { ...input };
  }

  private collectAffected(address: CellAddress): Map<string, CellAddress> {
    const affected = new Map<string, CellAddress>();
    const queue: CellAddress[] = [address];
    affected.set(cellAddressKey(address), { ...address });
    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) continue;
      for (const dependent of this.dependencies.getDependents(current)) {
        const key = cellAddressKey(dependent);
        if (affected.has(key)) continue;
        affected.set(key, dependent);
        queue.push(dependent);
      }
    }
    return affected;
  }

  private allFormulaAddresses(): Map<string, CellAddress> {
    const result = new Map<string, CellAddress>();
    for (const [key, cell] of this.cells) {
      if (cell.formula !== undefined) result.set(key, { ...cell.address });
    }
    return result;
  }

  private evaluateCell(
    address: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
  ): FormulaValue {
    const key = cellAddressKey(address);
    const cached = cache.get(key);
    if (cached !== undefined) return cached;
    if (visiting.has(key)) return createFormulaError('#CYCLE!', 'Circular reference detected');

    const cell = this.cells.get(key);
    if (!cell) {
      cache.set(key, null);
      return null;
    }
    if (cell.parseError) {
      cache.set(key, cell.parseError);
      cell.result = { value: cell.parseError, formula: cell.formula, dependencies: cell.result.dependencies };
      return cell.parseError;
    }
    if (!cell.ast) {
      cache.set(key, cell.result.value);
      return cell.result.value;
    }

    visiting.add(key);
    let value: FormulaValue;
    try {
      value = evaluateFormula(cell.ast, {
        currentCell: cell.address,
        readCell: (reference) => this.evaluateCell(reference, cache, visiting),
        readRange: (range) => this.readRange(range, cache, visiting),
        resolveName: (name) => this.resolveDefinedName(name),
      });
    } catch (error) {
      value = error instanceof FormulaReferenceError
        ? createFormulaError('#REF!', error.message)
        : createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Formula evaluation failed');
    } finally {
      visiting.delete(key);
    }

    cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
    cache.set(key, value);
    return value;
  }

  /** 定义名称求值:值文本按公式解析(相对当前单元格) */
  private resolveDefinedName(name: string): FormulaValue | undefined {
    const source = this.definedNames[name];
    if (source === undefined) return undefined;
    try {
      const parsed = this.parseFormula(source.startsWith('=') ? source : '=' + source);
      return evaluateFormula(parsed, {
        currentCell: { sheetId: this.defaultSheetId, row: 0, column: 0 },
        readCell: (reference) => this.getCellValue(reference),
        readRange: (range) => {
          const values: FormulaValue[] = [];
          for (let row = range.start.row; row <= range.end.row; row++) {
            for (let column = range.start.column; column <= range.end.column; column++) {
              values.push(this.getCellValue({ sheetId: range.start.sheetId, row, column }));
            }
          }
          return values;
        },
      });
    } catch {
      return createFormulaError('#NAME?', 'Cannot resolve name: ' + name);
    }
  }

  private readRange(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
  ): readonly FormulaValue[] {
    const values: FormulaValue[] = [];
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        values.push(this.evaluateCell({ sheetId: range.start.sheetId, row, column }, cache, visiting));
      }
    }
    return values;
  }
}

function formulaErrorFrom(error: unknown): FormulaError {
  if (error instanceof FormulaReferenceError) return createFormulaError('#REF!', error.message);
  if (error instanceof FormulaLexError || error instanceof FormulaSyntaxError) {
    return createFormulaError('#PARSE!', error.message, error.position);
  }
  return createFormulaError('#PARSE!', error instanceof Error ? error.message : 'Unable to parse formula');
}
