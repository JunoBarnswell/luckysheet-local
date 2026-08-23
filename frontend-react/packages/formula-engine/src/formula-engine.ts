import type { CellAddress, FormulaAst } from './ast';
import {
  DEFAULT_FORMULA_CAPABILITIES,
  type FormulaCapabilities,
} from './capabilities';
import { cellAddressKey, compareCellAddresses, parseCellAddress } from './address';
import { collectFormulaDependencies } from './dependencies';
import { evaluateFormula } from './evaluator';
import { formatFormula } from './ast-format';
import { remapAst } from './ast-rewrite';
import { FormulaLexError, FormulaReferenceError, FormulaSyntaxError } from './errors';
import { parseFormula as parseFormulaSource } from './parser';
import { RangeIndex, type FormulaDependency, type RangeDependency } from './range-index';
import { createFormulaError, isArrayValue, isFormulaError, type ArrayValue, type FormulaError, type FormulaValue, type ScalarValue } from './values';
import { normalizeDefinedNames, resolveDefinedNameSource } from './defined-names';
import { collectNameReferences, formulaUsesVolatile } from './formula-analysis';
import { normalizeSheetTables, resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import {
  assertCalculationTaskRequest,
  InlineCalculationTaskPort,
  type CalculationTaskPort,
  type CalculationTaskReport,
  type CalculationTaskRequest,
} from './calculation-task-port';
import {
  anchorDisplayValue,
  isSpillMatrix,
  resolveSpill,
  spillKey,
  spillValueAt,
  type ResolvedSpill,
} from './spill-resolver';

export interface SpillEnvironment {
  rowCount: number;
  columnCount: number;
  isOccupied: (row: number, column: number) => boolean;
}

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
  readonly recalculationMode?: RecalculationMode;
  /** Gated formula capabilities; omitted means all gated functions fail closed. */
  readonly capabilities?: FormulaCapabilities;
}

export type RecalculationMode = 'automatic' | 'manual';

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
  private spillEnvironments = new Map<string, SpillEnvironment>();
  private spills = new Map<string, ResolvedSpill>();
  private nameIndex = new Map<string, Set<string>>();
  private cellNameRefs = new Map<string, string[]>();
  private volatileCells = new Set<string>();
  private recalculationMode: RecalculationMode;
  private pendingRecalculationRoots = new Set<string>();
  private sheetTables = new Map<string, SheetTableRef>();
  private readonly capabilities: FormulaCapabilities;

  private readonly cells = new Map<string, StoredCell>();

  constructor(options: FormulaEngineOptions = {}) {
    this.defaultSheetId = options.defaultSheetId ?? 'Sheet1';
    this.recalculationMode = options.recalculationMode ?? 'automatic';
    this.capabilities = options.capabilities ?? DEFAULT_FORMULA_CAPABILITIES;
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
    this.spills.delete(spillKey(address));
    const result: FormulaResult = { value, dependencies: [] };
    this.cells.set(key, { address, result });
    if (this.recalculationMode === 'automatic') {
      this.recalculate(address);
    } else {
      this.pendingRecalculationRoots.add(key);
    }
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
      const extractedDependencies = collectFormulaDependencies(parsed, address, { sheetTables: this.sheetTables });
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
    this.updateFormulaMetadata(key, ast);
    if (this.recalculationMode === 'automatic') {
      this.recalculate(address);
    } else {
      this.evaluateChangedCell(address);
      this.pendingRecalculationRoots.add(key);
    }
    return this.getCellResult(address) ?? result;
  }

  clearCell(addressInput: CellAddressInput): RecalculationReport {
    const address = this.resolveAddress(addressInput);
    const key = cellAddressKey(address);
    this.dependencies.remove(address);
    this.spills.delete(spillKey(address));
    this.detachNameReferences(key);
    this.volatileCells.delete(key);
    this.cells.delete(key);
    return this.scheduleRecalculation(address) ?? { recalculated: [], results: new Map() };
  }

  getRecalculationMode(): RecalculationMode {
    return this.recalculationMode;
  }

  /**
   * Expose the calculation boundary used by a future Worker host.  The
   * default host is explicitly inline and synchronous; it shares the same
   * versioned, serializable request/result shape without pretending to offer
   * worker isolation.
   */
  createCalculationTaskPort(): CalculationTaskPort {
    return new InlineCalculationTaskPort((request) => this.executeCalculationTask(request));
  }

  executeCalculationTask(request: CalculationTaskRequest): CalculationTaskReport {
    assertCalculationTaskRequest(request);
    const reports = request.roots && request.roots.length > 0
      ? request.roots.map((root) => this.recalculate(root))
      : [this.recalculate()];
    const recalculated: CellAddress[] = [];
    const seen = new Set<string>();
    const results = new Map<string, FormulaResult>();
    for (const report of reports) {
      for (const address of report.recalculated) {
        const key = cellAddressKey(address);
        if (!seen.has(key)) {
          seen.add(key);
          recalculated.push({ ...address });
        }
      }
      for (const [key, result] of report.results) results.set(key, result);
    }
    const taskResults: CalculationTaskReport['results'][number][] = [];
    for (const [key, result] of results) {
      const address = this.cells.get(key)?.address;
      if (!address) continue;
      taskResults.push({
        address: { ...address },
        value: result.value,
        ...(result.formula === undefined ? {} : { formula: result.formula }),
        dependencies: result.dependencies,
      });
    }
    return { recalculated, results: taskResults };
  }

  setRecalculationMode(mode: RecalculationMode): void {
    this.recalculationMode = mode;
  }

  hasPendingRecalculation(): boolean {
    return this.pendingRecalculationRoots.size > 0;
  }

  setSheetTables(tables: readonly SheetTableRef[]): RecalculationReport {
    this.sheetTables = normalizeSheetTables(tables);
    const affected = this.allFormulaAddresses();
    if (this.recalculationMode === 'manual') {
      for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
      return { recalculated: [], results: new Map() };
    }
    return this.recalculateAffected(affected);
  }

  getSheetTables(): readonly SheetTableRef[] {
    return [...this.sheetTables.values()];
  }

  getCellResult(addressInput: CellAddressInput): FormulaResult | undefined {
    const address = this.resolveAddress(addressInput);
    return this.cells.get(cellAddressKey(address))?.result;
  }

  getCellValue(addressInput: CellAddressInput): FormulaValue {
    const address = this.resolveAddress(addressInput);
    const spillValue = this.getSpillValueAt(address.sheetId, address.row, address.column);
    if (spillValue !== undefined) return spillValue;
    return this.cells.get(cellAddressKey(address))?.result.value ?? null;
  }

  getDependencies(addressInput: CellAddressInput): readonly FormulaDependency[] {
    return this.dependencies.getDependencies(this.resolveAddress(addressInput));
  }

  getDependents(addressInput: CellAddressInput): readonly CellAddress[] {
    return this.dependencies.getDependents(this.resolveAddress(addressInput));
  }

  /** 批量设置定义名称(值为公式文本或引用文本,不带 =) */
  setDefinedNames(names: Record<string, string>): RecalculationReport {
    this.definedNames = normalizeDefinedNames(names);
    const affected = new Map<string, CellAddress>();
    for (const refs of this.nameIndex.values()) {
      for (const key of refs) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
      }
    }
    if (affected.size === 0) return { recalculated: [], results: new Map() };
    if (this.recalculationMode === 'manual') {
      for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
      return { recalculated: [], results: new Map() };
    }
    return this.recalculateAffected(affected);
  }

  setSpillEnvironment(sheetId: string, environment: SpillEnvironment | undefined): void {
    if (!environment) this.spillEnvironments.delete(sheetId);
    else this.spillEnvironments.set(sheetId, environment);
  }

  getSpillsForSheet(sheetId: string): ResolvedSpill[] {
    return [...this.spills.values()].filter((spill) => spill.sheetId === sheetId);
  }

  getSpillValueAt(sheetId: string, row: number, column: number): FormulaValue | undefined {
    for (const spill of this.spills.values()) {
      if (spill.sheetId !== sheetId) continue;
      const value = spillValueAt(spill, row, column);
      if (value !== undefined) return value;
    }
    return undefined;
  }

  getDefinedNames(): Record<string, string> {
    return { ...this.definedNames };
  }

  /** 清空全部公式与缓存(结构操作后整体重建前调用) */
  reset(): void {
    this.cells.clear();
    this.spills.clear();
    this.nameIndex.clear();
    this.cellNameRefs.clear();
    this.volatileCells.clear();
    this.pendingRecalculationRoots.clear();
    this.sheetTables.clear();
    this.dependencies.clear?.();
  }

  recalculateCell(addressInput: CellAddressInput): RecalculationReport {
    const address = this.resolveAddress(addressInput);
    const affected = new Map<string, CellAddress>();
    const key = cellAddressKey(address);
    const cell = this.cells.get(key);
    if (cell?.formula !== undefined) affected.set(key, { ...address });
    return this.recalculateAffected(affected);
  }

  recalculate(addressInput?: CellAddressInput): RecalculationReport {
    if (addressInput !== undefined) {
      const affected = this.collectAffected(this.resolveAddress(addressInput));
      for (const key of this.volatileCells) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
      }
      return this.recalculateAffected(affected);
    }

    const affected = this.recalculationMode === 'manual' && this.pendingRecalculationRoots.size > 0
      ? this.collectAffectedFromRoots(this.pendingRecalculationRoots)
      : this.allFormulaAddresses();
    for (const key of this.volatileCells) {
      const cell = this.cells.get(key);
      if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
    }
    const report = this.recalculateAffected(affected);
    this.pendingRecalculationRoots.clear();
    return report;
  }

  private scheduleRecalculation(address: CellAddress): RecalculationReport | undefined {
    if (this.recalculationMode === 'automatic') return this.recalculate(address);
    this.pendingRecalculationRoots.add(cellAddressKey(address));
    return undefined;
  }

  private evaluateChangedCell(address: CellAddress): void {
    const key = cellAddressKey(address);
    const cell = this.cells.get(key);
    if (!cell?.formula || !cell.ast) return;
    const cache = new Map<string, FormulaValue>();
    const visiting = new Set<string>();
    this.evaluateCell(address, cache, visiting);
  }

  private collectAffectedFromRoots(roots: ReadonlySet<string>): Map<string, CellAddress> {
    const affected = new Map<string, CellAddress>();
    for (const key of roots) {
      const cell = this.cells.get(key);
      if (!cell) continue;
      for (const [dependentKey, dependentAddress] of this.collectAffected(cell.address)) {
        affected.set(dependentKey, dependentAddress);
      }
    }
    return affected;
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
        || collectFormulaDependencies(cell.ast, cell.address, { sheetTables: this.sheetTables }).some((dependency) =>
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

  private recalculateAffected(affected: Map<string, CellAddress>): RecalculationReport {
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
        capabilities: this.capabilities,
        readCell: (reference) => this.evaluateCell(reference, cache, visiting),
        readRange: (range) => this.readRange(range, cache, visiting),
        resolveName: (name) => this.resolveDefinedName(name, cell.address, cache, visiting),
        resolveTableReference: (tableName, request) => {
          const resolved = resolveSheetTableReference(tableName, request, cell.address, this.sheetTables);
          if (isFormulaError(resolved)) return resolved;
          if ('start' in resolved && 'end' in resolved) return { kind: 'range', range: resolved };
          return this.evaluateCell(resolved as CellAddress, cache, visiting);
        },
      });
    } catch (error) {
      value = error instanceof FormulaReferenceError
        ? createFormulaError('#REF!', error.message)
        : createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Formula evaluation failed');
    } finally {
      visiting.delete(key);
    }

    cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
    this.refreshSpill(address, value, cache);
    const displayValue = this.cells.get(key)?.result.value ?? value;
    cache.set(key, displayValue);
    return displayValue;
  }

  private refreshSpill(address: CellAddress, value: FormulaValue, cache?: Map<string, FormulaValue>): void {
    const key = spillKey(address);
    if (!isSpillMatrix(value)) {
      this.spills.delete(key);
      return;
    }
    const environment = this.spillEnvironments.get(address.sheetId);
    if (!environment) {
      this.spills.delete(key);
      return;
    }
    const spill = resolveSpill({
      sheetId: address.sheetId,
      anchor: { row: address.row, column: address.column },
      values: value,
      rowCount: environment.rowCount,
      columnCount: environment.columnCount,
      isOccupied: environment.isOccupied,
    });
    this.spills.set(key, spill);
    if (spill.state === 'blocked') {
      const display = anchorDisplayValue(spill, value);
      const cell = this.cells.get(cellAddressKey(address));
      if (cell) cell.result = { ...cell.result, value: display };
      cache?.set(cellAddressKey(address), display);
    }
  }

  private updateFormulaMetadata(key: string, ast?: FormulaAst): void {
    this.detachNameReferences(key);
    if (!ast) {
      this.volatileCells.delete(key);
      return;
    }
    if (formulaUsesVolatile(ast)) this.volatileCells.add(key);
    else this.volatileCells.delete(key);
    const names = collectNameReferences(ast);
    this.cellNameRefs.set(key, names);
    for (const name of names) {
      const bucket = this.nameIndex.get(name) ?? new Set<string>();
      bucket.add(key);
      this.nameIndex.set(name, bucket);
    }
  }

  private detachNameReferences(key: string): void {
    for (const name of this.cellNameRefs.get(key) ?? []) {
      const bucket = this.nameIndex.get(name);
      bucket?.delete(key);
      if (bucket && bucket.size === 0) this.nameIndex.delete(name);
    }
    this.cellNameRefs.delete(key);
  }

  private resolveDefinedName(
    name: string,
    currentCell: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
  ): FormulaValue | undefined {
    const source = this.definedNames[name.toUpperCase()] ?? this.definedNames[name];
    if (source === undefined) return undefined;
    return resolveDefinedNameSource(source, {
      currentCell,
      readCell: (reference) => this.evaluateCell(reference, cache, visiting),
      readRangeMatrix: (range) => this.readRangeMatrix(range, cache, visiting),
    });
  }

  private readRangeMatrix(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
  ): ArrayValue {
    const matrix: ArrayValue = [];
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      const line: FormulaValue[] = [];
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        line.push(this.evaluateCell({ sheetId: range.start.sheetId, row, column }, cache, visiting));
      }
      matrix.push(line);
    }
    return matrix;
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
