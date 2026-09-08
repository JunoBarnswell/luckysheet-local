import { KernelInvocationError, kernelInvoke } from '../../kernel-client/src/index';
import { cellAddressKey, parseCellAddress } from './address';
import type { CellAddress, FormulaAst } from './ast';
import { formatFormula } from './ast-format';
import { parseFormula } from './parser';
import type { FormulaDependency } from './range-index';
import type { FormulaValue, ScalarValue } from './values';
import type { FormulaDefinedName } from './defined-names';
import type { SheetTableRef } from './sheet-table-resolver';
import type { ResolvedSpill } from './spill-resolver';
import type { WorkbookCalculationMode, WorkbookCalculationSettings } from './calculation-settings';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';
import type { ExcelNumericContext } from './numeric';
import type { WorkbookCollationContext } from './collation';

export type CellAddressInput = CellAddress | string;
export type RecalculationMode = WorkbookCalculationMode;
export interface FormulaCellOverride { readonly address: CellAddress; readonly value: ScalarValue; }
export interface FormulaResult {
  readonly value: FormulaValue;
  readonly formula?: string;
  readonly ast?: FormulaAst;
  readonly dependencies: readonly FormulaDependency[];
}
export interface FormulaCellEntry extends FormulaResult { readonly address: CellAddress; readonly formula: string; }
export interface FormulaEvaluationTraceStep { readonly expression: string; readonly value: FormulaValue; }
export interface FormulaEvaluationTrace { readonly value: FormulaValue; readonly steps: readonly FormulaEvaluationTraceStep[]; }
export interface RecalculationReport { readonly revision: number; readonly generation: number; readonly recalculatedCount: number; readonly pendingRecalculation: boolean; }
export interface FormulaEngineOptions {
  readonly unitId: string;
  /** Always reads the canonical model revision; the binding owns no workbook replica. */
  readonly revision: () => number;
  readonly defaultSheetId: string;
}
interface FormulaInspection {
  readonly revision: number;
  readonly generation: number;
  readonly entries: readonly FormulaCellEntry[];
  readonly dependents: readonly CellAddress[];
  readonly spills: readonly ResolvedSpill[];
  readonly pendingRecalculation: boolean;
  readonly nextCursor?: string | null;
}
interface FormulaMetadata {
  readonly dateSystem: ExcelDateSystem;
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  readonly numericContext: ExcelNumericContext;
  readonly collationContext: WorkbookCollationContext;
  readonly calculationSettings: WorkbookCalculationSettings;
  readonly definedNameModels: readonly FormulaDefinedName[];
}
interface Manifest {
  readonly revision: number;
  readonly metadata: FormulaMetadata;
  readonly sheets: readonly { readonly sheetId: string; readonly metadata: { readonly sheetTables?: readonly SheetTableRef[] } }[];
}

/** Revision-pinned Rust formula binding. All authored writes belong to model commands. */
export class FormulaEngine {
  readonly defaultSheetId: string;
  readonly unitId: string;
  private readonly readRevision: () => number;

  constructor(options: FormulaEngineOptions) {
    if (!options.unitId || !options.defaultSheetId || typeof options.revision !== 'function') {
      throw new KernelInvocationError({ code: 'FORMULA_BINDING_INVALID', message: 'Formula binding requires workbook identity, revision and worksheet identity.', recovery: 'bind-canonical-workbook' });
    }
    this.unitId = options.unitId;
    this.defaultSheetId = options.defaultSheetId;
    this.readRevision = options.revision;
  }

  private invoke<T extends { readonly revision: number }>(operation: string, params: Record<string, unknown> = {}): T {
    const revision = this.readRevision();
    if (!Number.isSafeInteger(revision) || revision < 0) throw new KernelInvocationError({ code: 'REVISION_INVALID', message: 'Formula binding revision is invalid.', object: this.unitId, recovery: 'reload-workbook' });
    const result = kernelInvoke<T>(operation, { ...params, unitId: this.unitId, revision });
    if (!result || typeof result !== 'object' || result.revision !== revision) throw new KernelInvocationError({ code: 'REVISION_MISMATCH', message: 'Formula result revision does not match its canonical input.', object: this.unitId, recovery: 'reload-workbook' });
    return result;
  }
  private address(input: CellAddressInput): CellAddress { return typeof input === 'string' ? parseCellAddress(input, this.defaultSheetId) : input; }
  private inspect(address?: CellAddress, projection: 'entries' | 'spills' | 'status' = 'entries', paging: { sheetId?: string; cursor?: string; limit?: number } = {}): FormulaInspection {
    const result = this.invoke<FormulaInspection>('formula.inspect', { projection, ...paging, ...(address ? { address } : {}) });
    if (!Array.isArray(result.entries) || !Array.isArray(result.dependents) || !Array.isArray(result.spills) || typeof result.pendingRecalculation !== 'boolean' || !Number.isSafeInteger(result.generation)) {
      throw new KernelInvocationError({ code: 'KERNEL_PROTOCOL_ERROR', message: 'Formula inspection response is incomplete.', object: this.unitId, recovery: 'reload-kernel' });
    }
    return result;
  }
  private manifest(): Manifest { return this.invoke('manifest'); }
  private metadata<K extends keyof FormulaMetadata>(key: K): FormulaMetadata[K] {
    const metadata = this.manifest().metadata;
    if (!metadata || !Object.hasOwn(metadata, key)) throw new KernelInvocationError({ code: 'FORMULA_METADATA_MISSING', message: 'Canonical formula metadata is missing: ' + key, object: this.unitId, recovery: 'migrate-workbook-metadata' });
    return metadata[key];
  }

  evaluateFormula(formula: string, address: CellAddressInput, overrides: readonly FormulaCellOverride[] = []): FormulaValue {
    return this.invoke<{ revision: number; value: FormulaValue }>('formula.evaluate', { address: this.address(address), formula, ...(overrides.length ? { overrides } : {}) }).value;
  }
  evaluateAst(ast: FormulaAst, currentCell: CellAddress, overrides: readonly FormulaCellOverride[] = []): FormulaValue {
    return this.evaluateFormula(formatFormula(ast), currentCell, overrides);
  }
  getCellResult(addressInput: CellAddressInput): FormulaResult | undefined {
    const address = this.address(addressInput);
    const { cell } = this.invoke<{ revision: number; cell: { value: FormulaValue; formula?: string } | null }>('cell.get', { address });
    if (cell === null) return undefined;
    if (!cell || !Object.hasOwn(cell, 'value')) throw new KernelInvocationError({ code: 'KERNEL_PROTOCOL_ERROR', message: 'Canonical cell response is incomplete.', object: cellAddressKey(address), recovery: 'reload-kernel' });
    if (cell.formula !== undefined) {
      const entry = this.inspect(address).entries.find(candidate => cellAddressKey(candidate.address) === cellAddressKey(address));
      if (!entry) throw new KernelInvocationError({ code: 'FORMULA_RESULT_MISSING', message: 'An authored formula has no canonical calculation result.', object: cellAddressKey(address), recovery: 'recalculate-workbook' });
      return entry;
    }
    return { value: cell.value, dependencies: [] };
  }
  getCellValue(addressInput: CellAddressInput): FormulaValue {
    return this.getCellResult(addressInput)?.value ?? null;
  }
  getFormulaEntriesPage(options: { cursor?: string; limit?: number; sheetId?: string } = {}): { revision: number; entries: readonly FormulaCellEntry[]; nextCursor?: string | null } {
    const result = this.inspect(undefined, 'entries', options);
    return { revision: result.revision, entries: result.entries, nextCursor: result.nextCursor };
  }
  getFormulaAst(address: CellAddressInput): FormulaAst | undefined { const formula = this.getCellResult(address)?.formula; return formula === undefined ? undefined : parseFormula(formula); }
  getDependencies(address: CellAddressInput): readonly FormulaDependency[] { return this.getCellResult(address)?.dependencies ?? []; }
  getDependents(address: CellAddressInput): readonly CellAddress[] { return this.inspect(this.address(address)).dependents; }
  evaluateFormulaWithTrace(address: CellAddressInput): FormulaEvaluationTrace | undefined {
    const resolved = this.address(address);
    if (this.getCellResult(resolved)?.formula === undefined) return undefined;
    return this.invoke<FormulaEvaluationTrace & { revision: number }>('formula.trace', { address: resolved });
  }
  getSpillsForSheetPage(sheetId: string, options: { cursor?: string; limit?: number } = {}): { revision: number; spills: readonly ResolvedSpill[]; nextCursor?: string | null } {
    const result = this.inspect(undefined, 'spills', { ...options, sheetId });
    return { revision: result.revision, spills: result.spills, nextCursor: result.nextCursor };
  }
  getSpillValueAt(sheetId: string, row: number, column: number): FormulaValue | undefined {
    const result = this.invoke<{ revision: number; value: FormulaValue | null; isSpill: boolean }>('formula.spillValue', { address: { sheetId, row, column } });
    return result.isSpill ? result.value : undefined;
  }
  recalculate(): RecalculationReport { return this.invoke<RecalculationReport>('formula.recalculate'); }
  async recalculateAsync(): Promise<RecalculationReport> { return this.recalculate(); }
  hasPendingRecalculation(): boolean { return this.inspect(undefined, 'status').pendingRecalculation; }
  getCalculationGeneration(): number { return this.inspect(undefined, 'status').generation; }
  getCalculationSettings(): WorkbookCalculationSettings { return this.metadata('calculationSettings'); }
  getRecalculationMode(): RecalculationMode { return this.getCalculationSettings().mode; }
  getCanonicalReferenceDate(): CanonicalExcelDateParts | undefined { return this.manifest().metadata.canonicalReferenceDate; }
  getDateSystem(): ExcelDateSystem { return this.metadata('dateSystem'); }
  getNumericContext(): ExcelNumericContext { return this.metadata('numericContext'); }
  getCollationContext(): WorkbookCollationContext { return this.metadata('collationContext'); }
  getDefinedNameModels(): FormulaDefinedName[] { return [...this.metadata('definedNameModels')]; }
  getDefinedNames(): Record<string, string> { return Object.fromEntries(this.getDefinedNameModels().filter(name => name.scope === 'workbook').map(name => [name.name, name.formula])); }
  getSheetTables(): readonly SheetTableRef[] { return this.manifest().sheets.flatMap(sheet => sheet.metadata.sheetTables ?? []); }
}
