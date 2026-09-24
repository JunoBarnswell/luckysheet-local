import type { CellAddress, FormulaAst, FormulaReferenceNode } from './ast';
import { cellAddressKey, compareCellAddresses, parseCellAddress } from './address';
import { collectFormulaDependencies } from './dependencies';
import {
  evaluateFormula,
  evaluateFormulaWithTrace,
  type FormulaCellOverride,
  type FormulaEvaluationContext,
  type FormulaEvaluationReference,
  type FormulaEvaluationTrace,
} from './evaluator';
import { formatFormula } from './ast-format';
import { offsetAst } from './ast-rewrite';
import { FormulaLexError, FormulaReferenceError, FormulaSyntaxError } from './errors';
import { parseFormula as parseFormulaSource } from './parser';
import { RangeIndex, type FormulaDependency, type FormulaSheetIdentity, type RangeDependency } from './range-index';
import { resolveFormulaSheetId } from './sheet-reference';
import { createFormulaError, isArrayValue, isFormulaError, type ArrayValue, type FormulaError, type FormulaValue, type ScalarValue } from './values';
import { normalizeDefinedNameModels, normalizeDefinedNames, parseDefinedNameFormula, resolveDefinedNameSource, type FormulaDefinedName } from './defined-names';
import { collectNameReferences, collectTableReferences, formulaUsesRowVisibility, formulaUsesVolatile } from './formula-analysis';
import { normalizeSheetTables, resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';
import { DEFAULT_EXCEL_NUMERIC_CONTEXT, normalizeExcelNumericContext, type ExcelNumericContext } from './numeric';
import { createCalculationEntropyContext, formulaRandom, type CalculationEntropyContext } from './random';
import { DEFAULT_WORKBOOK_COLLATION, normalizeWorkbookCollation, type WorkbookCollationContext } from './collation';
import { findFormulaComponents } from './circular';
import { createSnapshotVisibilityResolver, type ReferenceFormulaKind, type RowVisibilityResolver } from './reference-cursor';
import { DEFAULT_WORKBOOK_CALCULATION_SETTINGS, normalizeWorkbookCalculationSettings, type WorkbookCalculationMode, type WorkbookCalculationSettings } from './calculation-settings';
import {
  assertCalculationTaskRequest,
  InlineCalculationTaskPort,
  type CalculationInputUpdate,
  type CalculationTaskPort,
  type CalculationTaskReport,
  type CalculationTaskRequest,
  type CalculationTaskResult,
} from './calculation-task-port';
import {
  BrowserCalculationTaskPort,
  createBrowserCalculationWorker,
  type CalculationTaskState,
  type CalculationBrowserWorkerFactory,
} from './calculation-browser-task-port';
import {
  assertFormulaCalculationSnapshot,
  type FormulaCalculationSnapshot,
} from './calculation-state';
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
  /** Update the worker's persisted occupancy projection after an input delta. */
  applyOccupiedUpdate?: (address: CellAddress, occupied: boolean) => void;
  /** Grow the host runtime extent before a legal spill is resolved. */
  ensureExtent?: (rowCount: number, columnCount: number) => void;
  /** Optional exact occupancy materializer for Worker-bound calculation. */
  getOccupiedAddresses?: () => readonly { readonly row: number; readonly column: number }[];
}

export type CellAddressInput = CellAddress | string;

export type CellInput = { readonly value: ScalarValue } | { readonly formula: string };

export interface FormulaResult {
  readonly value: FormulaValue;
  readonly formula?: string;
  readonly ast?: FormulaAst;
  readonly dependencies: readonly FormulaDependency[];
}

/** Structured formula input/result used by audit and host-side projections. */
export interface FormulaCellEntry {
  readonly address: CellAddress;
  readonly formula: string;
  readonly value: FormulaValue;
  readonly ast?: FormulaAst;
  readonly dependencies: readonly FormulaDependency[];
}

export interface RecalculationReport {
  readonly recalculated: readonly CellAddress[];
  /** Formula outputs or spill projections whose canonical result changed. */
  readonly changedAddresses?: readonly CellAddress[];
  readonly results: ReadonlyMap<string, FormulaResult>;
}

function sameCalculationValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (typeof left !== 'object' || left === null || typeof right !== 'object' || right === null) return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left) && Array.isArray(right)
      && left.length === right.length
      && left.every((value, index) => sameCalculationValue(value, right[index]));
  }
  const leftKeys = Object.keys(left);
  const rightRecord = right as Record<string, unknown>;
  const rightKeys = Object.keys(rightRecord);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) => Object.prototype.hasOwnProperty.call(rightRecord, key)
    && sameCalculationValue((left as Record<string, unknown>)[key], rightRecord[key]));
}

function changedMapKeys<T>(previous: ReadonlyMap<string, T>, next: ReadonlyMap<string, T>): Set<string> {
  const changed = new Set<string>();
  for (const key of new Set([...previous.keys(), ...next.keys()])) {
    if (JSON.stringify(previous.get(key)) !== JSON.stringify(next.get(key))) changed.add(key);
  }
  return changed;
}

function changedDefinedNameTokens(
  previous: readonly FormulaDefinedName[],
  next: readonly FormulaDefinedName[],
  identity: (entry: FormulaDefinedName) => string,
): Set<string> {
  const previousByIdentity = new Map(previous.map((entry) => [identity(entry), entry] as const));
  const nextByIdentity = new Map(next.map((entry) => [identity(entry), entry] as const));
  const changed = new Set<string>();
  for (const key of new Set([...previousByIdentity.keys(), ...nextByIdentity.keys()])) {
    const before = previousByIdentity.get(key);
    const after = nextByIdentity.get(key);
    if (JSON.stringify(before) === JSON.stringify(after)) continue;
    const token = (after ?? before)?.name.trim().toUpperCase();
    if (token) changed.add(token);
  }
  return changed;
}

export interface FormulaEngineOptions {
  readonly defaultSheetId?: string;
  readonly sheetOrder?: readonly FormulaSheetIdentity[];
  readonly recalculationMode?: RecalculationMode;
  readonly dateSystem?: ExcelDateSystem;
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  readonly numericContext?: Partial<ExcelNumericContext>;
  /** Stable host/workbook seed; authoritative collaboration code may replace it. */
  readonly calculationEntropySeed?: string;
  readonly collationContext?: Partial<WorkbookCollationContext>;
  readonly calculationSettings?: Partial<WorkbookCalculationSettings>;
  readonly rowVisibilityResolver?: RowVisibilityResolver;
}

export interface CalculationTaskPortOptions {
  /** Allows deterministic tests or a host-owned browser Worker factory. */
  readonly workerFactory?: CalculationBrowserWorkerFactory;
  /** Set false only for explicitly synchronous hosts. */
  readonly useWorker?: boolean;
}

export type RecalculationMode = WorkbookCalculationMode;

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
  private readonly sheetOrder: readonly FormulaSheetIdentity[];
  /** Canonical scoped names. Workbook-only lookup is derived on demand. */
  private definedNameModels: FormulaDefinedName[] = [];
  private spillEnvironments = new Map<string, SpillEnvironment>();
  private spills = new Map<string, ResolvedSpill>();
  private nameIndex = new Map<string, Set<string>>();
  private cellNameRefs = new Map<string, string[]>();
  private tableReferenceIndex = new Map<string, Set<string>>();
  private cellTableRefs = new Map<string, string[]>();
  private volatileCells = new Set<string>();
  private visibilityDependentCells = new Set<string>();
  private recalculationMode: RecalculationMode;
  private pendingRecalculationRoots = new Set<string>();
  private sheetTables = new Map<string, SheetTableRef>();
  private calculationGeneration = 0;
  private calculationContextGeneration = 0;
  private inputUpdateSequence = 0;
  private pendingInputUpdates = new Map<string, { sequence: number; update: CalculationInputUpdate }>();
  private formulaTopologyGeneration = 0;
  private cachedCircularComponents: { generation: number; components: ReturnType<typeof findFormulaComponents> } | null = null;
  private nextTaskSequence = 0;
  private activeTaskId: string | null = null;
  private activeTaskPort: CalculationTaskPort | null = null;
  private defaultTaskPort: CalculationTaskPort | null = null;
  private readonly dateSystem: ExcelDateSystem;
  private readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  private readonly numericContext: ExcelNumericContext;
  private readonly calculationEntropySeed: string;
  private calculationCycleSequence = 0;
  private activeCalculationEntropy?: CalculationEntropyContext;
  private lastAppliedCalculationChanges: readonly CellAddress[] | null = null;
  private readonly collationContext: WorkbookCollationContext;
  private rowVisibilityResolver?: RowVisibilityResolver;
  private calculationSettings: WorkbookCalculationSettings;
  private iterationFallbackValues?: ReadonlyMap<string, FormulaValue>;
  private activeResultChangeCollector: Map<string, CellAddress> | null = null;
  private activeResultChangeBaselines: Map<string, FormulaValue> | null = null;
  private activeSpillChangeCollector: Map<string, CellAddress> | null = null;
  private activeSpillChangeBaselines: Map<string, ResolvedSpill | undefined> | null = null;
  private pendingSpillChangeBaselines = new Map<string, { address: CellAddress; spill: ResolvedSpill }>();

  private readonly cells = new Map<string, StoredCell>();
  private readonly inputRowsBySheet = new Map<string, Map<number, Set<number>>>();
  private readonly sortedInputRowsBySheet = new Map<string, readonly number[]>();
  private formulaCount = 0;

  constructor(options: FormulaEngineOptions = {}) {
    this.defaultSheetId = options.defaultSheetId ?? 'Sheet1';
    this.sheetOrder = normalizeFormulaSheetOrder(options.sheetOrder, this.defaultSheetId);
    this.calculationSettings = normalizeWorkbookCalculationSettings({
      ...DEFAULT_WORKBOOK_CALCULATION_SETTINGS,
      ...options.calculationSettings,
      mode: options.recalculationMode ?? options.calculationSettings?.mode ?? DEFAULT_WORKBOOK_CALCULATION_SETTINGS.mode,
    });
    this.recalculationMode = this.calculationSettings.mode;
    this.dateSystem = options.dateSystem ?? '1900';
    this.canonicalReferenceDate = options.canonicalReferenceDate ? structuredClone(options.canonicalReferenceDate) : undefined;
    this.numericContext = normalizeExcelNumericContext(options.numericContext ?? DEFAULT_EXCEL_NUMERIC_CONTEXT);
    this.calculationEntropySeed = options.calculationEntropySeed?.trim() || 'react-sheets-calculation';
    this.collationContext = normalizeWorkbookCollation(options.collationContext ?? DEFAULT_WORKBOOK_COLLATION);
    this.rowVisibilityResolver = options.rowVisibilityResolver;
    if (!this.defaultSheetId) throw new Error('FormulaEngine requires a default worksheet id');
    this.dependencies = new RangeIndex(this.sheetOrder);
  }

  /** Rebuild an isolated engine from a structured-clone-safe calculation snapshot. */
  static fromCalculationSnapshot(snapshot: FormulaCalculationSnapshot): FormulaEngine {
    assertFormulaCalculationSnapshot(snapshot);
    const engine = new FormulaEngine({
      defaultSheetId: snapshot.defaultSheetId,
      sheetOrder: snapshot.sheetOrder,
      recalculationMode: 'manual',
      calculationSettings: { ...snapshot.calculationSettings, mode: 'manual' },
      dateSystem: snapshot.dateSystem,
      canonicalReferenceDate: snapshot.canonicalReferenceDate,
      numericContext: snapshot.numericContext,
      calculationEntropySeed: snapshot.calculationEntropy.entropySeed,
      collationContext: snapshot.collationContext,
      rowVisibilityResolver: snapshot.visibility ? createSnapshotVisibilityResolver(snapshot.visibility) : undefined,
    });
    engine.activeCalculationEntropy = structuredClone(snapshot.calculationEntropy);
    engine.calculationCycleSequence = snapshot.calculationEntropy.cycleId;
    engine.definedNameModels = normalizeDefinedNameModels(snapshot.definedNameModels);
    engine.sheetTables = new Map(
      [...normalizeSheetTables(snapshot.sheetTables)].map(([name, table]) => [name, structuredClone(table)] as const),
    );
    for (const spillSpace of snapshot.spillSpaces) {
      const occupied = new Map(spillSpace.occupied.map((address) => [
        cellAddressKey(address),
        { row: address.row, column: address.column },
      ] as const));
      engine.spillEnvironments.set(spillSpace.sheetId, {
        rowCount: spillSpace.rowCount,
        columnCount: spillSpace.columnCount,
        isOccupied: (row, column) => occupied.has(cellAddressKey({ sheetId: spillSpace.sheetId, row, column })),
        getOccupiedAddresses: () => [...occupied.values()].map((address) => ({ ...address })),
        applyOccupiedUpdate: (address, isOccupied) => {
          const key = cellAddressKey(address);
          if (isOccupied) occupied.set(key, { row: address.row, column: address.column });
          else occupied.delete(key);
        },
      });
    }
    for (const cell of snapshot.cells) {
      if (cell.input.kind === 'formula') engine.loadFormula(cell.address, cell.input.formula);
      else engine.loadValue(cell.address, cell.input.value);
    }
    engine.pendingRecalculationRoots = new Set(snapshot.pendingRoots.map(cellAddressKey));
    engine.calculationSettings = structuredClone(snapshot.calculationSettings);
    engine.recalculationMode = snapshot.calculationSettings.mode;
    return engine;
  }

  setCell(addressInput: CellAddressInput, input: CellInput): FormulaResult {
    return 'formula' in input ? this.setFormula(addressInput, input.formula) : this.setValue(addressInput, input.value);
  }

  setValue(addressInput: CellAddressInput, value: ScalarValue): FormulaResult {
    const address = this.resolveAddress(addressInput);
    const result = this.loadValue(address, value);
    this.recordInputUpdate(address, { kind: 'value', value });
    this.markCalculationStateChanged();
    if (isAutomaticCalculationMode(this.recalculationMode)) {
      this.recalculate(address);
    } else {
      this.pendingRecalculationRoots.add(cellAddressKey(address));
    }
    return this.getCellResult(address) ?? result;
  }

  setFormula(addressInput: CellAddressInput, formula: string): FormulaResult {
    const address = this.resolveAddress(addressInput);
    const result = this.loadFormula(address, formula);
    this.recordInputUpdate(address, { kind: 'formula', formula });
    this.markCalculationStateChanged();
    if (isAutomaticCalculationMode(this.recalculationMode)) {
      this.recalculate(address);
    } else {
      this.evaluateChangedCell(address);
      this.pendingRecalculationRoots.add(cellAddressKey(address));
    }
    return this.getCellResult(address) ?? result;
  }

  clearCell(addressInput: CellAddressInput): RecalculationReport {
    const address = this.resolveAddress(addressInput);
    const key = cellAddressKey(address);
    const previous = this.cells.get(key);
    this.dependencies.remove(address);
    this.rememberSpillBaseline(address);
    this.spills.delete(spillKey(address));
    this.detachNameReferences(key);
    this.detachTableReferences(key);
    this.volatileCells.delete(key);
    this.cells.delete(key);
    this.unindexInputAddress(address);
    if (previous?.formula !== undefined) this.formulaCount = Math.max(0, this.formulaCount - 1);
    this.markFormulaTopologyChanged();
    this.recordInputUpdate(address, null);
    this.markCalculationStateChanged();
    return this.scheduleRecalculation(address) ?? { recalculated: [], results: new Map() };
  }

  getRecalculationMode(): RecalculationMode {
    return this.recalculationMode;
  }

  getCalculationSettings(): WorkbookCalculationSettings {
    return structuredClone(this.calculationSettings);
  }

  setCalculationSettings(settings: Partial<WorkbookCalculationSettings>): RecalculationReport {
    this.calculationSettings = normalizeWorkbookCalculationSettings({ ...this.calculationSettings, ...settings });
    this.recalculationMode = this.calculationSettings.mode;
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
    const affected = this.allFormulaAddresses();
    if (this.recalculationMode !== 'automatic') {
      for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
      return { recalculated: [], results: new Map() };
    }
    return this.recalculateAffected(affected);
  }

  getCanonicalReferenceDate(): CanonicalExcelDateParts | undefined {
    return this.canonicalReferenceDate ? structuredClone(this.canonicalReferenceDate) : undefined;
  }

  getDateSystem(): ExcelDateSystem {
    return this.dateSystem;
  }

  getNumericContext(): ExcelNumericContext {
    return { ...this.numericContext };
  }

  getCollationContext(): WorkbookCollationContext {
    return structuredClone(this.collationContext);
  }

  private beginCalculationEntropy(): CalculationEntropyContext {
    if (this.activeCalculationEntropy) return this.activeCalculationEntropy;
    this.calculationCycleSequence += 1;
    this.activeCalculationEntropy = createCalculationEntropyContext(this.calculationEntropySeed, this.calculationCycleSequence);
    return this.activeCalculationEntropy;
  }

  private randomForCell(address: CellAddress, functionName: string, occurrence = '0', elementIndex = 0): number | FormulaError {
    const entropy = this.activeCalculationEntropy;
    if (!entropy) return createFormulaError('#BLOCKED!', 'Volatile formula requires a calculation entropy context');
    return formulaRandom(entropy, address, functionName, occurrence, elementIndex);
  }

  /** Monotonic input/calculation generation used by derived consumers. */
  getCalculationGeneration(): number {
    return this.calculationGeneration;
  }

  /** Context revisions are stable across ordinary cell edits. */
  getCalculationContextGeneration(): number {
    return this.calculationContextGeneration;
  }

  getFormulaCount(): number {
    return this.formulaCount;
  }

  getInputAddressesInRange(range: {
    readonly sheetId: string;
    readonly startRow: number;
    readonly endRow: number;
    readonly startColumn: number;
    readonly endColumn: number;
  }): readonly CellAddress[] {
    if (!range.sheetId.trim()
      || !Number.isSafeInteger(range.startRow) || range.startRow < 0
      || !Number.isSafeInteger(range.endRow) || range.endRow < range.startRow
      || !Number.isSafeInteger(range.startColumn) || range.startColumn < 0
      || !Number.isSafeInteger(range.endColumn) || range.endColumn < range.startColumn) {
      throw new FormulaReferenceError('Calculation input range bounds are invalid');
    }
    const rows = this.inputRowsBySheet.get(range.sheetId);
    if (!rows) return [];
    let rowCoordinates = this.sortedInputRowsBySheet.get(range.sheetId);
    if (!rowCoordinates) {
      rowCoordinates = [...rows.keys()].sort((left, right) => left - right);
      this.sortedInputRowsBySheet.set(range.sheetId, rowCoordinates);
    }
    let low = 0;
    let high = rowCoordinates.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (rowCoordinates[middle]! < range.startRow) low = middle + 1;
      else high = middle;
    }
    const addresses: CellAddress[] = [];
    for (let index = low; index < rowCoordinates.length && rowCoordinates[index]! <= range.endRow; index += 1) {
      const row = rowCoordinates[index]!;
      const columns = rows.get(row)!;
      for (const column of columns) {
        if (column >= range.startColumn && column <= range.endColumn) {
          addresses.push({ sheetId: range.sheetId, row, column });
        }
      }
    }
    return addresses.sort(compareCellAddresses);
  }

  getPendingRecalculationRoots(): readonly CellAddress[] {
    return this.pendingCalculationRoots();
  }

  /**
   * Return input deltas that have not yet been acknowledged by the persistent
   * calculation Worker.  The sequence is monotonic so newer edits cannot be
   * accidentally removed when an older task completes late.
   */
  exportPendingCalculationInputs(): { inputs: readonly CalculationInputUpdate[]; inputRevision: number } {
    const entries = [...this.pendingInputUpdates.values()]
      .sort((left, right) => left.sequence - right.sequence)
      .map((entry) => structuredClone(entry.update));
    return { inputs: entries, inputRevision: this.inputUpdateSequence };
  }

  acknowledgeCalculationInputUpdates(inputRevision: number): void {
    if (!Number.isSafeInteger(inputRevision) || inputRevision < 0) return;
    for (const [key, entry] of this.pendingInputUpdates) {
      if (entry.sequence <= inputRevision) this.pendingInputUpdates.delete(key);
    }
  }

  /** Apply a batch of canonical model inputs without per-cell recalculation. */
  synchronizeInputs(updates: readonly CalculationInputUpdate[]): readonly CellAddress[] {
    const addresses: CellAddress[] = [];
    for (const update of updates) {
      const address = this.resolveAddress(update.address);
      addresses.push({ ...address });
      this.applyInputUpdate(update, true);
      addresses.push(...this.spillAnchorsAffectedBy(address));
    }
    if (updates.length === 0) return addresses;
    for (const address of addresses) this.pendingRecalculationRoots.add(cellAddressKey(address));
    this.markCalculationStateChanged();
    if (this.recalculationMode !== 'automatic') {
      for (const address of addresses) {
        const key = cellAddressKey(address);
        if (!this.cells.has(key)) continue;
        this.pendingRecalculationRoots.add(key);
        if (this.cells.get(key)?.formula !== undefined) this.evaluateChangedCell(address);
      }
    }
    return addresses;
  }

  /** Worker-only input apply path; no host journal and no nested task cancel. */
  applyCalculationTaskInputs(updates: readonly CalculationInputUpdate[]): void {
    for (const update of updates) {
      const address = this.resolveAddress(update.address);
      this.applyInputUpdate(update, false);
      this.pendingRecalculationRoots.add(cellAddressKey(address));
      const input = update.input;
      const occupied = input !== null && (input.kind === 'formula' || (input.value !== null && input.value !== ''));
      this.spillEnvironments.get(address.sheetId)?.applyOccupiedUpdate?.(address, occupied);
      for (const anchor of this.spillAnchorsAffectedBy(address)) this.pendingRecalculationRoots.add(cellAddressKey(anchor));
    }
  }

  /** Advance the calculation generation when visibility changes without cell writes. */
  notifyVisibilityChanged(): void {
    this.calculationContextGeneration += 1;
    for (const key of this.visibilityDependentCells) {
      if (this.cells.get(key)?.formula !== undefined) this.pendingRecalculationRoots.add(key);
    }
    this.markCalculationStateChanged();
  }

  /**
   * Creates the actual browser Worker transport when one is available. Node
   * and other non-browser hosts retain the explicit inline implementation.
   */
  createCalculationTaskPort(options: CalculationTaskPortOptions = {}): CalculationTaskPort {
    if (options.useWorker !== false) {
      const worker = options.workerFactory?.() ?? createBrowserCalculationWorker();
      if (worker) {
        return new BrowserCalculationTaskPort(
          worker,
          (workerContextGeneration): CalculationTaskState => {
            const pending = this.exportPendingCalculationInputs();
            const needsSnapshot = workerContextGeneration === null
              || workerContextGeneration !== this.calculationContextGeneration;
            return {
              ...(needsSnapshot ? { snapshot: this.exportCalculationSnapshot() } : {}),
              generation: this.calculationGeneration,
              inputs: pending.inputs,
              inputRevision: pending.inputRevision,
              calculationContextGeneration: this.calculationContextGeneration,
            };
          },
          (result, generation) => {
            return this.applyCalculationTaskResult(result, generation);
          },
          (inputRevision) => this.acknowledgeCalculationInputUpdates(inputRevision),
        );
      }
    }
    return new InlineCalculationTaskPort((request) => this.executeCalculationTask(request));
  }

  /**
   * Runs recalculation through the task transport. In a browser this posts an
   * isolated calculation snapshot to the real Worker; it never executes the
   * formula evaluator on the main thread. Callers must await it and refresh
   * their derived projection after completion.
   */
  async recalculateAsync(
    addressInput?: CellAddressInput | readonly CellAddressInput[],
    taskPort: CalculationTaskPort = this.defaultTaskPort ??= this.createCalculationTaskPort(),
    full = false,
  ): Promise<RecalculationReport> {
    this.lastAppliedCalculationChanges = null;
    if (this.activeTaskId && this.activeTaskPort) {
      this.activeTaskPort.cancel(this.activeTaskId);
      this.activeCalculationEntropy = undefined;
    }
    const revision = ++this.nextTaskSequence;
    const taskId = `calculation-${revision}`;
    const generation = this.calculationGeneration;
    this.activeTaskId = taskId;
    this.activeTaskPort = taskPort;
    const calculationEntropy = this.beginCalculationEntropy();
    const roots = addressInput === undefined
      ? undefined
      : (Array.isArray(addressInput) ? addressInput : [addressInput]).map((address) => this.resolveAddress(address));
    const result = await taskPort.submit({
      protocol: 'react-sheets.formula-calculation',
      version: 1,
      taskId,
      kind: 'recalculate',
      revision,
      calculationEntropy,
      ...(full ? { full: true } : {}),
      ...(roots === undefined ? {} : { roots }),
    }).finally(() => {
      if (this.activeCalculationEntropy === calculationEntropy) this.activeCalculationEntropy = undefined;
    });
    if (this.activeTaskId === taskId) {
      this.activeTaskId = null;
      this.activeTaskPort = null;
    }
    if (result.status === 'failed') throw new Error(result.error?.message ?? 'Formula calculation failed');
    if (result.status === 'cancelled' || generation !== this.calculationGeneration || !result.report) {
      return { recalculated: [], results: new Map() };
    }
    return this.recalculationReportFromTask(result.report);
  }

  cancelCalculation(): void {
    if (!this.activeTaskId || !this.activeTaskPort) return;
    this.activeTaskPort.cancel(this.activeTaskId);
    this.activeTaskId = null;
    this.activeTaskPort = null;
    this.activeCalculationEntropy = undefined;
  }

  disposeCalculationTasks(): void {
    this.cancelCalculation();
    const disposable = this.defaultTaskPort as (CalculationTaskPort & { dispose?: () => void }) | null;
    disposable?.dispose?.();
    this.defaultTaskPort = null;
  }

  executeCalculationTask(request: CalculationTaskRequest): CalculationTaskReport {
    assertCalculationTaskRequest(request);
    this.activeCalculationEntropy = request.calculationEntropy;
    if (request.calculationEntropy) this.calculationCycleSequence = Math.max(this.calculationCycleSequence, request.calculationEntropy.cycleId);
    const reports = request.full
      ? [this.recalculateAllFormulas()]
      : request.roots && request.roots.length > 0
        ? [this.recalculateRoots(request.roots)]
        : [this.recalculateSmart()];
    if (request.roots && request.roots.length > 0) {
      for (const root of request.roots) this.pendingRecalculationRoots.delete(cellAddressKey(this.resolveAddress(root)));
    }
    if (this.recalculationMode === 'automatic') this.pendingRecalculationRoots.clear();
    const recalculated: CellAddress[] = [];
    const changedAddresses = new Map<string, CellAddress>();
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
      for (const address of report.changedAddresses ?? []) {
        changedAddresses.set(cellAddressKey(address), { ...address });
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
    return {
      recalculated,
      changedAddresses: [...changedAddresses.values()].sort(compareCellAddresses),
      results: taskResults,
      spills: [...this.spills.values()].map(copySpill),
      pendingRoots: this.pendingCalculationRoots(),
    };
  }

  /** Excel F9 smart recalculation: only dirty roots, their dependents, and volatile formulas. */
  private recalculateSmart(): RecalculationReport {
    const affected = this.pendingRecalculationRoots.size > 0
      ? this.collectAffectedFromRoots(this.pendingRecalculationRoots)
      : new Map<string, CellAddress>();
    for (const key of this.volatileCells) {
      const cell = this.cells.get(key);
      if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
    }
    const report = this.recalculateAffected(affected);
    this.pendingRecalculationRoots.clear();
    return report;
  }

  /** Recalculate a union of roots once; never traverse the same dependency subtree per root. */
  private recalculateRoots(addressInputs: readonly CellAddressInput[]): RecalculationReport {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
      const affected = new Map<string, CellAddress>();
      for (const addressInput of addressInputs) {
        const address = this.resolveAddress(addressInput);
        for (const [key, dependent] of this.collectAffected(address)) affected.set(key, dependent);
      }
      for (const key of this.volatileCells) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
      }
      return this.recalculateAffected(affected);
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  /** Return the structured-clone-safe inputs for a Worker calculation task. */
  exportCalculationSnapshot(): FormulaCalculationSnapshot {
    const cells = [...this.cells.values()]
      .map((cell) => ({
        address: { ...cell.address },
        input: cell.formula === undefined
          ? { kind: 'value' as const, value: cell.result.value as ScalarValue }
          : { kind: 'formula' as const, formula: cell.formula },
      }))
      .sort((left, right) => compareCellAddresses(left.address, right.address));

    const occupiedBySheet = new Map<string, Map<string, CellAddress>>();
    for (const cell of this.cells.values()) {
      if (!isOccupiedInput(cell)) continue;
      const occupied = occupiedBySheet.get(cell.address.sheetId) ?? new Map<string, CellAddress>();
      occupied.set(cellAddressKey(cell.address), { ...cell.address });
      occupiedBySheet.set(cell.address.sheetId, occupied);
    }
    const spillSpaces = [...this.spillEnvironments.entries()]
      .map(([sheetId, environment]) => {
        const occupied = occupiedBySheet.get(sheetId) ?? new Map<string, CellAddress>();
        for (const address of environment.getOccupiedAddresses?.() ?? []) {
          if (address.row < 0 || address.column < 0) continue;
          occupied.set(cellAddressKey({ sheetId, row: address.row, column: address.column }), {
            sheetId,
            row: address.row,
            column: address.column,
          });
        }
        return {
          sheetId,
          rowCount: environment.rowCount,
          columnCount: environment.columnCount,
          occupied: [...occupied.values()].sort(compareCellAddresses),
        };
      })
      .sort((left, right) => left.sheetId.localeCompare(right.sheetId));

    return {
      defaultSheetId: this.defaultSheetId,
      sheetOrder: this.sheetOrder.map((sheet) => ({ ...sheet })),
      calculationSettings: structuredClone(this.calculationSettings),
      dateSystem: this.dateSystem,
      canonicalReferenceDate: this.canonicalReferenceDate ? structuredClone(this.canonicalReferenceDate) : undefined,
      numericContext: { ...this.numericContext },
      calculationEntropy: this.activeCalculationEntropy ?? createCalculationEntropyContext(this.calculationEntropySeed, this.calculationCycleSequence),
      collationContext: structuredClone(this.collationContext),
      ...(this.rowVisibilityResolver?.snapshot ? { visibility: structuredClone(this.rowVisibilityResolver.snapshot()) } : {}),
      cells,
      definedNameModels: this.getDefinedNameModels(),
      sheetTables: this.getSheetTables(),
      spillSpaces,
      pendingRoots: this.pendingCalculationRoots(),
    };
  }

  /**
   * Apply only a result produced from the current calculation generation.
   * A late task cannot overwrite a newer edit, even if it reaches the port.
   */
  applyCalculationTaskResult(result: CalculationTaskResult, generation: number): boolean {
    if (generation !== this.calculationGeneration || result.status !== 'completed' || !result.report) return false;
    const changedAddresses = new Map<string, CellAddress>();
    for (const entry of result.report.results) {
      const key = cellAddressKey(entry.address);
      const cell = this.cells.get(key);
      if (!cell || cell.formula === undefined || cell.formula !== entry.formula) {
        throw new Error(`CALCULATION_RESULT_FORMULA_MISMATCH: worker result does not match ${key}`);
      }
      if (!sameCalculationValue(cell.result.value, entry.value)) {
        changedAddresses.set(key, { ...entry.address });
      }
      cell.result = {
        value: entry.value,
        formula: cell.formula,
        ast: cell.ast,
        dependencies: entry.dependencies.map(copyDependency),
      };
    }
    if (result.report.spills !== undefined) {
      const nextSpills = new Map(result.report.spills.map((spill) => [
        spillKey({ sheetId: spill.sheetId, row: spill.anchor.row, column: spill.anchor.column }),
        copySpill(spill),
      ] as const));
      for (const [key, spill] of this.spills) {
        if (!sameCalculationValue(spill, nextSpills.get(key))) {
          changedAddresses.set(cellAddressKey({ sheetId: spill.sheetId, ...spill.anchor }), { sheetId: spill.sheetId, ...spill.anchor });
        }
      }
      for (const [key, spill] of nextSpills) {
        if (!sameCalculationValue(spill, this.spills.get(key))) {
          changedAddresses.set(cellAddressKey({ sheetId: spill.sheetId, ...spill.anchor }), { sheetId: spill.sheetId, ...spill.anchor });
        }
      }
      for (const { address, spill } of this.pendingSpillChangeBaselines.values()) {
        if (!sameCalculationValue(spill, nextSpills.get(spillKey(address)))) {
          changedAddresses.set(cellAddressKey(address), { ...address });
        }
      }
      this.spills.clear();
      for (const [key, spill] of nextSpills) this.spills.set(key, spill);
    }
    if (result.report.pendingRoots !== undefined) {
      this.pendingRecalculationRoots = new Set(result.report.pendingRoots.map(cellAddressKey));
    }
    this.pendingSpillChangeBaselines.clear();
    this.lastAppliedCalculationChanges = [...changedAddresses.values()].sort(compareCellAddresses);
    return true;
  }

  setRecalculationMode(mode: RecalculationMode): void {
    this.recalculationMode = mode;
    this.calculationSettings = normalizeWorkbookCalculationSettings({ ...this.calculationSettings, mode });
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
  }

  hasPendingRecalculation(): boolean {
    return this.pendingRecalculationRoots.size > 0;
  }

  setSheetTables(tables: readonly SheetTableRef[], recalculate = true): RecalculationReport {
    const nextTables = new Map(
      [...normalizeSheetTables(tables)].map(([name, table]) => [name, structuredClone(table)] as const),
    );
    const changedTables = changedMapKeys(this.sheetTables, nextTables);
    if (changedTables.size === 0) return { recalculated: [], results: new Map() };
    this.sheetTables = nextTables;
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
    const affected = this.formulasReferencing(this.tableReferenceIndex, changedTables);
    this.refreshFormulaDependencies(affected);
    if (!recalculate || this.recalculationMode !== 'automatic') {
      for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
      return { recalculated: [], results: new Map() };
    }
    return this.recalculateAffected(affected);
  }

  getSheetTables(): readonly SheetTableRef[] {
    return [...this.sheetTables.values()].map((table) => structuredClone(table));
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

  /** Return all authored formula cells in deterministic address order. */
  listFormulaCells(): readonly CellAddress[] {
    return [...this.cells.values()]
      .filter((cell) => cell.formula !== undefined)
      .map((cell) => ({ ...cell.address }))
      .sort(compareCellAddresses);
  }

  /** Return a stable, cloned view of authored formulas and their current results. */
  getFormulaEntries(): readonly FormulaCellEntry[] {
    return [...this.cells.values()]
      .filter((cell): cell is StoredCell & { formula: string } => cell.formula !== undefined)
      .map((cell) => ({
        address: { ...cell.address },
        formula: cell.formula,
        value: structuredClone(cell.result.value),
        ...(cell.ast === undefined ? {} : { ast: structuredClone(cell.ast) }),
        dependencies: cell.result.dependencies.map(copyDependency),
      }))
      .sort((left, right) => compareCellAddresses(left.address, right.address));
  }

  /** Return the parsed AST for a formula without exposing mutable engine state. */
  getFormulaAst(addressInput: CellAddressInput): FormulaAst | undefined {
    const cell = this.cells.get(cellAddressKey(this.resolveAddress(addressInput)));
    return cell?.ast === undefined ? undefined : structuredClone(cell.ast);
  }

  /** Evaluate one formula and expose a data-only, ordered AST trace for auditing. */
  evaluateFormulaWithTrace(addressInput: CellAddressInput): FormulaEvaluationTrace | undefined {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
    const address = this.resolveAddress(addressInput);
    const cell = this.cells.get(cellAddressKey(address));
    if (!cell?.formula || !cell.ast) {
      return cell?.parseError ? { value: structuredClone(cell.parseError), steps: [] } : undefined;
    }
    const cache = new Map<string, FormulaValue>();
    const visiting = new Set<string>();
    let trace: FormulaEvaluationTrace;
    try {
      trace = evaluateFormulaWithTrace(cell.ast, this.createEvaluationContext(cell, cache, visiting, []));
    } catch (error) {
      const value = error instanceof FormulaReferenceError
        ? createFormulaError('#REF!', error.message)
        : createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Formula evaluation failed');
      trace = { value, steps: [] };
    }
    cell.result = { value: trace.value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
    return trace;
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  getDependencies(addressInput: CellAddressInput): readonly FormulaDependency[] {
    return this.dependencies.getDependencies(this.resolveAddress(addressInput));
  }

  getDependents(addressInput: CellAddressInput): readonly CellAddress[] {
    return this.dependencies.getDependents(this.resolveAddress(addressInput));
  }

  /**
   * Index a non-calculation formula owner for structural rewrites without
   * adding it to the calculation dependency graph. Names and table references
   * remain owned by their canonical name/table models; only explicit cell
   * references are spatially indexed here.
   */
  setStructuralFormulaReference(addressInput: CellAddressInput, sourceId: string, formula: string): void {
    const address = this.resolveAddress(addressInput);
    let dependencies: readonly FormulaDependency[];
    try {
      const ast = this.parseFormula(formula);
      dependencies = collectFormulaDependencies(ast, address, { sheetOrder: this.sheetOrder });
    } catch {
      this.dependencies.setStructuralReference(address, sourceId, [], true);
      return;
    }
    try {
      this.dependencies.setStructuralReference(address, sourceId, dependencies);
    } catch (error) {
      if (!(error instanceof FormulaReferenceError)) throw error;
      this.dependencies.setStructuralReference(address, sourceId, [], true);
    }
  }

  removeStructuralFormulaReference(addressInput: CellAddressInput, sourceId: string): boolean {
    return this.dependencies.removeStructuralReference(this.resolveAddress(addressInput), sourceId);
  }

  /**
   * Set the canonical scoped name collection. A workbook Record remains a
   * narrow input compatibility form; it is immediately normalized into the
   * same scoped collection and is never stored separately.
   */
  setDefinedNames(names: Record<string, string> | readonly FormulaDefinedName[]): RecalculationReport {
    if (Array.isArray(names)) return this.setDefinedNameModels(names as readonly FormulaDefinedName[]);
    const models = Object.entries(normalizeDefinedNames(names as Record<string, string>))
      .map(([name, formula]) => ({ name, formula, scope: 'workbook' as const }));
    return this.setDefinedNameModels(models);
  }

  setDefinedNameModels(names: readonly FormulaDefinedName[], recalculate = true): RecalculationReport {
    const nextNames = normalizeDefinedNameModels(names);
    const changedNames = changedDefinedNameTokens(
      this.definedNameModels,
      nextNames,
      (entry) => this.definedNameIdentity(entry),
    );
    if (changedNames.size === 0) return { recalculated: [], results: new Map() };
    this.definedNameModels = nextNames;
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
    const affected = this.formulasReferencing(this.nameIndex, changedNames);
    this.refreshFormulaDependencies(affected);
    if (affected.size === 0) return { recalculated: [], results: new Map() };
    if (!recalculate || this.recalculationMode !== 'automatic') {
      for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
      return { recalculated: [], results: new Map() };
    }
    return this.recalculateAffected(affected);
  }

  setSpillEnvironment(sheetId: string, environment: SpillEnvironment | undefined): void {
    if (!environment) this.spillEnvironments.delete(sheetId);
    else this.spillEnvironments.set(sheetId, environment);
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
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
    const result: Record<string, string> = {};
    for (const entry of this.definedNameModels) {
      if (entry.scope === 'workbook') result[entry.name.toUpperCase()] = entry.formula;
    }
    return result;
  }

  getDefinedNameModels(): FormulaDefinedName[] {
    return structuredClone(this.definedNameModels);
  }

  /** 清空全部公式与缓存(结构操作后整体重建前调用) */
  reset(): void {
    this.cells.clear();
    this.inputRowsBySheet.clear();
    this.sortedInputRowsBySheet.clear();
    this.spills.clear();
    this.nameIndex.clear();
    this.cellNameRefs.clear();
    this.tableReferenceIndex.clear();
    this.cellTableRefs.clear();
    this.volatileCells.clear();
    this.visibilityDependentCells.clear();
    this.pendingRecalculationRoots.clear();
    this.sheetTables.clear();
    this.dependencies.clear?.();
    this.formulaCount = 0;
    this.markFormulaTopologyChanged();
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
  }

  recalculateCell(addressInput: CellAddressInput): RecalculationReport {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
    const address = this.resolveAddress(addressInput);
    const affected = new Map<string, CellAddress>();
    const key = cellAddressKey(address);
    const cell = this.cells.get(key);
    if (cell?.formula !== undefined) affected.set(key, { ...address });
    return this.recalculateAffected(affected);
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  recalculate(addressInput?: CellAddressInput): RecalculationReport {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
      if (addressInput !== undefined) {
        const affected = this.collectAffected(this.resolveAddress(addressInput));
        for (const key of this.volatileCells) {
          const cell = this.cells.get(key);
          if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
        }
        return this.recalculateAffected(affected);
      }

      const affected = this.recalculationMode !== 'automatic' && this.pendingRecalculationRoots.size > 0
        ? this.collectAffectedFromRoots(this.pendingRecalculationRoots)
        : this.allFormulaAddresses();
      for (const key of this.volatileCells) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined) affected.set(key, { ...cell.address });
      }
      const report = this.recalculateAffected(affected);
      this.pendingRecalculationRoots.clear();
      return report;
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  private scheduleRecalculation(address: CellAddress): RecalculationReport | undefined {
    if (isAutomaticCalculationMode(this.recalculationMode)) return this.recalculate(address);
    this.pendingRecalculationRoots.add(cellAddressKey(address));
    return undefined;
  }

  private evaluateChangedCell(address: CellAddress): void {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
    const key = cellAddressKey(address);
    const cell = this.cells.get(key);
    if (!cell?.formula || !cell.ast) return;
    const cache = new Map<string, FormulaValue>();
    const visiting = new Set<string>();
    this.evaluateCell(address, cache, visiting);
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  private collectAffectedFromRoots(roots: ReadonlySet<string>): Map<string, CellAddress> {
    const affected = new Map<string, CellAddress>();
    const queue: CellAddress[] = [];
    for (const key of roots) {
      const cell = this.cells.get(key);
      if (!cell) continue;
      if (affected.has(key)) continue;
      affected.set(key, { ...cell.address });
      queue.push({ ...cell.address });
    }
    let head = 0;
    while (head < queue.length) {
      const current = queue[head++];
      if (!current) continue;
      this.addSpillAnchorsAffectedBy(current, affected, queue);
      for (const dependent of this.dependencies.getDependents(current)) {
        const key = cellAddressKey(dependent);
        if (affected.has(key)) continue;
        affected.set(key, dependent);
        queue.push(dependent);
      }
    }
    return affected;
  }

  private applyInputUpdate(update: CalculationInputUpdate, trackForWorker: boolean): void {
    const address = this.resolveAddress(update.address);
    const addressKey = cellAddressKey(address);
    if (update.input === null) {
      const key = addressKey;
      this.dependencies.remove(address);
      this.rememberSpillBaseline(address);
      this.spills.delete(spillKey(address));
      this.detachNameReferences(key);
      this.detachTableReferences(key);
      this.volatileCells.delete(key);
      const previous = this.cells.get(key);
      this.cells.delete(key);
      this.unindexInputAddress(address);
      this.visibilityDependentCells.delete(key);
      if (previous?.formula !== undefined) this.formulaCount = Math.max(0, this.formulaCount - 1);
      this.markFormulaTopologyChanged();
    } else if (update.input.kind === 'formula') {
      this.loadFormula(address, update.input.formula);
    } else {
      this.loadValue(address, update.input.value);
    }
    if (trackForWorker) this.recordInputUpdate(address, update.input);
  }

  private rememberSpillBaseline(address: CellAddress): void {
    const key = cellAddressKey(address);
    const spill = this.spills.get(spillKey(address));
    if (spill && !this.pendingSpillChangeBaselines.has(key)) {
      this.pendingSpillChangeBaselines.set(key, { address: { ...address }, spill: copySpill(spill) });
    }
  }

  private recordInputUpdate(address: CellAddress, input: CalculationInputUpdate['input']): void {
    this.inputUpdateSequence += 1;
    const update: CalculationInputUpdate = {
      address: { ...address },
      input: input === null ? null : structuredClone(input),
    };
    this.pendingInputUpdates.set(cellAddressKey(address), { sequence: this.inputUpdateSequence, update });
  }

  private markFormulaTopologyChanged(): void {
    this.formulaTopologyGeneration += 1;
    this.cachedCircularComponents = null;
  }

  private indexInputAddress(address: CellAddress): void {
    let rows = this.inputRowsBySheet.get(address.sheetId);
    if (!rows) {
      rows = new Map<number, Set<number>>();
      this.inputRowsBySheet.set(address.sheetId, rows);
    }
    let columns = rows.get(address.row);
    if (!columns) {
      columns = new Set<number>();
      rows.set(address.row, columns);
      this.sortedInputRowsBySheet.delete(address.sheetId);
    }
    columns.add(address.column);
  }

  private unindexInputAddress(address: CellAddress): void {
    const rows = this.inputRowsBySheet.get(address.sheetId);
    const columns = rows?.get(address.row);
    columns?.delete(address.column);
    if (columns?.size === 0) {
      rows?.delete(address.row);
      this.sortedInputRowsBySheet.delete(address.sheetId);
    }
    if (rows?.size === 0) this.inputRowsBySheet.delete(address.sheetId);
  }

  private loadValue(address: CellAddress, value: ScalarValue): FormulaResult {
    const key = cellAddressKey(address);
    const previous = this.cells.get(key);
    this.dependencies.remove(address);
    this.rememberSpillBaseline(address);
    this.spills.delete(spillKey(address));
    this.detachNameReferences(key);
    this.detachTableReferences(key);
    this.volatileCells.delete(key);
    this.visibilityDependentCells.delete(key);
    const result: FormulaResult = { value, dependencies: [] };
    this.cells.set(key, { address: { ...address }, result });
    this.indexInputAddress(address);
    if (previous?.formula !== undefined) {
      this.formulaCount = Math.max(0, this.formulaCount - 1);
      this.markFormulaTopologyChanged();
    }
    return result;
  }

  private loadFormula(address: CellAddress, formula: string): FormulaResult {
    const key = cellAddressKey(address);
    const previous = this.cells.get(key);
    this.rememberSpillBaseline(address);
    this.spills.delete(spillKey(address));
    let ast: FormulaAst | undefined;
    let formulaDependencies: readonly FormulaDependency[] = [];
    let parseError: FormulaError | undefined;

    try {
      const parsed = this.parseFormula(formula);
      const extractedDependencies = collectFormulaDependencies(parsed, address, {
        sheetTables: this.sheetTables,
        sheetOrder: this.sheetOrder,
      });
      const expandedDependencies = this.expandNameDependencies(extractedDependencies, address, new Set<string>());
      this.dependencies.set(address, expandedDependencies);
      ast = parsed;
      formulaDependencies = expandedDependencies;
    } catch (error) {
      parseError = formulaErrorFrom(error);
      this.dependencies.set(address, [], true);
    }

    const result: FormulaResult = parseError
      ? { value: parseError, formula, dependencies: [] }
      : { value: null, formula, ast, dependencies: formulaDependencies };
    this.cells.set(key, { address: { ...address }, formula, ast, parseError, result });
    this.indexInputAddress(address);
    this.updateFormulaMetadata(key, ast, formulaDependencies);
    if (previous?.formula === undefined) this.formulaCount += 1;
    this.markFormulaTopologyChanged();
    return result;
  }

  private pendingCalculationRoots(): CellAddress[] {
    return [...this.pendingRecalculationRoots]
      .map((key) => this.cells.get(key)?.address)
      .filter((address): address is CellAddress => address !== undefined)
      .map((address) => ({ ...address }))
      .sort(compareCellAddresses);
  }

  private recalculationReportFromTask(report: CalculationTaskReport): RecalculationReport {
    const results = new Map<string, FormulaResult>();
    for (const entry of report.results) {
      const cell = this.cells.get(cellAddressKey(entry.address));
      if (!cell || cell.formula === undefined || cell.formula !== entry.formula) {
        throw new Error(`CALCULATION_RESULT_FORMULA_MISMATCH: task result does not match ${cellAddressKey(entry.address)}`);
      }
      results.set(cellAddressKey(entry.address), cell.result);
    }
    return {
      recalculated: report.recalculated.map((address) => ({ ...address })),
      changedAddresses: (this.lastAppliedCalculationChanges ?? report.changedAddresses ?? []).map((address) => ({ ...address })),
      results,
    };
  }

  private markCalculationStateChanged(): void {
    this.calculationGeneration += 1;
    this.activeCalculationEntropy = undefined;
    if (this.activeTaskId && this.activeTaskPort) {
      this.activeTaskPort.cancel(this.activeTaskId);
      this.activeTaskId = null;
      this.activeTaskPort = null;
    }
  }

  private recalculateAffected(affected: Map<string, CellAddress>): RecalculationReport {
    const previousCollector = this.activeResultChangeCollector;
    const previousResultBaselines = this.activeResultChangeBaselines;
    const previousSpillCollector = this.activeSpillChangeCollector;
    const previousSpillBaselines = this.activeSpillChangeBaselines;
    const resultChanges = new Map<string, CellAddress>();
    const resultBaselines = new Map<string, FormulaValue>();
    const spillChanges = new Map<string, CellAddress>();
    const spillBaselines = new Map<string, ResolvedSpill | undefined>(
      [...this.pendingSpillChangeBaselines].map(([key, entry]) => [key, entry.spill] as const),
    );
    this.activeResultChangeCollector = resultChanges;
    this.activeResultChangeBaselines = resultBaselines;
    this.activeSpillChangeCollector = spillChanges;
    this.activeSpillChangeBaselines = spillBaselines;
    try {
      return this.recalculateAffectedCore(affected, resultChanges, spillChanges);
    } finally {
      this.activeResultChangeCollector = previousCollector;
      this.activeResultChangeBaselines = previousResultBaselines;
      this.activeSpillChangeCollector = previousSpillCollector;
      this.activeSpillChangeBaselines = previousSpillBaselines;
    }
  }

  private recalculateAffectedCore(
    affected: Map<string, CellAddress>,
    resultChanges: Map<string, CellAddress>,
    spillChanges: Map<string, CellAddress>,
  ): RecalculationReport {
    const evaluationCache = new Map<string, FormulaValue>();
    const recalculated: CellAddress[] = [];
    const results = new Map<string, FormulaResult>();

    const circularComponents = this.getCircularComponents().filter((component) =>
      component.cyclic && component.members.some((address) => affected.has(cellAddressKey(address))),
    );
    const handledCircularCells = new Set<string>();
    const circularFallback = new Map<string, FormulaValue>();
    for (const component of circularComponents) {
      for (const address of component.members) circularFallback.set(cellAddressKey(address), this.cells.get(cellAddressKey(address))?.result.value ?? null);
    }
    this.iterationFallbackValues = circularFallback;
    try {
      for (const component of circularComponents) {
        this.evaluateCircularComponent(component.members, evaluationCache);
        for (const address of component.members) handledCircularCells.add(cellAddressKey(address));
      }
    } finally {
      this.iterationFallbackValues = undefined;
    }

    const affectedEntries = [...affected.entries()]
      .map(([key, address]) => ({ key, address }))
      .sort((left, right) => compareCellAddresses(left.address, right.address));
    for (const { key, address } of affectedEntries) {
      const cell = this.cells.get(key);
      if (cell?.formula === undefined) continue;
      if (!handledCircularCells.has(key)) this.evaluateCell(address, evaluationCache, new Set<string>());
      recalculated.push({ ...address });
      const result = this.cells.get(key)?.result;
      if (result) results.set(key, result);
    }

    // Recursive evaluation can update formula prerequisites not present in a
    // partial root's dependent set. Include every evaluated formula so the
    // Worker result can faithfully update the host cache without recomputing.
    for (const key of evaluationCache.keys()) {
      const cell = this.cells.get(key);
      if (cell?.formula === undefined) continue;
      const result = cell.result;
      results.set(key, result);
    }

    for (const { address, spill } of this.pendingSpillChangeBaselines.values()) {
      this.recordSpillProjectionChange(address, spill, this.spills.get(spillKey(address)));
    }
    this.pendingSpillChangeBaselines.clear();

    return {
      recalculated,
      changedAddresses: [...new Map([...resultChanges, ...spillChanges]).values()].sort(compareCellAddresses),
      results,
    };
  }

  private getCircularComponents(): ReturnType<typeof findFormulaComponents> {
    if (this.cachedCircularComponents?.generation === this.formulaTopologyGeneration) {
      return this.cachedCircularComponents.components;
    }
    const graphNodes = [...this.cells.values()]
      .filter((cell) => cell.formula !== undefined)
      .map((cell) => ({ address: cell.address, dependencies: cell.result.dependencies }));
    const components = findFormulaComponents(graphNodes);
    this.cachedCircularComponents = { generation: this.formulaTopologyGeneration, components };
    return components;
  }

  private evaluateCircularComponent(
    members: readonly CellAddress[],
    cache: Map<string, FormulaValue>,
  ): void {
    const orderedMembers = [...members].sort(compareCellAddresses);
    if (!this.calculationSettings.iterativeCalculation) {
      for (const address of orderedMembers) {
        const key = cellAddressKey(address);
        const cell = this.cells.get(key);
        if (!cell?.formula) continue;
        const value = createFormulaError('#NUM!', `Circular reference component: ${orderedMembers.map(cellAddressKey).join(',')}`);
        const previousValue = cell.result.value;
        cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
        this.recordCalculationResultChange(address, previousValue, value);
        cache.set(key, value);
        this.spills.delete(spillKey(address));
      }
      return;
    }

    let previous = new Map<string, FormulaValue>(orderedMembers.map((address) => {
      const key = cellAddressKey(address);
      return [key, this.iterationFallbackValues?.get(key) ?? this.cells.get(key)?.result.value ?? null];
    }));
    for (let pass = 0; pass < this.calculationSettings.maximumIterations; pass += 1) {
      const iterationCache = new Map<string, FormulaValue>();
      for (const address of orderedMembers) this.evaluateCell(address, iterationCache, new Set<string>());
      const current = new Map<string, FormulaValue>();
      for (const address of orderedMembers) {
        const key = cellAddressKey(address);
        current.set(key, this.cells.get(key)?.result.value ?? null);
      }
      const delta = Math.max(...orderedMembers.map((address) => formulaValueDelta(previous.get(cellAddressKey(address)), current.get(cellAddressKey(address)))));
      previous = current;
      for (const [key, value] of current) cache.set(key, value);
      if (delta <= this.calculationSettings.maximumChange) break;
      this.iterationFallbackValues = new Map([...(this.iterationFallbackValues ?? []), ...current]);
    }
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
    let head = 0;
    while (head < queue.length) {
      const current = queue[head++];
      if (!current) continue;
      this.addSpillAnchorsAffectedBy(current, affected, queue);
      for (const dependent of this.dependencies.getDependents(current)) {
        const key = cellAddressKey(dependent);
        if (affected.has(key)) continue;
        affected.set(key, dependent);
        queue.push(dependent);
      }
    }
    return affected;
  }

  private recalculateAllFormulas(): RecalculationReport {
    const ownsEntropy = this.activeCalculationEntropy === undefined;
    this.beginCalculationEntropy();
    try {
      const report = this.recalculateAffected(this.allFormulaAddresses());
      this.pendingRecalculationRoots.clear();
      return report;
    } finally {
      if (ownsEntropy) this.activeCalculationEntropy = undefined;
    }
  }

  private spillAnchorsAffectedBy(address: CellAddress): CellAddress[] {
    const anchors: CellAddress[] = [];
    for (const spill of this.spills.values()) {
      if (spill.sheetId !== address.sheetId
        || (spill.anchor.row === address.row && spill.anchor.column === address.column)
        || address.row < spill.range.startRow || address.row > spill.range.endRow
        || address.column < spill.range.startColumn || address.column > spill.range.endColumn) continue;
      const anchor = { sheetId: spill.sheetId, row: spill.anchor.row, column: spill.anchor.column };
      const key = cellAddressKey(anchor);
      const anchorCell = this.cells.get(key);
      if (anchorCell?.formula === undefined) continue;
      anchors.push(anchor);
    }
    return anchors;
  }

  private addSpillAnchorsAffectedBy(address: CellAddress, affected: Map<string, CellAddress>, queue: CellAddress[]): void {
    for (const anchor of this.spillAnchorsAffectedBy(address)) {
      const key = cellAddressKey(anchor);
      if (affected.has(key)) continue;
      affected.set(key, anchor);
      queue.push(anchor);
    }
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
    overrides: readonly FormulaCellOverride[] = [],
  ): FormulaValue {
    const key = cellAddressKey(address);
    const cached = cache.get(key);
    if (cached !== undefined) return cached;
    if (visiting.has(key)) {
      const fallback = this.iterationFallbackValues?.get(key);
      return fallback === undefined
        ? createFormulaError('#NUM!', 'Circular reference requires iterative calculation')
        : fallback;
    }

    const cell = this.cells.get(key);
    if (!cell) {
      cache.set(key, null);
      return null;
    }
    if (cell.parseError) {
      const previousValue = cell.result.value;
      cache.set(key, cell.parseError);
      cell.result = { value: cell.parseError, formula: cell.formula, dependencies: cell.result.dependencies };
      this.recordCalculationResultChange(address, previousValue, cell.parseError);
      return cell.parseError;
    }
    if (!cell.ast) {
      cache.set(key, cell.result.value);
      return cell.result.value;
    }

    visiting.add(key);
    let value: FormulaValue;
    try {
      value = evaluateFormula(cell.ast, this.createEvaluationContext(cell, cache, visiting, overrides));
    } catch (error) {
      value = error instanceof FormulaReferenceError
        ? createFormulaError('#REF!', error.message)
        : createFormulaError('#VALUE!', error instanceof Error ? error.message : 'Formula evaluation failed');
    } finally {
      visiting.delete(key);
    }

    const previousValue = cell.result.value;
    cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
    this.refreshSpill(address, value, cache);
    const displayValue = this.cells.get(key)?.result.value ?? value;
    this.recordCalculationResultChange(address, previousValue, displayValue);
    cache.set(key, displayValue);
    return displayValue;
  }

  private recordCalculationResultChange(address: CellAddress, before: FormulaValue, after: FormulaValue): void {
    const collector = this.activeResultChangeCollector;
    const baselines = this.activeResultChangeBaselines;
    if (!collector || !baselines) return;
    const key = cellAddressKey(address);
    const baseline = baselines.has(key) ? baselines.get(key)! : before;
    if (!baselines.has(key)) baselines.set(key, before);
    if (sameCalculationValue(baseline, after)) collector.delete(key);
    else collector.set(key, { ...address });
  }

  private recordSpillProjectionChange(address: CellAddress, before: ResolvedSpill | undefined, after: ResolvedSpill | undefined): void {
    const collector = this.activeSpillChangeCollector;
    const baselines = this.activeSpillChangeBaselines;
    if (!collector || !baselines) return;
    const key = cellAddressKey(address);
    const baseline = baselines.has(key) ? baselines.get(key) : before;
    if (!baselines.has(key)) baselines.set(key, before);
    if (sameCalculationValue(baseline, after)) collector.delete(key);
    else collector.set(key, { ...address });
  }

  private createEvaluationContext(
    cell: StoredCell,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[],
  ): FormulaEvaluationContext {
    return {
        currentCell: cell.address,
        sheetOrder: this.sheetOrder,
        dateSystem: this.dateSystem,
        canonicalReferenceDate: this.canonicalReferenceDate,
        numericContext: this.numericContext,
        collationContext: this.collationContext,
        rowVisibility: this.rowVisibilityResolver,
        readFormulaKind: (reference) => this.formulaKindAt(reference),
        random: (functionName, occurrence, elementIndex) => this.randomForCell(cell.address, functionName, occurrence, elementIndex),
        readCell: (reference) => this.readCellWithOverrides(reference, cache, visiting, overrides),
        readRange: (range) => this.readRange(range, cache, visiting, overrides),
        readRangeMatrix: (range) => this.readRangeMatrix(range, cache, visiting, overrides),
        readSpillRange: (anchor) => {
          this.evaluateCell(anchor, cache, visiting, overrides);
          const spill = this.spills.get(spillKey(anchor));
          return spill ? {
            kind: 'range' as const,
            start: { sheetId: spill.range.sheetId, row: spill.range.startRow, column: spill.range.startColumn },
            end: { sheetId: spill.range.sheetId, row: spill.range.endRow, column: spill.range.endColumn },
          } : undefined;
        },
        readSpillValue: (address) => {
          this.evaluateCell(address, cache, visiting, overrides);
          return this.getSpillValueAt(address.sheetId, address.row, address.column);
        },
        resolveName: (name) => this.resolveDefinedName(name, cell.address, cache, visiting),
        resolveTableReference: (tableName, request) => {
          const resolved = resolveSheetTableReference(tableName, request, cell.address, this.sheetTables);
          if (isFormulaError(resolved)) return resolved;
          if ('start' in resolved && 'end' in resolved) return { kind: 'range', range: resolved };
          return this.evaluateCell(resolved as CellAddress, cache, visiting, overrides);
        },
        resolveReference: (reference) => this.resolveReference(reference, cell.address),
        evaluateWithCellOverrides: (ast, nestedOverrides) => evaluateFormula(ast, this.createEvaluationContext(cell, new Map<string, FormulaValue>(), new Set<string>(), [...overrides, ...nestedOverrides])),
      };
  }

  private readCellWithOverrides(
    address: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[],
  ): FormulaValue {
    const override = overrides.find((candidate) => cellAddressKey(candidate.address) === cellAddressKey(address));
    return override ? structuredClone(override.value) : this.evaluateCell(address, cache, visiting, overrides);
  }

  private refreshSpill(address: CellAddress, value: FormulaValue, cache?: Map<string, FormulaValue>): void {
    const key = spillKey(address);
    const previous = this.spills.get(key);
    if (!isSpillMatrix(value)) {
      this.spills.delete(key);
      if (previous) this.recordSpillProjectionChange(address, previous, undefined);
      return;
    }
    const environment = this.spillEnvironments.get(address.sheetId);
    if (!environment) {
      this.spills.delete(key);
      if (previous) this.recordSpillProjectionChange(address, previous, undefined);
      return;
    }
    environment.ensureExtent?.(address.row + value.length, address.column + Math.max(0, ...value.map((row) => row.length)));
    const spill = resolveSpill({
      sheetId: address.sheetId,
      anchor: { row: address.row, column: address.column },
      values: value,
      rowCount: environment.rowCount,
      columnCount: environment.columnCount,
      isOccupied: environment.isOccupied,
    });
    this.spills.set(key, spill);
    if (!sameCalculationValue(previous, spill)) this.recordSpillProjectionChange(address, previous, spill);
    if (spill.state === 'blocked') {
      const display = anchorDisplayValue(spill, value);
      const cell = this.cells.get(cellAddressKey(address));
      if (cell) cell.result = { ...cell.result, value: display };
      cache?.set(cellAddressKey(address), display);
    }
  }

  private updateFormulaMetadata(key: string, ast?: FormulaAst, dependencies: readonly FormulaDependency[] = []): void {
    this.detachNameReferences(key);
    this.detachTableReferences(key);
    if (!ast) {
      this.volatileCells.delete(key);
      this.visibilityDependentCells.delete(key);
      return;
    }
    if (formulaUsesVolatile(ast)) this.volatileCells.add(key);
    else this.volatileCells.delete(key);
    const cell = this.cells.get(key);
    if (!cell) throw new Error(`FORMULA_INDEX_INVARIANT: formula owner ${key} is missing while indexing row-visibility dependencies`);
    if (this.formulaDependsOnRowVisibility(ast, cell.address.sheetId, new Set<string>())) this.visibilityDependentCells.add(key);
    else this.visibilityDependentCells.delete(key);
    const names = [...new Set([
      ...collectNameReferences(ast),
      ...dependencies
        .filter((dependency): dependency is Extract<FormulaDependency, { kind: 'name' }> => dependency.kind === 'name')
        .map((dependency) => dependency.name.trim().toUpperCase()),
    ])];
    this.cellNameRefs.set(key, names);
    for (const name of names) {
      const bucket = this.nameIndex.get(name) ?? new Set<string>();
      bucket.add(key);
      this.nameIndex.set(name, bucket);
    }
    const tables = this.collectDefinedNameTableReferences(ast, cell.address, new Set<string>());
    this.cellTableRefs.set(key, tables);
    for (const table of tables) {
      const bucket = this.tableReferenceIndex.get(table) ?? new Set<string>();
      bucket.add(key);
      this.tableReferenceIndex.set(table, bucket);
    }
  }

  private formulaDependsOnRowVisibility(ast: FormulaAst, ownerSheetId: string, visitedNames: Set<string>): boolean {
    if (formulaUsesRowVisibility(ast)) return true;
    for (const reference of collectNameReferences(ast)) {
      const normalized = reference.trim().toUpperCase();
      const definition = this.definedNameModels.find((entry) => entry.scope === 'sheet'
        && entry.sheetId === ownerSheetId && entry.name.trim().toUpperCase() === normalized)
        ?? this.definedNameModels.find((entry) => entry.scope === 'workbook'
          && entry.name.trim().toUpperCase() === normalized);
      if (!definition) continue;
      const key = `${definition.scope}:${definition.sheetId ?? ''}:${normalized}`;
      if (visitedNames.has(key)) continue;
      visitedNames.add(key);
      const nameAst = parseDefinedNameFormula(definition.formula);
      if (nameAst && this.formulaDependsOnRowVisibility(nameAst, ownerSheetId, visitedNames)) return true;
    }
    return false;
  }

  private detachNameReferences(key: string): void {
    for (const name of this.cellNameRefs.get(key) ?? []) {
      const bucket = this.nameIndex.get(name);
      bucket?.delete(key);
      if (bucket && bucket.size === 0) this.nameIndex.delete(name);
    }
    this.cellNameRefs.delete(key);
  }

  private detachTableReferences(key: string): void {
    for (const table of this.cellTableRefs.get(key) ?? []) {
      const bucket = this.tableReferenceIndex.get(table);
      bucket?.delete(key);
      if (bucket?.size === 0) this.tableReferenceIndex.delete(table);
    }
    this.cellTableRefs.delete(key);
  }

  private collectDefinedNameTableReferences(ast: FormulaAst, owner: CellAddress, visiting: Set<string>): string[] {
    const tables = new Set(collectTableReferences(ast));
    for (const name of collectNameReferences(ast)) {
      const definition = this.findDefinedName(name, owner);
      if (!definition) continue;
      const identity = this.definedNameIdentity(definition);
      if (visiting.has(identity)) continue;
      visiting.add(identity);
      const source = parseDefinedNameFormula(definition.formula);
      if (!source) continue;
      for (const table of this.collectDefinedNameTableReferences(source, owner, visiting)) {
        tables.add(table);
      }
    }
    return [...tables];
  }

  private formulasReferencing(
    index: ReadonlyMap<string, ReadonlySet<string>>,
    references: ReadonlySet<string>,
  ): Map<string, CellAddress> {
    const affected = new Map<string, CellAddress>();
    for (const reference of references) {
      for (const key of index.get(reference) ?? []) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined && cell.ast) affected.set(key, { ...cell.address });
      }
    }
    return affected;
  }

  private refreshFormulaDependencies(owners: ReadonlyMap<string, CellAddress>): void {
    let topologyChanged = false;
    for (const [key, address] of owners) {
      const cell = this.cells.get(key);
      if (!cell?.formula || !cell.ast) continue;
      const dependencies = this.expandNameDependencies(
        collectFormulaDependencies(cell.ast, address, {
          sheetTables: this.sheetTables,
          sheetOrder: this.sheetOrder,
        }),
        address,
        new Set<string>(),
      );
      if (!sameCalculationValue(cell.result.dependencies, dependencies)) topologyChanged = true;
      this.dependencies.set(address, dependencies);
      cell.result = { ...cell.result, dependencies };
      this.updateFormulaMetadata(key, cell.ast, dependencies);
    }
    if (topologyChanged) this.markFormulaTopologyChanged();
  }

  private resolveDefinedName(
    name: string,
    currentCell: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
  ): FormulaValue | undefined {
    const definition = this.findDefinedName(name, currentCell);
    if (!definition) return undefined;
    const identity = this.definedNameIdentity(definition);
    const visitingKey = `name:${identity}`;
    if (visiting.has(visitingKey)) return createFormulaError('#NUM!', `Circular defined name dependency: ${identity}`);
    visiting.add(visitingKey);
    try {
      return resolveDefinedNameSource(definition.formula, {
        currentCell,
        sheetOrder: this.sheetOrder,
        anchor: definition.anchor,
        readCell: (reference) => this.evaluateCell(reference, cache, visiting),
        readRangeMatrix: (range) => this.readRangeMatrix(range, cache, visiting),
        resolveName: (nestedName) => this.resolveDefinedName(nestedName, currentCell, cache, visiting),
        numericContext: this.numericContext,
      });
    } finally {
      visiting.delete(visitingKey);
    }
  }

  private expandNameDependencies(
    dependencies: readonly FormulaDependency[],
    owner: CellAddress,
    visiting: Set<string>,
  ): FormulaDependency[] {
    const expanded = [...dependencies];
    for (const dependency of dependencies) {
      if (dependency.kind !== 'name') continue;
      const definition = this.findDefinedName(dependency.name, owner);
      if (!definition) continue;
      const identity = this.definedNameIdentity(definition);
      if (visiting.has(identity)) continue;
      const source = parseDefinedNameFormula(definition.formula);
      if (!source) continue;
      const projected = definition.anchor
        ? offsetAst(source, owner.row - definition.anchor.row, owner.column - definition.anchor.column)
        : source;
      const nested = collectFormulaDependencies(projected, owner, {
        sheetTables: this.sheetTables,
        sheetOrder: this.sheetOrder,
      });
      expanded.push(...this.expandNameDependencies(nested, owner, new Set([...visiting, identity])));
    }
    return expanded;
  }

  private findDefinedName(name: string, currentCell: CellAddress): FormulaDefinedName | undefined {
    const normalized = name.trim().toUpperCase();
    const sheetKey = currentCell.sheetId.trim().toUpperCase();
    return this.definedNameModels.find((entry) => entry.scope === 'sheet'
      && entry.sheetId?.trim().toUpperCase() === sheetKey
      && entry.name.trim().toUpperCase() === normalized)
      ?? this.definedNameModels.find((entry) => entry.scope === 'workbook'
        && entry.name.trim().toUpperCase() === normalized);
  }

  private definedNameIdentity(definition: FormulaDefinedName): string {
    return definition.scope === 'sheet'
      ? `sheet:${definition.sheetId?.trim().toUpperCase()}:${definition.name.trim().toUpperCase()}`
      : `workbook:${definition.name.trim().toUpperCase()}`;
  }

  private formulaKindAt(address: CellAddress): ReferenceFormulaKind {
    const ast = this.cells.get(cellAddressKey(address))?.ast;
    if (ast?.type !== 'function-call') return 'ordinary';
    const name = ast.name.trim().toUpperCase();
    return name === 'SUBTOTAL' ? 'subtotal' : name === 'AGGREGATE' ? 'aggregate' : 'ordinary';
  }

  private resolveReference(reference: FormulaReferenceNode, currentCell: CellAddress): FormulaEvaluationReference | FormulaError {
    switch (reference.type) {
      case 'whole-column-reference': {
        const sheetId = resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, this.sheetOrder);
        const extent = this.spillEnvironments.get(sheetId);
        if (!extent || extent.rowCount < 1) return createFormulaError('#REF!', `Worksheet extent unavailable for ${sheetId}`);
        return {
          kind: 'reference',
          ranges: [{
            kind: 'range',
            start: { sheetId, row: 0, column: reference.startColumn },
            end: { sheetId, row: extent.rowCount - 1, column: reference.endColumn },
          }],
        };
      }
      case 'whole-row-reference': {
        const sheetId = resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, this.sheetOrder);
        const extent = this.spillEnvironments.get(sheetId);
        if (!extent || extent.columnCount < 1) return createFormulaError('#REF!', `Worksheet extent unavailable for ${sheetId}`);
        return {
          kind: 'reference',
          ranges: [{
            kind: 'range',
            start: { sheetId, row: reference.startRow, column: 0 },
            end: { sheetId, row: reference.endRow, column: extent.columnCount - 1 },
          }],
        };
      }
      case 'reference-union': {
        const ranges: RangeDependency[] = [];
        for (const item of reference.references) {
          const resolved = this.resolveReference(item, currentCell);
          if (isFormulaError(resolved)) return resolved;
          ranges.push(...resolved.ranges);
        }
        return { kind: 'reference', ranges };
      }
      case 'reference-intersection': {
        const left = this.resolveReference(reference.left, currentCell);
        const right = this.resolveReference(reference.right, currentCell);
        if (isFormulaError(left)) return left;
        if (isFormulaError(right)) return right;
        const ranges: RangeDependency[] = [];
        for (const leftRange of left.ranges) {
          for (const rightRange of right.ranges) {
            if (leftRange.start.sheetId !== rightRange.start.sheetId) continue;
            const startRow = Math.max(leftRange.start.row, rightRange.start.row);
            const endRow = Math.min(leftRange.end.row, rightRange.end.row);
            const startColumn = Math.max(leftRange.start.column, rightRange.start.column);
            const endColumn = Math.min(leftRange.end.column, rightRange.end.column);
            if (startRow <= endRow && startColumn <= endColumn) {
              ranges.push({
                kind: 'range',
                start: { sheetId: leftRange.start.sheetId, row: startRow, column: startColumn },
                end: { sheetId: leftRange.start.sheetId, row: endRow, column: endColumn },
              });
            }
          }
        }
        return ranges.length > 0 ? { kind: 'reference', ranges } : createFormulaError('#NULL!', 'Reference intersection is empty');
      }
      case 'sheet-range-reference':
        return createFormulaError('#REF!', '3-D reference requires an ordered worksheet resolver');
      case 'external-reference':
        return createFormulaError('#REF!', `External workbook is unavailable: ${reference.qualifier.workbookId}`);
      default:
        return createFormulaError('#REF!', `Unsupported structured reference: ${reference.type}`);
    }
  }

  private readRangeMatrix(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): ArrayValue {
    const matrix: ArrayValue = [];
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      const line: FormulaValue[] = [];
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        line.push(this.evaluateCellOrSpill({ sheetId: range.start.sheetId, row, column }, cache, visiting, overrides));
      }
      matrix.push(line);
    }
    return matrix;
  }

  private readRange(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): readonly FormulaValue[] {
    const values: FormulaValue[] = [];
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        values.push(this.evaluateCellOrSpill({ sheetId: range.start.sheetId, row, column }, cache, visiting, overrides));
      }
    }
    return values;
  }

  private evaluateCellOrSpill(
    address: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): FormulaValue {
    const value = this.evaluateCell(address, cache, visiting, overrides);
    return this.getSpillValueAt(address.sheetId, address.row, address.column) ?? value;
  }
}

function formulaErrorFrom(error: unknown): FormulaError {
  if (error instanceof FormulaReferenceError) return createFormulaError('#REF!', error.message);
  if (error instanceof FormulaLexError || error instanceof FormulaSyntaxError) {
    return createFormulaError('#PARSE!', error.message, error.position);
  }
  return createFormulaError('#PARSE!', error instanceof Error ? error.message : 'Unable to parse formula');
}

function normalizeFormulaSheetOrder(
  sheetOrder: readonly FormulaSheetIdentity[] | undefined,
  defaultSheetId: string,
): readonly FormulaSheetIdentity[] {
  const source = sheetOrder ?? [{ id: defaultSheetId, name: defaultSheetId }];
  if (source.length === 0) throw new Error('FormulaEngine requires at least one worksheet identity');
  const ids = new Set<string>();
  const names = new Set<string>();
  const normalized = source.map((sheet) => {
    if (!sheet.id.trim() || !sheet.name.trim()) throw new Error('FormulaEngine worksheet identities cannot be empty');
    const normalizedName = sheet.name.toLowerCase();
    if (ids.has(sheet.id) || names.has(normalizedName)) throw new Error('FormulaEngine worksheet identities must be unique');
    ids.add(sheet.id);
    names.add(normalizedName);
    return { id: sheet.id, name: sheet.name };
  });
  if (!ids.has(defaultSheetId)) throw new Error('FormulaEngine default worksheet is missing from worksheet order');
  return normalized;
}

function isAutomaticCalculationMode(mode: RecalculationMode): boolean {
  return mode === 'automatic';
}

function formulaValueDelta(previous: FormulaValue | undefined, current: FormulaValue | undefined): number {
  if (typeof previous === 'number' && typeof current === 'number') return Math.abs(current - previous);
  return JSON.stringify(previous) === JSON.stringify(current) ? 0 : Number.POSITIVE_INFINITY;
}

function isOccupiedInput(cell: StoredCell): boolean {
  return cell.formula !== undefined || (cell.result.value !== null && cell.result.value !== '');
}

function copyDependency(dependency: FormulaDependency): FormulaDependency {
  return dependency.kind === 'cell'
    ? { kind: 'cell', address: { ...dependency.address } }
    : dependency.kind === 'range'
      ? { kind: 'range', start: { ...dependency.start }, end: { ...dependency.end } }
      : dependency.kind === 'reference'
        ? { kind: 'reference', reference: structuredClone(dependency.reference) }
        : { kind: 'name', name: dependency.name };
}

function copySpill(spill: ResolvedSpill): ResolvedSpill {
  return {
    sheetId: spill.sheetId,
    anchor: { ...spill.anchor },
    range: { ...spill.range },
    values: spill.values.map((row) => row.map((value) => {
      if (typeof value === 'object' && value !== null && 'kind' in value) {
        return { kind: 'error' as const, code: value.code, message: value.message };
      }
      return value;
    })),
    state: spill.state,
    ...(spill.blocker === undefined ? {} : { blocker: { ...spill.blocker } }),
  };
}
