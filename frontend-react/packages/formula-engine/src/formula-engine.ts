import { assertExternalCalculationLink, externalReferenceRange, externalSheetKey, type ExternalCalculationLink } from './external-links';
import { createFormulaInputFault } from './input-fault';
import type { CellAddress, FormulaAst, FormulaReferenceNode } from './ast';
import { cellAddressKey, compareCellAddresses, parseCellAddress } from './address';
import { collectFormulaDependencies, collectFormulaReferenceNodes, resolveRangeReference } from './dependencies';
import {
  evaluateFormula,
  evaluateFormulaOperand,
  evaluateFormulaWithTrace,
  type FormulaCellOverride,
  type FormulaEvaluationContext,
  type FormulaEvaluationValue,
  type FormulaEvaluationReference,
  type FormulaEvaluationTrace,
} from './evaluator';
import { formatFormula } from './ast-format';
import { offsetAst } from './ast-rewrite';
import { FormulaLexError, FormulaReferenceError, FormulaSyntaxError } from './errors';
import { parseFormula as parseFormulaSource } from './parser';
import {
  RangeIndex,
  type FormulaDependency,
  type FormulaSheetIdentity,
  type RangeDependency,
} from './range-index';
import type { DefinedNameReferenceIndexUpdate, DefinedNameReferenceOwnerIdentity } from './reference-index';
import { resolveFormulaSheetId } from './sheet-reference';
import { createFormulaError, isArrayValue, isFormulaError, type ArrayValue, type FormulaError, type FormulaValue, type ScalarValue } from './values';
import { normalizeDefinedNameModels, normalizeDefinedNames, parseDefinedNameFormula, type FormulaDefinedName } from './defined-names';
import { collectNameReferences, collectTableReferences, formulaUsesRowVisibility, formulaUsesVolatile } from './formula-analysis';
import { normalizeSheetTables, resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import { canonicalExcelDateFromUtcDate, type CanonicalExcelDateParts, type ExcelDateSystem } from './excel-date';
import { DEFAULT_EXCEL_NUMERIC_CONTEXT, normalizeExcelNumericContext, type ExcelNumericContext } from './numeric';
import { createCalculationEntropyContext, formulaRandom, type CalculationEntropyContext } from './random';
import { DEFAULT_WORKBOOK_COLLATION, normalizeWorkbookCollation, type WorkbookCollationContext } from './collation';
import { analyzeFormulaGraph, type CircularComponent, type FormulaGraphAnalysis } from './circular';
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
  CALCULATION_TASK_VERSION,
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
  type SpillBlockerRange,
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
  /** Static non-cell geometry that blocks a spill, such as merged/table ranges. */
  getBlockedRanges?: () => readonly SpillBlockerRange[];
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

export interface FormulaDefinedNameModelDelta {
  readonly owner: DefinedNameReferenceOwnerIdentity;
  readonly before: FormulaDefinedName;
  readonly after: FormulaDefinedName;
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

function definedNameReferenceOwner(definition: FormulaDefinedName): DefinedNameReferenceIndexUpdate['owner'] {
  return {
    scope: definition.scope,
    name: definition.name,
    ...(definition.sheetId ? { sheetId: definition.sheetId } : {}),
  };
}

function isDefinedNameScope(scope: unknown): scope is FormulaDefinedName['scope'] {
  return scope === 'workbook' || scope === 'sheet';
}

function definedNameReferenceIndexUpdate(definition: FormulaDefinedName): DefinedNameReferenceIndexUpdate {
  const owner = definedNameReferenceOwner(definition);
  const context = definition.anchor ?? (definition.scope === 'sheet'
    ? { sheetId: definition.sheetId!, row: 0, column: 0 }
    : undefined);
  try {
    const source = definition.formula.trim();
    const ast = parseFormulaSource(source.startsWith('=') ? source : `=${source}`);
    return {
      owner,
      references: collectFormulaReferenceNodes(ast),
      context,
      anchor: definition.anchor,
    };
  } catch {
    return { owner, references: [], context, anchor: definition.anchor, failure: 'invalid-formula' };
  }
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
  readonly blockedRanges?: readonly { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number }[];
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
  evaluated?: boolean;
  valueDependencies?: readonly FormulaDependency[];
}

export class FormulaEngine {
  readonly defaultSheetId: string;
  readonly dependencies: RangeIndex;
  private sheetOrder: readonly FormulaSheetIdentity[];
  /** Canonical scoped names. Workbook-only lookup is derived on demand. */
  private definedNamesByIdentity = new Map<string, FormulaDefinedName>();
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
  private cachedCircularComponents: { generation: number; byCell: ReadonlyMap<string, CircularComponent>; graph: FormulaGraphAnalysis; order: ReadonlyMap<string, number> } | null = null;
  private activeEvaluationOwners: ReadonlySet<string> | null = null;
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
  private pendingSpillRecalculationOwners: Set<string> | null = null;
  private spillDependentRoots: Map<string, CellAddress> | null = null;
  private pendingSpillChangeBaselines = new Map<string, { address: CellAddress; spill: ResolvedSpill }>();

  private readonly cells = new Map<string, StoredCell>();
  private readonly inputRowsBySheet = new Map<string, Map<number, Set<number>>>();
  private readonly sortedInputRowsBySheet = new Map<string, readonly number[]>();
  private externalLinks = new Map<string, ExternalCalculationLink>();
  private externalCells = new Map<string, Map<string, { address: CellAddress; value: FormulaValue }>>();
  private externalOwners = new Map<string, Set<string>>();
  private cellExternalTokens = new Map<string, Set<string>>();
  private pendingExternalLinks = new Map<string, { sequence: number; link: ExternalCalculationLink }>();

  getExternalCalculationLinks(): readonly ExternalCalculationLink[] { return [...this.externalLinks.values()].map(link => structuredClone(link)); }

  /** Data comes from a host-authorized read; evaluation never starts network requests. */
  applyExternalCalculationLinks(links: readonly ExternalCalculationLink[], trackForWorker = true): void {
    for (const link of links) assertExternalCalculationLink(link);
    for (const link of links) {
      const previous = this.externalLinks.get(link.token.toUpperCase());
      if (previous && previous.id !== link.id) throw new Error('EXTERNAL_LINK_TOKEN_CONFLICT');
      if (previous && previous.sourceUnitId === link.sourceUnitId && previous.subject === link.subject && (link.accessRevision < previous.accessRevision
        || link.accessRevision === previous.accessRevision && link.sourceRevision < previous.sourceRevision)) continue;
      if (sameCalculationValue(previous, link)) continue;
      for (const sheet of previous?.sheets ?? []) this.externalCells.delete(externalSheetKey(link.id, sheet.id));
      this.externalLinks.set(link.token.toUpperCase(), structuredClone(link));
      for (const sheet of link.sheets) this.externalCells.set(externalSheetKey(link.id, sheet.id), new Map());
      for (const cell of link.cells) {
        const sheetId = externalSheetKey(link.id, cell.address.sheetId);
        const address = { ...cell.address, sheetId };
        this.externalCells.get(sheetId)!.set(`${address.row}:${address.column}`, { address, value: structuredClone(cell.value) });
      }
      for (const owner of this.externalOwners.get(link.token.toUpperCase()) ?? []) this.pendingRecalculationRoots.add(owner);
      if (trackForWorker) {
        this.inputUpdateSequence += 1;
        this.pendingExternalLinks.set(link.id, { sequence: this.inputUpdateSequence, link: structuredClone(link) });
      }
      this.markCalculationStateChanged();
    }
  }

  private isBlockedRange(range: RangeDependency): boolean {
    return this.blockedRanges.some(blocked => blocked.sheetId === range.start.sheetId && blocked.startRow <= range.end.row && range.start.row <= blocked.endRow && blocked.startColumn <= range.end.column && range.start.column <= blocked.endColumn);
  }

  private externalRangeValues(range: RangeDependency, sparse: boolean): Iterable<FormulaValue> | undefined {
    const cells = this.externalCells.get(range.start.sheetId);
    if (!cells) return undefined;
    if (sparse) return [...cells.values()].filter(cell => cell.address.row >= range.start.row && cell.address.row <= range.end.row
      && cell.address.column >= range.start.column && cell.address.column <= range.end.column).map(cell => cell.value);
    return (function* () {
      for (let row = range.start.row; row <= range.end.row; row++) for (let column = range.start.column; column <= range.end.column; column++) yield cells.get(`${row}:${column}`)?.value ?? null;
    })();
  }

  private externalRangeMatrix(range: RangeDependency): ArrayValue | undefined {
    const cells = this.externalCells.get(range.start.sheetId);
    if (!cells) return undefined;
    if ((range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1) > 100000) return [[createFormulaError('#NUM!', 'External matrix exceeds calculation limit')]];
    return Array.from({ length: range.end.row - range.start.row + 1 }, (_, index) => Array.from({ length: range.end.column - range.start.column + 1 }, (_, column) => cells.get(`${range.start.row + index}:${range.start.column + column}`)?.value ?? null));
  }

  private readonly blockedRanges: NonNullable<FormulaEngineOptions['blockedRanges']>;
  private recordFormulaOwners = new Map<string, { tableId: string; recordId: string; fieldId: string; address: CellAddress }>();

  private recordFieldOwners = new Map<string, CellAddress>();

  setRecordFormulaOwners(owners: readonly { tableId: string; recordId: string; fieldId: string; address: CellAddress }[]): void {
    const next = new Map(owners.map(owner => [cellAddressKey(owner.address), structuredClone(owner)]));
    const fields = new Map(owners.map(owner => [JSON.stringify([owner.tableId, owner.recordId, owner.fieldId]), owner.address]));
    if (fields.size !== owners.length || next.size !== owners.length) throw new Error('RECORD_CALCULATION_OWNER_CONFLICT');
    if (JSON.stringify([...next]) === JSON.stringify([...this.recordFormulaOwners])) return;
    this.recordFormulaOwners = next;
    this.recordFieldOwners = fields;
    this.calculationContextGeneration += 1;
  }
  getRecordFormulaOwners(): readonly { tableId: string; recordId: string; fieldId: string; address: CellAddress }[] { return [...this.recordFormulaOwners.values()]; }
  getRecordFormulaOwnerAt(address: CellAddress): Readonly<{ tableId: string; recordId: string; fieldId: string; address: CellAddress }> | undefined { return this.recordFormulaOwners.get(cellAddressKey(address)); }
  getRecordFieldResult(tableId: string, recordId: string, fieldId: string): FormulaResult | undefined {
    const address = this.recordFieldOwners.get(JSON.stringify([tableId, recordId, fieldId]));
    return address ? this.getCellResult(address) : undefined;
  }

  private formulaCount = 0;

  constructor(options: FormulaEngineOptions = {}) {
    this.blockedRanges = structuredClone(options.blockedRanges ?? []);
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
      blockedRanges: snapshot.blockedRanges,
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
    engine.setDefinedNameModels(snapshot.definedNameModels, false);
    engine.setRecordFormulaOwners(snapshot.recordFormulaOwners ?? []);
    engine.applyExternalCalculationLinks(snapshot.externalLinks ?? [], false);
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
        getBlockedRanges: () => spillSpace.blockedRanges.map((range) => ({ ...range })),
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
    for (const spillSpace of snapshot.spillSpaces) {
      for (const spill of spillSpace.spills) {
        engine.spills.set(spillKey({ sheetId: spill.sheetId, row: spill.anchor.row, column: spill.anchor.column }), copySpill(spill));
      }
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
    for (const token of this.cellExternalTokens.get(key) ?? []) this.externalOwners.get(token)?.delete(key);
    this.cellExternalTokens.delete(key);
    this.volatileCells.delete(key);
    this.cells.delete(key);
    this.unindexInputAddress(address);
    if (previous?.formula !== undefined) this.formulaCount = Math.max(0, this.formulaCount - 1);
    if (previous?.formula !== undefined) this.markFormulaTopologyChanged();
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

  updateSheetNames(sheetOrder: readonly FormulaSheetIdentity[]): void {
    const next = normalizeFormulaSheetOrder(sheetOrder, this.defaultSheetId);
    if (next.length !== this.sheetOrder.length
      || next.some((sheet, index) => sheet.id !== this.sheetOrder[index]?.id)) {
      throw new Error('FORMULA_SHEET_IDENTITY_ORDER_MISMATCH: sheet-name updates cannot change worksheet identity order');
    }
    this.applySheetOrder(next, false);
  }

  /** Update worksheet order without rebuilding the calculation context. */
  updateSheetOrder(sheetOrder: readonly FormulaSheetIdentity[]): void {
    const next = normalizeFormulaSheetOrder(sheetOrder, this.defaultSheetId);
    const currentIds = new Set(this.sheetOrder.map((sheet) => sheet.id));
    if (next.length !== currentIds.size || next.some((sheet) => !currentIds.has(sheet.id))) {
      throw new Error('FORMULA_SHEET_IDENTITY_SET_MISMATCH: order updates cannot add or remove worksheet identities');
    }
    this.applySheetOrder(next, true);
  }

  private applySheetOrder(next: readonly FormulaSheetIdentity[], acceptReorder: boolean): void {
    const namesChanged = next.some((sheet, index) => sheet.name !== this.sheetOrder[index]?.name);
    const orderChanged = next.some((sheet, index) => sheet.id !== this.sheetOrder[index]?.id);
    if (!namesChanged && !orderChanged) return;
    this.sheetOrder = next;
    this.dependencies.setSheetOrder(next);
    this.dependencies.updateDefinedNameReferences(this.getDefinedNameModels().map(definedNameReferenceIndexUpdate));
    if (orderChanged && acceptReorder) {
      const roots = this.dependencies.getSheetRangeFormulaOwners();
      const affectedNameTokens = new Set<string>();
      for (const definition of this.getDefinedNameModels()) {
        let ast: FormulaAst;
        try {
          ast = parseFormulaSource(definition.formula.trim().startsWith('=') ? definition.formula : `=${definition.formula}`);
        } catch (error) {
          if (error instanceof FormulaLexError || error instanceof FormulaSyntaxError) continue;
          throw error;
        }
        if (collectFormulaReferenceNodes(ast).some((reference) => reference.type === 'sheet-range-reference')) {
          affectedNameTokens.add(definition.name.trim().toUpperCase());
        }
      }
      const namedRoots = this.formulasReferencing(this.nameIndex, affectedNameTokens);
      for (const [key, address] of namedRoots) this.pendingRecalculationRoots.add(key);
      for (const address of roots) this.pendingRecalculationRoots.add(cellAddressKey(address));
      if (roots.length > 0 || namedRoots.size > 0) this.markFormulaTopologyChanged();
    }
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
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
    if (!entropy) return createFormulaInputFault('#BLOCKED!', 'Volatile formula requires a calculation entropy context', 'runtime-unavailable', 'calculation-entropy');
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

  getFormulaTopologyRevision(): number {
    return this.formulaTopologyGeneration;
  }

  getCalculationInputRevision(): number {
    return this.inputUpdateSequence;
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
  exportPendingCalculationInputs(): { inputs: readonly CalculationInputUpdate[]; externalLinks: readonly ExternalCalculationLink[]; inputRevision: number } {
    const entries = [...this.pendingInputUpdates.values()]
      .sort((left, right) => left.sequence - right.sequence)
      .map((entry) => structuredClone(entry.update));
    return { inputs: entries, externalLinks: [...this.pendingExternalLinks.values()].sort((a, b) => a.sequence - b.sequence).map(entry => structuredClone(entry.link)), inputRevision: this.inputUpdateSequence };
  }

  acknowledgeCalculationInputUpdates(inputRevision: number): void {
    if (!Number.isSafeInteger(inputRevision) || inputRevision < 0) return;
    for (const [key, entry] of this.pendingExternalLinks) if (entry.sequence <= inputRevision) this.pendingExternalLinks.delete(key);
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

  /** Queue existing spill anchors whose candidate ranges intersect changed blocker geometry. */
  notifySpillBlockersChanged(sheetId: string, ranges: readonly SpillBlockerRange[]): readonly CellAddress[] {
    if (ranges.length === 0) return [];
    const affected = new Map<string, CellAddress>();
    for (const spill of this.spills.values()) {
      if (spill.sheetId !== sheetId || !ranges.some((range) =>
        spill.range.startRow <= range.endRow && range.startRow <= spill.range.endRow
        && spill.range.startColumn <= range.endColumn && range.startColumn <= spill.range.endColumn)) continue;
      const address = { sheetId, row: spill.anchor.row, column: spill.anchor.column };
      const key = cellAddressKey(address);
      if (this.cells.get(key)?.formula === undefined) continue;
      affected.set(key, address);
    }
    if (affected.size === 0) return [];
    for (const key of affected.keys()) this.pendingRecalculationRoots.add(key);
    this.calculationContextGeneration += 1;
    this.markCalculationStateChanged();
    return [...affected.values()].sort(compareCellAddresses);
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
              externalLinks: pending.externalLinks,
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
      version: CALCULATION_TASK_VERSION,
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
      if (cell?.formula !== undefined) for (const [dependentKey, address] of this.collectAffected(cell.address)) affected.set(dependentKey, address);
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
        if (cell?.formula !== undefined) for (const [dependentKey, address] of this.collectAffected(cell.address)) affected.set(dependentKey, address);
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
          blockedRanges: (environment.getBlockedRanges?.() ?? []).map((range) => ({ ...range })),
          spills: this.getSpillsForSheet(sheetId)
            .sort((left, right) => compareCellAddresses(
              { sheetId: left.sheetId, row: left.anchor.row, column: left.anchor.column },
              { sheetId: right.sheetId, row: right.anchor.row, column: right.anchor.column },
            ))
            .map(copySpill),
        };
      })
      .sort((left, right) => left.sheetId.localeCompare(right.sheetId));

    return {
      defaultSheetId: this.defaultSheetId,
      blockedRanges: this.blockedRanges,
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
      externalLinks: this.getExternalCalculationLinks(),
      recordFormulaOwners: this.getRecordFormulaOwners(),
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
      cell.evaluated = true;
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
      this.getDefinedNameModels(),
      nextNames,
      (entry) => this.definedNameIdentity(entry),
    );
    if (changedNames.size === 0) return { recalculated: [], results: new Map() };
    const previousByIdentity = new Map(this.definedNamesByIdentity);
    const nextByIdentity = new Map(nextNames.map((entry) => [this.definedNameIdentity(entry), entry] as const));
    const referenceUpdates: DefinedNameReferenceIndexUpdate[] = [];
    for (const [identity, previous] of previousByIdentity) {
      const next = nextByIdentity.get(identity);
      if (JSON.stringify(previous) === JSON.stringify(next)) continue;
      if (!next) {
        referenceUpdates.push({ owner: definedNameReferenceOwner(previous), remove: true });
      }
    }
    for (const [identity, next] of nextByIdentity) {
      const previous = previousByIdentity.get(identity);
      if (JSON.stringify(previous) === JSON.stringify(next)) continue;
      referenceUpdates.push(definedNameReferenceIndexUpdate(next));
    }
    this.dependencies.updateDefinedNameReferences(referenceUpdates);
    this.definedNamesByIdentity = new Map(nextNames.map((entry) => [this.definedNameIdentity(entry), entry] as const));
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

  applyDefinedNameModelDeltas(
    deltas: readonly FormulaDefinedNameModelDelta[],
    recalculate = true,
  ): RecalculationReport {
    if (deltas.length === 0) return { recalculated: [], results: new Map() };
    const normalized = deltas.map((delta) => ({
      owner: { ...delta.owner },
      before: normalizeDefinedNameModels([delta.before])[0]!,
      after: normalizeDefinedNameModels([delta.after])[0]!,
    }));
    const identities = new Set<string>();
    const referenceUpdates: DefinedNameReferenceIndexUpdate[] = [];
    const changedNames = new Set<string>();
    const pending: FormulaDefinedNameModelDelta[] = [];
    for (const delta of normalized) {
      const identity = this.definedNameIdentity(delta.before);
      const declaredIdentity = this.definedNameIdentity({
        ...delta.before,
        name: delta.owner.name,
        scope: delta.owner.scope,
        ...(delta.owner.sheetId ? { sheetId: delta.owner.sheetId } : { sheetId: undefined }),
      });
      if ((delta.owner.scope === 'workbook' && delta.owner.sheetId !== undefined)
        || !isDefinedNameScope(delta.owner.scope)
        || (delta.owner.scope === 'sheet' && !delta.owner.sheetId)
        || identity !== this.definedNameIdentity(delta.after)
        || identity !== declaredIdentity
        || identities.has(identity)) {
        throw new Error('STRUCTURAL_DEFINED_NAME_DELTA_MISMATCH: defined-name owner identity is not unique and stable');
      }
      identities.add(identity);
      const current = this.definedNamesByIdentity.get(identity);
      if (!current) {
        throw new Error(`STRUCTURAL_DEFINED_NAME_DELTA_MISMATCH: canonical owner changed before ${identity}`);
      }
      const currentJson = JSON.stringify(current);
      const beforeJson = JSON.stringify(delta.before);
      const afterJson = JSON.stringify(delta.after);
      if (currentJson === afterJson || beforeJson === afterJson) continue;
      if (currentJson !== beforeJson) {
        throw new Error(`STRUCTURAL_DEFINED_NAME_DELTA_MISMATCH: canonical owner changed before ${identity}`);
      }
      pending.push(delta);
      referenceUpdates.push(definedNameReferenceIndexUpdate(delta.after));
      changedNames.add(delta.after.name.trim().toUpperCase());
    }
    if (changedNames.size === 0) return { recalculated: [], results: new Map() };
    this.dependencies.updateDefinedNameReferences(referenceUpdates);
    for (const delta of pending) {
      const identity = this.definedNameIdentity(delta.after);
      this.definedNamesByIdentity.set(identity, delta.after);
    }
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
    for (const entry of this.definedNamesByIdentity.values()) {
      if (entry.scope === 'workbook') result[entry.name.toUpperCase()] = entry.formula;
    }
    return result;
  }

  getDefinedNameModels(): FormulaDefinedName[] {
    return structuredClone([...this.definedNamesByIdentity.values()]);
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
    this.dependencies.updateDefinedNameReferences(this.getDefinedNameModels().map(definedNameReferenceIndexUpdate));
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
          if (cell?.formula !== undefined) for (const [dependentKey, address] of this.collectAffected(cell.address)) affected.set(dependentKey, address);
        }
        return this.recalculateAffected(affected);
      }

      const affected = this.recalculationMode !== 'automatic' && this.pendingRecalculationRoots.size > 0
        ? this.collectAffectedFromRoots(this.pendingRecalculationRoots)
        : this.allFormulaAddresses();
      for (const key of this.volatileCells) {
        const cell = this.cells.get(key);
        if (cell?.formula !== undefined) for (const [dependentKey, address] of this.collectAffected(cell.address)) affected.set(dependentKey, address);
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
      const [sheetId, row, column] = JSON.parse(key) as [string, number, number];
      const address = cell?.address ?? { sheetId, row, column };
      if (affected.has(key)) continue;
      affected.set(key, { ...address });
      queue.push({ ...address });
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
    for (const token of this.cellExternalTokens.get(key) ?? []) this.externalOwners.get(token)?.delete(key);
    this.cellExternalTokens.delete(key);
      this.volatileCells.delete(key);
      const previous = this.cells.get(key);
      this.cells.delete(key);
      this.unindexInputAddress(address);
      this.visibilityDependentCells.delete(key);
      if (previous?.formula !== undefined) this.formulaCount = Math.max(0, this.formulaCount - 1);
      if (previous?.formula !== undefined) this.markFormulaTopologyChanged();
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
    for (const token of this.cellExternalTokens.get(key) ?? []) this.externalOwners.get(token)?.delete(key);
    this.cellExternalTokens.delete(key);
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
    let valueDependencies: readonly FormulaDependency[] = [];
    let parseError: FormulaError | undefined;

    try {
      const parsed = this.parseFormula(formula);
      const extractedDependencies = collectFormulaDependencies(parsed, address, {
        sheetTables: this.sheetTables,
        sheetOrder: this.sheetOrder,
        resolveNameAst: (name, owner) => this.resolveDefinedNameAst(name, owner),
      });
      const expandedDependencies = this.expandNameDependencies(extractedDependencies, address, new Set<string>());
      this.dependencies.set(address, expandedDependencies);
      ast = parsed;
      formulaDependencies = expandedDependencies;
      valueDependencies = this.collectValueDependencies(parsed, address);
    } catch (error) {
      parseError = formulaErrorFrom(error);
      this.dependencies.set(address, [], true);
    }

    const result: FormulaResult = parseError
      ? { value: parseError, formula, dependencies: [] }
      : { value: null, formula, ast, dependencies: formulaDependencies };
    this.cells.set(key, { address: { ...address }, formula, ast, parseError, result, valueDependencies });
    this.indexInputAddress(address);
    this.updateFormulaMetadata(key, ast, formulaDependencies);
    if (previous?.formula === undefined) this.formulaCount += 1;
    if (previous?.formula === undefined
      || !sameCalculationValue(previous.result.dependencies, formulaDependencies)
      || !sameCalculationValue(previous.valueDependencies, valueDependencies)) this.markFormulaTopologyChanged();
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
    const previousEvaluationOwners = this.activeEvaluationOwners;
    const previousCollector = this.activeResultChangeCollector;
    const previousResultBaselines = this.activeResultChangeBaselines;
    const previousSpillCollector = this.activeSpillChangeCollector;
    const previousSpillBaselines = this.activeSpillChangeBaselines;
    const previousPendingSpillOwners = this.pendingSpillRecalculationOwners;
    const previousSpillDependentRoots = this.spillDependentRoots;
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
    this.pendingSpillRecalculationOwners = new Set(
      [...affected.values()]
        .filter((address) => this.cells.get(cellAddressKey(address))?.formula !== undefined)
        .map(spillKey),
    );
    this.spillDependentRoots = new Map();
    try {
      const report = this.recalculateAffectedCore(affected, resultChanges, spillChanges);
      const recalculated = new Map(report.recalculated.map((address) => [cellAddressKey(address), address]));
      const results = new Map(report.results);
      let passes = 0;
      while (this.spillDependentRoots.size > 0) {
        if (++passes > this.formulaCount + 1) throw new Error('CALCULATION_SPILL_DID_NOT_CONVERGE: spill dependency graph failed to stabilize');
        const roots = [...this.spillDependentRoots.values()];
        this.spillDependentRoots.clear();
        const next = new Map<string, CellAddress>();
        for (const root of roots) for (const [key, address] of this.collectAffected(root)) next.set(key, address);
        const dependentReport = this.recalculateAffectedCore(next, resultChanges, spillChanges);
        for (const address of dependentReport.recalculated) recalculated.set(cellAddressKey(address), address);
        for (const [key, result] of dependentReport.results) results.set(key, result);
      }
      return { recalculated: [...recalculated.values()], results, changedAddresses: [...new Map([...resultChanges, ...spillChanges]).values()].sort(compareCellAddresses) };
    } finally {
      this.activeEvaluationOwners = previousEvaluationOwners;
      this.activeResultChangeCollector = previousCollector;
      this.activeResultChangeBaselines = previousResultBaselines;
      this.activeSpillChangeCollector = previousSpillCollector;
      this.activeSpillChangeBaselines = previousSpillBaselines;
      this.pendingSpillRecalculationOwners = previousPendingSpillOwners;
      this.spillDependentRoots = previousSpillDependentRoots;
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

    const circularByCell = this.getCircularComponentIndex();
    const graph = this.cachedCircularComponents!;
    const dirty = this.collectAffectedFromRoots(this.pendingRecalculationRoots);
    for (const [key, address] of affected) dirty.set(key, address);
    const evaluationOwners = new Map(affected);
    const queue = [...affected.keys()];
    for (let head = 0; head < queue.length; head++) {
      for (const prerequisite of graph.graph.prerequisites.get(queue[head]!) ?? []) {
        const cell = this.cells.get(prerequisite)!;
        if (evaluationOwners.has(prerequisite) || (cell.evaluated && !dirty.has(prerequisite))) continue;
        evaluationOwners.set(prerequisite, cell.address);
        queue.push(prerequisite);
      }
    }
    this.activeEvaluationOwners = new Set(evaluationOwners.keys());
    for (const address of evaluationOwners.values()) this.pendingSpillRecalculationOwners?.add(spillKey(address));
    const selectedCircularComponents = new Map<string, CircularComponent>();
    for (const key of evaluationOwners.keys()) {
      const component = circularByCell.get(key);
      if (component?.cyclic) selectedCircularComponents.set(cellAddressKey(component.members[0]!), component);
    }
    const circularComponents = [...selectedCircularComponents.values()];
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

    const affectedEntries = [...evaluationOwners.entries()]
      .map(([key, address]) => ({ key, address }))
      .sort((left, right) => (graph.order.get(left.key) ?? -1) - (graph.order.get(right.key) ?? -1));
    for (const { key, address } of affectedEntries) {
      const cell = this.cells.get(key);
      if (cell?.formula === undefined) continue;
      if (!handledCircularCells.has(key)) this.evaluateCell(address, evaluationCache, new Set<string>());
      if (affected.has(key)) recalculated.push({ ...address });
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

  private collectValueDependencies(ast: FormulaAst | undefined, address: CellAddress): readonly FormulaDependency[] {
    if (!ast) return [];
    return this.expandNameDependencies(collectFormulaDependencies(ast, address, {
      sheetTables: this.sheetTables, sheetOrder: this.sheetOrder,
      resolveNameAst: (name, owner) => this.resolveDefinedNameAst(name, owner), valueDependencies: true,
    }), address, new Set(), true);
  }

  private getCircularComponentIndex(): ReadonlyMap<string, CircularComponent> {
    if (this.cachedCircularComponents?.generation === this.formulaTopologyGeneration) {
      return this.cachedCircularComponents.byCell;
    }
    const graphNodes = [...this.cells.values()]
      .filter((cell) => cell.formula !== undefined)
      .map((cell) => ({ address: cell.address, dependencies: cell.ast ? this.collectValueDependencies(cell.ast, cell.address) : cell.result.dependencies }));
    const byCell = new Map<string, CircularComponent>();
    const graph = analyzeFormulaGraph(graphNodes, this.sheetOrder);
    for (const component of graph.components) {
      if (!component.cyclic) continue;
      for (const address of component.members) byCell.set(cellAddressKey(address), component);
    }
    this.cachedCircularComponents = { generation: this.formulaTopologyGeneration, byCell, graph, order: new Map(graph.calculationOrder.map((key, index) => [key, index])) };
    return byCell;
  }

  private evaluateCircularComponent(
    members: readonly CellAddress[],
    cache: Map<string, FormulaValue>,
  ): void {
    const orderedMembers = [...members].sort(compareCellAddresses);
    if (!this.calculationSettings.iterativeCalculation) {
      const diagnostic = `Circular reference component: ${orderedMembers.map(cellAddressKey).join(',')}`;
      for (const address of orderedMembers) {
        const key = cellAddressKey(address);
        const cell = this.cells.get(key);
        if (!cell?.formula) continue;
        const value = createFormulaError('#NUM!', diagnostic);
        const previousValue = cell.result.value;
        cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
        cell.evaluated = true;
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
      const delta = orderedMembers.reduce((maximum, address) => Math.max(maximum, formulaValueDelta(previous.get(cellAddressKey(address)), current.get(cellAddressKey(address)))), 0);
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
    if (overrides.length === 0 && cell.evaluated && this.activeEvaluationOwners && !this.activeEvaluationOwners.has(key)) {
      cache.set(key, cell.result.value);
      return cell.result.value;
    }
    if (cell.parseError) {
      const previousValue = cell.result.value;
      cache.set(key, cell.parseError);
      cell.result = { value: cell.parseError, formula: cell.formula, dependencies: cell.result.dependencies };
      cell.evaluated = true;
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
    if (overrides.length > 0) { cache.set(key, value); return value; }
    cell.result = { value, formula: cell.formula, ast: cell.ast, dependencies: cell.result.dependencies };
    cell.evaluated = true;
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
    if (!sameCalculationValue(before, after)) {
      for (const spill of [before, after]) {
        if (!spill || spill.state !== 'ok') continue;
        for (const dependent of this.dependencies.getRangeDependents(spill.sheetId, spill.range)) {
          if (cellAddressKey(dependent) !== cellAddressKey(address)) this.spillDependentRoots?.set(cellAddressKey(dependent), dependent);
        }
      }
    }
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
        calculationReferenceDate: this.activeCalculationEntropy
          ? canonicalExcelDateFromUtcDate(new Date(
            this.activeCalculationEntropy.calculationTimeUtcMs
              - this.activeCalculationEntropy.calculationTimeZoneOffsetMinutes * 60_000,
          ), this.dateSystem)
          : undefined,
        numericContext: this.numericContext,
        collationContext: this.collationContext,
        rowVisibility: this.rowVisibilityResolver,
        readFormulaKind: (reference) => this.formulaKindAt(reference),
        random: (functionName, occurrence, elementIndex) => this.randomForCell(cell.address, functionName, occurrence, elementIndex),
        readCell: (reference) => this.blockedRanges.some(range => range.sheetId === reference.sheetId && range.startRow <= reference.row && reference.row <= range.endRow && range.startColumn <= reference.column && reference.column <= range.endColumn) ? createFormulaInputFault('#BLOCKED!', 'Hidden formula input', 'access-denied', reference.sheetId) : this.externalCells.get(reference.sheetId)?.get(`${reference.row}:${reference.column}`)?.value ?? this.readCellWithOverrides(reference, cache, visiting, overrides),
        readRange: (range) => this.isBlockedRange(range) ? [createFormulaInputFault('#BLOCKED!', 'Hidden formula range', 'access-denied', range.start.sheetId)] : this.externalRangeValues(range, false) ?? this.readRange(range, cache, visiting, overrides),
        readRangeMatrix: (range) => this.isBlockedRange(range) ? [[createFormulaInputFault('#BLOCKED!', 'Hidden formula range', 'access-denied', range.start.sheetId)]] : this.externalRangeMatrix(range) ?? this.readRangeMatrix(range, cache, visiting, overrides),
        readSparseRange: (range) => this.isBlockedRange(range) ? [createFormulaInputFault('#BLOCKED!', 'Hidden formula range', 'access-denied', range.start.sheetId)] : this.externalRangeValues(range, true) ?? this.readSparseRange(range, cache, visiting, overrides),
        readSparseRangeCells: (range) => this.isBlockedRange(range)
          ? [{ address: range.start, value: createFormulaInputFault('#BLOCKED!', 'Hidden formula range', 'access-denied', range.start.sheetId) }]
          : this.externalCells.has(range.start.sheetId)
            ? [...this.externalCells.get(range.start.sheetId)!.values()].filter(cell => cell.address.row >= range.start.row && cell.address.row <= range.end.row && cell.address.column >= range.start.column && cell.address.column <= range.end.column).map(cell => ({ address: { ...cell.address, sheetId: range.start.sheetId }, value: cell.value }))
            : this.readSparseRangeCells(range, cache, visiting, overrides),
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
        resolveName: (name) => this.resolveDefinedName(name, cell.address, cache, visiting, overrides),
        resolveFunction: (name) => {
          const definition = this.findDefinedName(name, cell.address);
          if (!definition) return undefined;
          const source = parseDefinedNameFormula(definition.formula);
          const parsed = source && definition.anchor ? offsetAst(source, cell.address.row - definition.anchor.row, cell.address.column - definition.anchor.column) : source;
          return parsed;
        },
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

  private findCellOverride(address: CellAddress, overrides: readonly FormulaCellOverride[]): FormulaCellOverride | undefined {
    for (let index = overrides.length - 1; index >= 0; index--) {
      const candidate = overrides[index]!;
      if (candidate.address.sheetId === address.sheetId && candidate.address.row === address.row && candidate.address.column === address.column) return candidate;
    }
    return undefined;
  }

  private readCellWithOverrides(
    address: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[],
  ): FormulaValue {
    const override = this.findCellOverride(address, overrides);
    return override ? override.value : this.evaluateCellOrSpill(address, cache, visiting, overrides);
  }

  private refreshSpill(address: CellAddress, value: FormulaValue, cache?: Map<string, FormulaValue>): void {
    const key = spillKey(address);
    this.pendingSpillRecalculationOwners?.delete(key);
    const previous = this.spills.get(key);
    if (this.recordFormulaOwners.has(cellAddressKey(address)) || !isSpillMatrix(value)) {
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
    environment.ensureExtent?.(address.row + value.length, address.column + value.reduce((maximum, row) => Math.max(maximum, row.length), 0));
    const spill = resolveSpill({
      sheetId: address.sheetId,
      anchor: { row: address.row, column: address.column },
      values: value,
      rowCount: environment.rowCount,
      columnCount: environment.columnCount,
      isOccupied: environment.isOccupied,
      blockedRanges: [
        ...(environment.getBlockedRanges?.() ?? []),
        ...[...this.spills.values()]
          .filter((candidate) => candidate.sheetId === address.sheetId
            && candidate.state === 'ok'
            && !this.pendingSpillRecalculationOwners?.has(spillKey({
              sheetId: candidate.sheetId,
              row: candidate.anchor.row,
              column: candidate.anchor.column,
            }))
            && (candidate.anchor.row !== address.row || candidate.anchor.column !== address.column))
          .map((candidate) => ({
            startRow: candidate.range.startRow,
            endRow: candidate.range.endRow,
            startColumn: candidate.range.startColumn,
            endColumn: candidate.range.endColumn,
          })),
      ],
    });
    this.spills.set(key, spill);
    if (!sameCalculationValue(previous, spill)) this.recordSpillProjectionChange(address, previous, spill);
    if (spill.state !== 'ok') {
      const display = anchorDisplayValue(spill, value);
      const cell = this.cells.get(cellAddressKey(address));
      if (cell) cell.result = { ...cell.result, value: display };
      cache?.set(cellAddressKey(address), display);
    }
  }

  private updateFormulaMetadata(key: string, ast?: FormulaAst, dependencies: readonly FormulaDependency[] = []): void {
    this.detachNameReferences(key);
    this.detachTableReferences(key);
    for (const token of this.cellExternalTokens.get(key) ?? []) this.externalOwners.get(token)?.delete(key);
    this.cellExternalTokens.delete(key);
    if (!ast) {
      this.volatileCells.delete(key);
      this.visibilityDependentCells.delete(key);
      return;
    }
    const cell = this.cells.get(key);
    if (!cell) throw new Error(`FORMULA_INDEX_INVARIANT: formula owner ${key} is missing while indexing row-visibility dependencies`);
    if (this.formulaDependsOnNamedProperty(ast, cell.address.sheetId, formulaUsesVolatile, new Set())) this.volatileCells.add(key);
    else this.volatileCells.delete(key);
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
    const tokens = new Set<string>();
    const visitExternal = (source: FormulaAst, visited: Set<string>): void => {
      for (const ref of collectFormulaReferenceNodes(source)) if (ref.type === 'external-reference') tokens.add(ref.qualifier.workbookId.toUpperCase());
      for (const name of collectNameReferences(source)) {
        const definition = this.findDefinedName(name, cell.address);
        if (!definition || visited.has(name.toUpperCase())) continue;
        visited.add(name.toUpperCase());
        try { visitExternal(parseFormulaSource(definition.formula.startsWith('=') ? definition.formula : `=${definition.formula}`), visited); } catch { /* The normal formula/name error is retained. */ }
      }
    };
    visitExternal(ast, new Set());
    this.cellExternalTokens.set(key, tokens);
    for (const token of tokens) { const owners = this.externalOwners.get(token) ?? new Set<string>(); owners.add(key); this.externalOwners.set(token, owners); }
    const tables = this.collectDefinedNameTableReferences(ast, cell.address, new Set<string>());
    this.cellTableRefs.set(key, tables);
    for (const table of tables) {
      const bucket = this.tableReferenceIndex.get(table) ?? new Set<string>();
      bucket.add(key);
      this.tableReferenceIndex.set(table, bucket);
    }
  }

  private formulaDependsOnRowVisibility(ast: FormulaAst, ownerSheetId: string, visitedNames: Set<string>): boolean {
    return this.formulaDependsOnNamedProperty(ast, ownerSheetId, formulaUsesRowVisibility, visitedNames);
  }

  private formulaDependsOnNamedProperty(ast: FormulaAst, ownerSheetId: string, predicate: (ast: FormulaAst) => boolean, visitedNames: Set<string>): boolean {
    if (predicate(ast)) return true;
    for (const reference of collectNameReferences(ast)) {
      const normalized = reference.trim().toUpperCase();
      const definition = this.findDefinedName(reference, { sheetId: ownerSheetId, row: 0, column: 0 });
      if (!definition) continue;
      const key = `${definition.scope}:${definition.sheetId ?? ''}:${normalized}`;
      if (visitedNames.has(key)) continue;
      visitedNames.add(key);
      const nameAst = parseDefinedNameFormula(definition.formula);
      if (nameAst && this.formulaDependsOnNamedProperty(nameAst, ownerSheetId, predicate, visitedNames)) return true;
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
          resolveNameAst: (name, owner) => this.resolveDefinedNameAst(name, owner),
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
    overrides: readonly FormulaCellOverride[],
  ): FormulaEvaluationValue | undefined {
    const definition = this.findDefinedName(name, currentCell);
    if (!definition) return undefined;
    const identity = this.definedNameIdentity(definition);
    const visitingKey = `name:${identity}`;
    if (visiting.has(visitingKey)) return createFormulaError('#NUM!', `Circular defined name dependency: ${identity}`);
    visiting.add(visitingKey);
    try {
      const source = definition.formula.trim();
      const parsed = parseFormulaSource(source.startsWith('=') ? source : `=${source}`);
      const ast = definition.anchor ? offsetAst(parsed, currentCell.row - definition.anchor.row, currentCell.column - definition.anchor.column) : parsed;
      return evaluateFormulaOperand(ast, this.createEvaluationContext({ address: currentCell, result: { value: null, dependencies: [] } }, cache, visiting, overrides));
    } catch (error) {
      return formulaErrorFrom(error);
    } finally {
      visiting.delete(visitingKey);
    }
  }

  private expandNameDependencies(
    dependencies: readonly FormulaDependency[],
    owner: CellAddress,
    visiting: Set<string>,
    valueDependencies = false,
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
        valueDependencies,
        resolveNameAst: (name, owner) => this.resolveDefinedNameAst(name, owner),
      });
      expanded.push(...this.expandNameDependencies(nested, owner, new Set([...visiting, identity]), valueDependencies));
    }
    return expanded;
  }

  private findDefinedName(name: string, currentCell: CellAddress): FormulaDefinedName | undefined {
    const normalized = name.trim().toUpperCase();
    const sheetKey = currentCell.sheetId.trim();
    return this.definedNamesByIdentity.get(`sheet:${sheetKey}:${normalized}`)
      ?? this.definedNamesByIdentity.get(`workbook:${normalized}`);
  }

  private definedNameIdentity(definition: FormulaDefinedName): string {
    return definition.scope === 'sheet'
      ? `sheet:${definition.sheetId?.trim()}:${definition.name.trim().toUpperCase()}`
      : `workbook:${definition.name.trim().toUpperCase()}`;
  }

  private resolveDefinedNameAst(name: string, owner: CellAddress): FormulaAst | undefined {
    const definition = this.findDefinedName(name, owner);
    if (!definition) return undefined;
    const ast = parseDefinedNameFormula(definition.formula);
    return ast && definition.anchor ? offsetAst(ast, owner.row - definition.anchor.row, owner.column - definition.anchor.column) : ast;
  }

  private formulaKindAt(address: CellAddress): ReferenceFormulaKind {
    const ast = this.cells.get(cellAddressKey(address))?.ast;
    if (ast?.type !== 'function-call') return 'ordinary';
    const name = ast.name.trim().toUpperCase();
    return name === 'SUBTOTAL' ? 'subtotal' : name === 'AGGREGATE' ? 'aggregate' : 'ordinary';
  }

  private resolveReference(reference: FormulaReferenceNode, currentCell: CellAddress): FormulaEvaluationReference | FormulaError {
    switch (reference.type) {
      case 'cell-reference': {
        const address = { ...reference.reference, sheetId: resolveFormulaSheetId(reference.reference.sheetId, currentCell.sheetId, this.sheetOrder) };
        return { kind: 'reference', ranges: [{ kind: 'range', start: address, end: address }] };
      }
      case 'range-reference': return { kind: 'reference', ranges: [resolveRangeReference(reference, currentCell, this.sheetOrder)] };
      case 'table-reference': {
        const result = resolveSheetTableReference(reference.tableName, reference, currentCell, this.sheetTables);
        if (isFormulaError(result)) return result;
        return { kind: 'reference', ranges: ['start' in result ? result : { kind: 'range', start: result, end: result }] };
      }

      case 'whole-column-reference': {
        const sheetId = resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, this.sheetOrder);
        return {
          kind: 'reference',
          ranges: [{
            kind: 'range',
            start: { sheetId, row: 0, column: reference.startColumn },
            end: { sheetId, row: 1_048_575, column: reference.endColumn },
          }],
        };
      }
      case 'whole-row-reference': {
        const sheetId = resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, this.sheetOrder);
        return {
          kind: 'reference',
          ranges: [{
            kind: 'range',
            start: { sheetId, row: reference.startRow, column: 0 },
            end: { sheetId, row: reference.endRow, column: 16_383 },
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
        return this.resolveSheetRangeReference(reference, currentCell);
      case 'external-reference': {
        const resolved = externalReferenceRange(reference, this.externalLinks.get(reference.qualifier.workbookId.toUpperCase()));
        return isFormulaError(resolved) ? resolved : { kind: 'reference', ranges: [resolved] };
      }
      default:
        return createFormulaError('#REF!', `Unsupported structured reference: ${reference.type}`);
    }
  }

  private resolveSheetRangeReference(
    reference: Extract<FormulaReferenceNode, { type: 'sheet-range-reference' }>,
    currentCell: CellAddress,
  ): FormulaEvaluationReference | FormulaError {
    const startSheetId = resolveFormulaSheetId(reference.qualifier.startSheetId, currentCell.sheetId, this.sheetOrder);
    const endSheetId = resolveFormulaSheetId(reference.qualifier.endSheetId, currentCell.sheetId, this.sheetOrder);
    const startIndex = this.sheetOrder.findIndex((sheet) => sheet.id === startSheetId);
    const endIndex = this.sheetOrder.findIndex((sheet) => sheet.id === endSheetId);
    if (startIndex < 0 || endIndex < 0) {
      return createFormulaError('#REF!', '3-D reference worksheet boundary is unresolved');
    }

    const ranges: RangeDependency[] = [];
    for (const sheet of this.sheetOrder.slice(Math.min(startIndex, endIndex), Math.max(startIndex, endIndex) + 1)) {
      switch (reference.reference.type) {
        case 'cell-reference': {
          const address = { ...reference.reference.reference, sheetId: sheet.id };
          ranges.push({ kind: 'range', start: address, end: { ...address } });
          break;
        }
        case 'range-reference': {
          const qualified = {
            ...reference.reference,
            start: { ...reference.reference.start, reference: { ...reference.reference.start.reference, sheetId: sheet.id } },
            end: { ...reference.reference.end, reference: { ...reference.reference.end.reference, sheetId: sheet.id } },
          };
          ranges.push(resolveRangeReference(qualified, { ...currentCell, sheetId: sheet.id }, this.sheetOrder));
          break;
        }
        case 'whole-column-reference': {
          ranges.push({ kind: 'range', start: { sheetId: sheet.id, row: 0, column: reference.reference.startColumn }, end: { sheetId: sheet.id, row: 1_048_575, column: reference.reference.endColumn } });
          break;
        }
        case 'whole-row-reference': {
          ranges.push({ kind: 'range', start: { sheetId: sheet.id, row: reference.reference.startRow, column: 0 }, end: { sheetId: sheet.id, row: reference.reference.endRow, column: 16_383 } });
          break;
        }
        case 'invalid-reference':
          return createFormulaError('#REF!', '3-D reference contains a deleted cell');
      }
    }
    return { kind: 'reference', ranges };
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
  ): Iterable<FormulaValue> {
    return (function* (engine: FormulaEngine) {
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        yield engine.evaluateCellOrSpill({ sheetId: range.start.sheetId, row, column }, cache, visiting, overrides);
      }
    }
    })(this);
  }

  private *readSparseRange(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): Iterable<FormulaValue> {
    for (const cell of this.readSparseRangeCells(range, cache, visiting, overrides)) yield cell.value;
  }

  private *readSparseRangeCells(
    range: RangeDependency,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): Iterable<{ readonly address: CellAddress; readonly value: FormulaValue }> {
    const inside = (address: CellAddress) => address.sheetId === range.start.sheetId
      && address.row >= range.start.row && address.row <= range.end.row
      && address.column >= range.start.column && address.column <= range.end.column;
    const seen = new Set<string>();
    const addresses = this.getInputAddressesInRange({
      sheetId: range.start.sheetId, startRow: range.start.row, endRow: range.end.row,
      startColumn: range.start.column, endColumn: range.end.column,
    });
    for (const address of [...addresses, ...overrides.filter(({ address }) => inside(address)).map(({ address }) => address)]) {
      const key = cellAddressKey(address);
      if (seen.has(key)) continue;
      seen.add(key);
      yield { address, value: this.evaluateCellOrSpill(address, cache, visiting, overrides) };
    }
    // Spill cells are projections of their single anchor. They are not inputs.
    for (const spill of this.spills.values()) {
      if (spill.sheetId !== range.start.sheetId || spill.state !== 'ok') continue;
      for (let row = Math.max(spill.range.startRow, range.start.row); row <= Math.min(spill.range.endRow, range.end.row); row += 1) {
        for (let column = Math.max(spill.range.startColumn, range.start.column); column <= Math.min(spill.range.endColumn, range.end.column); column += 1) {
          const address = { sheetId: spill.sheetId, row, column };
          const key = cellAddressKey(address);
          if (seen.has(key)) continue;
          seen.add(key);
          yield { address, value: this.evaluateCellOrSpill(address, cache, visiting, overrides) };
        }
      }
    }
  }

  private evaluateCellOrSpill(
    address: CellAddress,
    cache: Map<string, FormulaValue>,
    visiting: Set<string>,
    overrides: readonly FormulaCellOverride[] = [],
  ): FormulaValue {
    const override = this.findCellOverride(address, overrides);
    if (override) return override.value;
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
