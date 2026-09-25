import type { CellAddress } from './ast';
import { cellAddressKey } from './address';
import { isWorkbookCalculationSettings, type WorkbookCalculationSettings } from './calculation-settings';
import type { SheetTableRef } from './sheet-table-resolver';
import type { ScalarValue } from './values';
import type { ResolvedSpill, SpillBlockerRange } from './spill-resolver';
import type { FormulaDefinedName } from './defined-names';
import type { CanonicalExcelDateParts, ExcelDateSystem } from './excel-date';
import type { ExcelNumericContext } from './numeric';
import type { CalculationEntropyContext } from './random';
import type { WorkbookCollationContext } from './collation';
import { assertFormulaVisibilitySnapshot, type FormulaVisibilitySnapshot } from './reference-cursor';

/**
 * A data-only copy of the formula inputs required by an isolated calculation
 * task. It is transient work, never a second workbook persistence model.
 */
export interface FormulaCalculationSnapshot {
  readonly defaultSheetId: string;
  readonly sheetOrder: readonly { readonly id: string; readonly name: string }[];
  readonly calculationSettings: WorkbookCalculationSettings;
  readonly dateSystem: ExcelDateSystem;
  readonly canonicalReferenceDate?: CanonicalExcelDateParts;
  readonly numericContext: ExcelNumericContext;
  readonly calculationEntropy: CalculationEntropyContext;
  readonly collationContext: WorkbookCollationContext;
  readonly visibility?: FormulaVisibilitySnapshot;
  readonly cells: readonly FormulaCellSnapshot[];
  readonly definedNameModels: readonly FormulaDefinedName[];
  readonly sheetTables: readonly SheetTableRef[];
  readonly spillSpaces: readonly FormulaSpillSpaceSnapshot[];
  readonly pendingRoots: readonly CellAddress[];
}

export interface FormulaCellSnapshot {
  readonly address: CellAddress;
  readonly input: FormulaCellInputSnapshot;
}

export type FormulaCellInputSnapshot =
  | { readonly kind: 'value'; readonly value: ScalarValue }
  | { readonly kind: 'formula'; readonly formula: string };

/**
 * Spill occupancy cannot contain a callback when crossing a Worker boundary.
 * The host materializes the occupied coordinates for the calculation task.
 */
export interface FormulaSpillSpaceSnapshot {
  readonly sheetId: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly occupied: readonly CellAddress[];
  readonly blockedRanges: readonly SpillBlockerRange[];
  readonly spills: readonly ResolvedSpill[];
}

export function assertFormulaCalculationSnapshot(value: unknown): asserts value is FormulaCalculationSnapshot {
  if (!isRecord(value)) throw new Error('Calculation snapshot must be an object');
  if (typeof value.defaultSheetId !== 'string' || value.defaultSheetId.length === 0) {
    throw new Error('Calculation snapshot requires a default worksheet id');
  }
  if (!Array.isArray(value.sheetOrder) || value.sheetOrder.length === 0
    || !value.sheetOrder.every(isSheetIdentity)
    || !value.sheetOrder.some((sheet) => sheet.id === value.defaultSheetId)) {
    throw new Error('Calculation snapshot has an invalid worksheet order');
  }
  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();
  for (const sheet of value.sheetOrder) {
    const normalizedName = sheet.name.toLowerCase();
    if (sheetIds.has(sheet.id) || sheetNames.has(normalizedName)) {
      throw new Error('Calculation snapshot worksheet identities are not unique');
    }
    sheetIds.add(sheet.id);
    sheetNames.add(normalizedName);
  }
  if (!isWorkbookCalculationSettings(value.calculationSettings)) throw new Error('Calculation snapshot has invalid calculation settings');
  if (value.dateSystem !== '1900' && value.dateSystem !== '1904') throw new Error('Calculation snapshot has an invalid date system');
  if (value.canonicalReferenceDate !== undefined && !isCanonicalDateParts(value.canonicalReferenceDate)) {
    throw new Error('Calculation snapshot has an invalid canonical reference date');
  }
  if (!isExcelNumericContext(value.numericContext)) throw new Error('Calculation snapshot has an invalid numeric context');
  if (!isCalculationEntropyContext(value.calculationEntropy)) throw new Error('Calculation snapshot has an invalid calculation entropy');
  if (!isWorkbookCollationContext(value.collationContext)) throw new Error('Calculation snapshot has an invalid collation context');
  if (value.visibility !== undefined) assertFormulaVisibilitySnapshot(value.visibility);
  if (!Array.isArray(value.cells) || !value.cells.every(isFormulaCellSnapshot)) {
    throw new Error('Calculation snapshot has invalid cells');
  }
  if (!Array.isArray(value.definedNameModels) || !value.definedNameModels.every(isFormulaDefinedName)) throw new Error('Calculation snapshot has invalid defined names');
  if (!Array.isArray(value.sheetTables) || !value.sheetTables.every(isSheetTableRef)) {
    throw new Error('Calculation snapshot has invalid sheet tables');
  }
  if (!Array.isArray(value.spillSpaces) || !value.spillSpaces.every(isFormulaSpillSpaceSnapshot)) {
    throw new Error('Calculation snapshot has invalid spill spaces');
  }
  const formulaAddresses = new Set(value.cells
    .filter((cell): cell is FormulaCellSnapshot & { readonly input: { readonly kind: 'formula'; readonly formula: string } } => cell.input.kind === 'formula')
    .map(({ address }) => cellAddressKey(address)));
  const spillSpaceSheetIds = new Set<string>();
  const spillAnchorAddresses = new Set<string>();
  for (const spillSpace of value.spillSpaces) {
    if (!sheetIds.has(spillSpace.sheetId) || spillSpaceSheetIds.has(spillSpace.sheetId)) {
      throw new Error('Calculation snapshot spill spaces have invalid worksheet identities');
    }
    spillSpaceSheetIds.add(spillSpace.sheetId);
    for (const spill of spillSpace.spills) {
      const address = { sheetId: spill.sheetId, row: spill.anchor.row, column: spill.anchor.column };
      const key = cellAddressKey(address);
      if (!formulaAddresses.has(key)) {
        throw new Error('Calculation snapshot spill has no formula anchor');
      }
      if (spillAnchorAddresses.has(key)) throw new Error('Calculation snapshot has duplicate spill anchors');
      spillAnchorAddresses.add(key);
    }
  }
  if (!Array.isArray(value.pendingRoots) || !value.pendingRoots.every(isCellAddress)) {
    throw new Error('Calculation snapshot has invalid dirty roots');
  }
}

function isSheetIdentity(value: unknown): value is { readonly id: string; readonly name: string } {
  return isRecord(value)
    && typeof value.id === 'string' && value.id.trim().length > 0
    && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isFormulaCellSnapshot(value: unknown): value is FormulaCellSnapshot {
  return isRecord(value)
    && isCellAddress(value.address)
    && isFormulaCellInputSnapshot(value.input);
}

function isFormulaCellInputSnapshot(value: unknown): value is FormulaCellInputSnapshot {
  if (!isRecord(value)) return false;
  if (value.kind === 'formula') return typeof value.formula === 'string';
  return value.kind === 'value' && isScalarValue(value.value);
}

function isFormulaSpillSpaceSnapshot(value: unknown): value is FormulaSpillSpaceSnapshot {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || value.sheetId.length === 0) return false;
  const sheetId = value.sheetId;
  return isNonNegativeInteger(value.rowCount)
    && isNonNegativeInteger(value.columnCount)
    && Array.isArray(value.occupied)
    && value.occupied.every((address) => isCellAddress(address) && address.sheetId === sheetId)
    && Array.isArray(value.blockedRanges)
    && value.blockedRanges.every(isSpillBlockerRange)
    && Array.isArray(value.spills)
    && value.spills.every((spill) => isResolvedSpillSnapshot(spill, sheetId));
}

function isSpillBlockerRange(value: unknown): value is SpillBlockerRange {
  return isRecord(value)
    && isNonNegativeInteger(value.startRow)
    && isNonNegativeInteger(value.endRow)
    && value.endRow >= value.startRow
    && isNonNegativeInteger(value.startColumn)
    && isNonNegativeInteger(value.endColumn)
    && value.endColumn >= value.startColumn;
}

function isResolvedSpillSnapshot(value: unknown, sheetId: string): value is ResolvedSpill {
  if (!isRecord(value) || value.sheetId !== sheetId || !isRecord(value.anchor) || !isRecord(value.range)) return false;
  const { anchor, range } = value;
  return isNonNegativeInteger(anchor.row)
    && isNonNegativeInteger(anchor.column)
    && range.sheetId === sheetId
    && isNonNegativeInteger(range.startRow)
    && isNonNegativeInteger(range.endRow)
    && range.endRow >= range.startRow
    && isNonNegativeInteger(range.startColumn)
    && isNonNegativeInteger(range.endColumn)
    && range.endColumn >= range.startColumn
    && range.startRow === anchor.row
    && range.startColumn === anchor.column
    && (value.state === 'ok' || value.state === 'blocked' || value.state === 'spill-error')
    && (value.state === 'blocked'
      ? isRecord(value.blocker) && isNonNegativeInteger(value.blocker.row) && isNonNegativeInteger(value.blocker.column)
      : value.blocker === undefined)
    && Array.isArray(value.values)
    && value.values.every((row) => Array.isArray(row) && row.every(isSpillValue));
}

function isSpillValue(value: unknown): boolean {
  if (isScalarValue(value)) return true;
  return isRecord(value)
    && value.kind === 'error'
    && typeof value.code === 'string'
    && (value.message === undefined || typeof value.message === 'string');
}

function isSheetTableRef(value: unknown): value is SheetTableRef {
  if (!isRecord(value)) return false;
  const range = value.range;
  return typeof value.id === 'string'
    && typeof value.sheetId === 'string'
    && typeof value.name === 'string'
    && isRecord(range)
    && typeof range.sheetId === 'string'
    && isNonNegativeInteger(range.startRow)
    && isNonNegativeInteger(range.endRow)
    && isNonNegativeInteger(range.startColumn)
    && isNonNegativeInteger(range.endColumn)
    && typeof value.hasHeaderRow === 'boolean'
    && typeof value.hasTotalRow === 'boolean'
    && Array.isArray(value.columns)
    && value.columns.every((column) => isRecord(column) && typeof column.id === 'string' && typeof column.name === 'string');
}

function isCellAddress(value: unknown): value is CellAddress {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && value.sheetId.length > 0
    && isNonNegativeInteger(value.row)
    && isNonNegativeInteger(value.column);
}

function isScalarValue(value: unknown): value is ScalarValue {
  return value === null
    || typeof value === 'number'
    || typeof value === 'string'
    || typeof value === 'boolean';
}

function isFormulaDefinedName(value: unknown): value is FormulaDefinedName {
  return isRecord(value)
    && typeof value.name === 'string'
    && value.name.trim().length > 0
    && typeof value.formula === 'string'
    && value.formula.trim().length > 0
    && (value.scope === 'workbook' || value.scope === 'sheet')
    && (value.scope === 'workbook'
      ? value.sheetId === undefined
      : typeof value.sheetId === 'string' && value.sheetId.trim().length > 0)
    && (value.anchor === undefined || isCellAddress(value.anchor));
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isCanonicalDateParts(value: unknown): value is CanonicalExcelDateParts {
  if (!isRecord(value)) return false;
  const integerInRange = (candidate: unknown, minimum: number, maximum: number): candidate is number => typeof candidate === 'number' && Number.isInteger(candidate) && candidate >= minimum && candidate <= maximum;
  return integerInRange(value.year, 1, 9999)
    && integerInRange(value.month, 1, 12)
    && integerInRange(value.day, 1, 31)
    && integerInRange(value.hour, 0, 23)
    && integerInRange(value.minute, 0, 59)
    && integerInRange(value.second, 0, 59)
    && integerInRange(value.millisecond, 0, 999);
}

function isExcelNumericContext(value: unknown): value is ExcelNumericContext {
  if (!isRecord(value)) return false;
  return typeof value.significantDigits === 'number'
    && Number.isSafeInteger(value.significantDigits)
    && value.significantDigits >= 1
    && value.significantDigits <= 15;
}

function isCalculationEntropyContext(value: unknown): value is CalculationEntropyContext {
  if (!isRecord(value)) return false;
  return typeof value.cycleId === 'number'
    && Number.isSafeInteger(value.cycleId)
    && value.cycleId >= 0
    && typeof value.entropySeed === 'string'
    && value.entropySeed.trim().length > 0
    && typeof value.passIndex === 'number'
    && Number.isSafeInteger(value.passIndex)
    && value.passIndex >= 0;
}

function isWorkbookCollationContext(value: unknown): value is WorkbookCollationContext {
  if (!isRecord(value)) return false;
  return typeof value.cultureId === 'string'
    && typeof value.caseSensitive === 'boolean'
    && typeof value.accentSensitive === 'boolean'
    && (value.numericTextMode === 'lexical' || value.numericTextMode === 'numeric')
    && (value.blankOrder === 'first' || value.blankOrder === 'last')
    && Array.isArray(value.typeOrder)
    && value.typeOrder.length === 5
    && new Set(value.typeOrder).size === 5
    && Array.isArray(value.customLists)
    && value.customLists.every((list) => Array.isArray(list) && list.every((entry) => typeof entry === 'string'));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
