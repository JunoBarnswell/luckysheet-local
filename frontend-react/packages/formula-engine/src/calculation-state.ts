import type { CellAddress } from './ast';
import { isWorkbookCalculationSettings, type WorkbookCalculationSettings } from './calculation-settings';
import type { SheetTableRef } from './sheet-table-resolver';
import type { ScalarValue } from './values';
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
}

export function assertFormulaCalculationSnapshot(value: unknown): asserts value is FormulaCalculationSnapshot {
  if (!isRecord(value)) throw new Error('Calculation snapshot must be an object');
  if (typeof value.defaultSheetId !== 'string' || value.defaultSheetId.length === 0) {
    throw new Error('Calculation snapshot requires a default worksheet id');
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
  if (!Array.isArray(value.pendingRoots) || !value.pendingRoots.every(isCellAddress)) {
    throw new Error('Calculation snapshot has invalid dirty roots');
  }
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
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && value.sheetId.length > 0
    && isNonNegativeInteger(value.rowCount)
    && isNonNegativeInteger(value.columnCount)
    && Array.isArray(value.occupied)
    && value.occupied.every(isCellAddress);
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
  return Number.isInteger(value.year) && value.year >= 1 && value.year <= 9999
    && Number.isInteger(value.month) && value.month >= 1 && value.month <= 12
    && Number.isInteger(value.day) && value.day >= 1 && value.day <= 31
    && Number.isInteger(value.hour) && value.hour >= 0 && value.hour <= 23
    && Number.isInteger(value.minute) && value.minute >= 0 && value.minute <= 59
    && Number.isInteger(value.second) && value.second >= 0 && value.second <= 59
    && Number.isInteger(value.millisecond) && value.millisecond >= 0 && value.millisecond <= 999;
}

function isExcelNumericContext(value: unknown): value is ExcelNumericContext {
  if (!isRecord(value)) return false;
  return Number.isSafeInteger(value.significantDigits)
    && value.significantDigits >= 1
    && value.significantDigits <= 15;
}

function isCalculationEntropyContext(value: unknown): value is CalculationEntropyContext {
  if (!isRecord(value)) return false;
  return Number.isSafeInteger(value.cycleId)
    && value.cycleId >= 0
    && typeof value.entropySeed === 'string'
    && value.entropySeed.trim().length > 0
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
