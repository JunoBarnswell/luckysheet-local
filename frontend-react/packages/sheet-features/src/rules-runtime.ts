import type {
  CellData,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorksheetModel,
  WorkbookModel,
} from '@react-sheets/core-model';
import { resolveFilterCellValue } from '@react-sheets/core-model';
import type { FormulaEngine, FormulaValue, RowVisibility, RowVisibilityResolver } from '@react-sheets/formula-engine';
import { isArrayValue, isFormulaError } from '@react-sheets/formula-engine';
import { formatValue, type NumberFormatContext } from '@react-sheets/number-format';
import { createConditionalFormatRuntime, type ConditionalFormatRuntime, type ConditionalOverlay } from './data-features';
import { computeFilterHiddenRows, validateDataInput, type DataValidationResult, type FilterCellReader, type FilterVisualResolver, type RuleEvaluationOptions } from './data-features';
import { resolveAutoFilters, type FilterOwner, type ResolvedAutoFilter } from './sheet-table-features';
import { RuleIntervalIndex, resolveValidationRule, validationRuleIndex } from './rule-index';
export { RuleIntervalIndex, resolveValidationRule, validationRuleIndex } from './rule-index';

/**
 * One resolved cell contract for formula, rules, filter, find, chart and
 * print consumers. Authored storage and visibility are never conflated.
 */
export interface ResolvedCell {
  readonly address: { readonly sheetId: string; readonly row: number; readonly column: number };
  readonly cell?: CellData;
  readonly authoredValue: CellData['value'];
  readonly authoredFormula?: string;
  readonly calculatedValue: FormulaValue;
  readonly value: FormulaValue;
  readonly displayValue: string;
  readonly numberFormat?: string;
  readonly visibility: RowVisibility;
  readonly source: 'empty' | 'authored' | 'formula' | 'spill';
}

export interface ResolvedCellReader {
  resolve(sheetId: string, row: number, column: number): ResolvedCell;
  resolveRange(range: RangeRef): ResolvedCell[];
}

export interface ResolvedCellReaderOptions {
  readonly workbook: WorkbookModel;
  readonly formula?: FormulaEngine;
  readonly rowVisibility?: RowVisibilityResolver;
  readonly dateSystem?: '1900' | '1904';
  readonly locale?: string;
}

function scalarDisplayValue(value: FormulaValue, format: string | undefined, context: NumberFormatContext): string {
  if (isFormulaError(value)) return value.code;
  if (isArrayValue(value)) {
    const first = value[0]?.[0] ?? null;
    if (isFormulaError(first)) return first.code;
    return formatValue(first as string | number | boolean | null, format, context);
  }
  return formatValue(value as string | number | boolean | null, format, context);
}

/** Convert the core model's cached scalar/error shape into the formula engine's canonical value. */
function toFormulaValue(value: CellData['value'] | NonNullable<CellData['formulaValue']>): FormulaValue {
  if (value === null || typeof value !== 'object') return value;
  if (value.kind === 'error') {
    return { kind: 'error', code: value.code, message: value.message ?? '' };
  }
  throw new Error('RULE_RUNTIME_UNSUPPORTED_FORMULA_VALUE');
}

/** Build the canonical read projection; no caller may read CellData directly for derived behaviour. */
export function createResolvedCellReader(options: ResolvedCellReaderOptions): ResolvedCellReader {
  const numberFormatContext: NumberFormatContext = { dateSystem: options.dateSystem ?? '1900', locale: options.locale };
  const resolve = (sheetId: string, row: number, column: number): ResolvedCell => {
    const sheet = options.workbook.getSheet(sheetId);
    if (!Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0) throw new Error('RESOLVED_CELL_INVALID_ADDRESS');
    const cell = sheet.cells.get(row, column);
    const address = { sheetId, row, column };
    const authoredValue = cell?.value ?? null;
    const numberFormat = cell?.numberFormat ?? cell?.style?.numberFormat;
    const visibility = options.rowVisibility?.resolve(sheetId, row) ?? { manualHidden: false, filterHidden: false, outlineHidden: false };
    const spill = options.formula?.getSpillValueAt(sheetId, row, column);
    let calculatedValue: FormulaValue = toFormulaValue(authoredValue);
    let source: ResolvedCell['source'] = cell ? 'authored' : 'empty';
    if (spill !== undefined) {
      calculatedValue = spill;
      source = 'spill';
    } else if (cell?.formula !== undefined && !cell.formulaMetadata?.preservedOnly) {
      const result = options.formula?.getCellResult(address);
      if (!result) throw new Error(`RESOLVED_CELL_FORMULA_UNAVAILABLE: ${sheetId}!${row}:${column}`);
      calculatedValue = result.value;
      source = 'formula';
    } else if (cell?.formulaValue !== undefined) {
      calculatedValue = toFormulaValue(cell.formulaValue);
      source = 'formula';
    }
    return {
      address,
      ...(cell ? { cell: structuredClone(cell) } : {}),
      authoredValue,
      ...(cell?.formula !== undefined ? { authoredFormula: cell.formula } : {}),
      calculatedValue,
      value: calculatedValue,
      displayValue: scalarDisplayValue(calculatedValue, numberFormat, numberFormatContext),
      ...(numberFormat ? { numberFormat } : {}),
      visibility,
      source,
    };
  };
  return {
    resolve,
    resolveRange: (range) => {
      const result: ResolvedCell[] = [];
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) result.push(resolve(range.sheetId, row, column));
      }
      return result;
    },
  };
}

export interface RuleRuntimeSnapshot {
  readonly resolved: ResolvedCell;
  readonly validation?: DataValidationRule;
  readonly conditionalFormats: readonly ConditionalFormatRule[];
  readonly filterOwners: readonly ResolvedAutoFilter[];
  readonly overlay?: ConditionalOverlay;
}

/** Shared rule/index owner for a worksheet. */
export class RuleRuntime {
  readonly validation: RuleIntervalIndex<DataValidationRule>;
  readonly conditionalFormats: RuleIntervalIndex<ConditionalFormatRule>;
  private readonly conditionalRuntime: ConditionalFormatRuntime;
  private filterOwners: readonly ResolvedAutoFilter[];

  constructor(
    private readonly sheet: WorksheetModel,
    private readonly cells?: ResolvedCellReader,
    private readonly formula?: FormulaEngine,
  ) {
    this.validation = validationRuleIndex(sheet);
    this.conditionalFormats = new RuleIntervalIndex<ConditionalFormatRule>();
    for (const rule of sheet.conditionalFormats) for (const range of rule.ranges) this.conditionalFormats.add(range, rule);
    this.conditionalRuntime = createConditionalFormatRuntime(sheet, cells, formula);
    this.filterOwners = resolveAutoFilters(sheet);
  }

  resolve(row: number, column: number): RuleRuntimeSnapshot {
    const cell = this.sheet.cells.get(row, column);
    const authoredValue = cell?.value ?? null;
    const fallbackValue = cell?.formulaValue === undefined
      ? toFormulaValue(authoredValue)
      : toFormulaValue(cell.formulaValue);
    const resolved: ResolvedCell = this.cells?.resolve(this.sheet.id, row, column) ?? {
      address: { sheetId: this.sheet.id, row, column },
      ...(cell ? { cell: structuredClone(cell) } : {}),
      authoredValue,
      ...(cell?.formula ? { authoredFormula: cell.formula } : {}),
      calculatedValue: fallbackValue,
      value: fallbackValue,
      displayValue: scalarDisplayValue(fallbackValue, cell?.numberFormat ?? cell?.style?.numberFormat, {}),
      ...(cell?.numberFormat ?? cell?.style?.numberFormat ? { numberFormat: cell?.numberFormat ?? cell?.style?.numberFormat } : {}),
      visibility: { manualHidden: false, filterHidden: false, outlineHidden: false },
      source: cell ? 'authored' : 'empty',
    };
    const filterOwners = this.filterOwners.filter(({ autoFilter }) => autoFilter.range.startRow <= row && row <= autoFilter.range.endRow && autoFilter.range.startColumn <= column && column <= autoFilter.range.endColumn);
    return {
      resolved,
      validation: resolveValidationRule(this.sheet, row, column),
      conditionalFormats: this.conditionalFormats.query(this.sheet.id, row, column),
      filterOwners,
      overlay: this.conditionalRuntime.resolveCell(row, column),
    };
  }

  invalidate(): void {
    this.conditionalRuntime.invalidate();
    this.filterOwners = resolveAutoFilters(this.sheet);
  }

  validateInput(row: number, column: number, value: CellData['value']): DataValidationResult {
    const options: RuleEvaluationOptions = {
      ...(this.formula ? { formulaEngine: this.formula } : {}),
      ...(this.cells ? { readCell: (_sheet, targetRow, targetColumn) => this.cells!.resolve(this.sheet.id, targetRow, targetColumn).value } : {}),
    };
    return validateDataInput(this.sheet, row, column, value, options);
  }

  filterCellReader(): FilterCellReader {
    return (row, column) => {
      const resolved = this.cells?.resolve(this.sheet.id, row, column);
      if (resolved) return resolveFilterCellValue(resolved.cell, resolved.value);
      return resolveFilterCellValue(this.sheet.cells.get(row, column));
    };
  }

  computeHiddenRows(dateSystem: '1900' | '1904' = '1900', visualResolver?: FilterVisualResolver): Set<number> {
    const visual = visualResolver ?? ((row: number, column: number, cell?: CellData) => {
      const overlay = this.conditionalRuntime.resolveCell(row, column);
      return { style: { ...(cell?.style ?? {}), ...(overlay?.style ?? {}) }, nativeColor: cell?.filterMetadata?.color, nativeIcon: cell?.filterMetadata?.icon };
    });
    return computeFilterHiddenRows(this.sheet, this.filterCellReader(), dateSystem, visual);
  }
}

export function createRuleRuntime(sheet: WorksheetModel, cells?: ResolvedCellReader, formula?: FormulaEngine): RuleRuntime {
  return new RuleRuntime(sheet, cells, formula);
}

/** The filter owner is resolved once for every point; overlapping owners fail close. */
export function resolveUniqueFilterOwner(sheet: WorksheetModel, row: number, column: number): FilterOwner | undefined {
  const owners = resolveAutoFilters(sheet).filter(({ autoFilter }) => autoFilter.range.startRow <= row && row <= autoFilter.range.endRow && autoFilter.range.startColumn <= column && column <= autoFilter.range.endColumn);
  const unique = new Map(owners.map(({ owner }) => [owner.kind === 'table' ? `table:${owner.tableId}` : 'worksheet', owner]));
  if (unique.size > 1) throw new Error(`RULE_OWNER_AMBIGUOUS: overlapping worksheet/table filters at ${sheet.id}!${row}:${column}`);
  return owners[0]?.owner;
}
