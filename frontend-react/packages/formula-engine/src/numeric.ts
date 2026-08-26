import { createFormulaError, isFormulaError, type FormulaError, type FormulaValue } from './values';

export const EXCEL_SIGNIFICANT_DIGITS = 15;

export interface ExcelNumericContext {
  readonly significantDigits: number;
}

export const DEFAULT_EXCEL_NUMERIC_CONTEXT: ExcelNumericContext = Object.freeze({
  significantDigits: EXCEL_SIGNIFICANT_DIGITS,
});

export function normalizeExcelNumericContext(context?: Partial<ExcelNumericContext>): ExcelNumericContext {
  const significantDigits = context?.significantDigits ?? EXCEL_SIGNIFICANT_DIGITS;
  if (!Number.isSafeInteger(significantDigits) || significantDigits < 1 || significantDigits > EXCEL_SIGNIFICANT_DIGITS) {
    throw new Error(`Excel numeric precision must be an integer from 1 to ${EXCEL_SIGNIFICANT_DIGITS}`);
  }
  return { significantDigits };
}

/** Central scalar coercion used by arithmetic and numeric functions. */
export function coerceExcelNumber(value: FormulaValue | undefined): number | FormulaError {
  if (isFormulaError(value)) return value;
  if (value === undefined || value === null) return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : createFormulaError('#NUM!', 'Non-finite numeric value');
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = Number(value.trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return createFormulaError('#VALUE!', 'Expected a number');
}

/** Normalize a finite result to Excel's 15 significant-digit precision. */
export function normalizeExcelPrecision(value: number, contextOrDigits: ExcelNumericContext | number = DEFAULT_EXCEL_NUMERIC_CONTEXT): number {
  if (!Number.isFinite(value) || value === 0) return value;
  const significantDigits = typeof contextOrDigits === 'number' ? contextOrDigits : contextOrDigits.significantDigits;
  const exponent = Math.floor(Math.log10(Math.abs(value)));
  const digits = significantDigits - exponent - 1;
  return roundDecimal(value, digits, 'half-away-from-zero');
}

export function roundExcel(value: number, digits = 0): number {
  return roundDecimal(value, Math.trunc(digits), 'half-away-from-zero');
}

export function roundExcelUp(value: number, digits = 0): number {
  return roundDecimal(value, Math.trunc(digits), 'away-from-zero');
}

export function roundExcelDown(value: number, digits = 0): number {
  return roundDecimal(value, Math.trunc(digits), 'toward-zero');
}

export function truncateExcel(value: number, digits = 0): number {
  return roundExcelDown(value, digits);
}

function roundDecimal(value: number, digits: number, mode: 'half-away-from-zero' | 'away-from-zero' | 'toward-zero'): number {
  if (!Number.isFinite(value) || value === 0) return value;
  if (!Number.isSafeInteger(digits) || digits > 308 || digits < -308) return value;
  const sign = value < 0 ? -1 : 1;
  const shifted = decimalShift(Math.abs(value), digits);
  if (!Number.isFinite(shifted)) return value;
  const magnitude = mode === 'half-away-from-zero'
    ? Math.floor(shifted + 0.5)
    : mode === 'away-from-zero' ? Math.ceil(shifted) : Math.floor(shifted);
  return sign * decimalShift(magnitude, -digits);
}

/** Decimal exponent shifting avoids binary midpoint errors such as -1.475. */
function decimalShift(value: number, digits: number): number {
  const [coefficient, exponentText] = value.toString().split('e');
  const exponent = Number(exponentText ?? 0) + digits;
  return Number(`${coefficient}e${exponent}`);
}
