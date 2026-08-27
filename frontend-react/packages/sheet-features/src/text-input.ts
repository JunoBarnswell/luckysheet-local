import { clearFormulaProvenance, type CellData, type CellEditorKind, type CellStyle, type WorkbookModel } from '@react-sheets/core-model';
import {
  canonicalExcelDateToSerial,
  formatFormula,
  parseFormula,
  type CanonicalExcelDateParts,
  type ExcelDateSystem,
} from '@react-sheets/formula-engine';

/** The lexical owner of an input string. Each source has an explicit contract. */
export type CellInputSourceKind = 'direct-entry' | 'clipboard-text' | 'import-text' | 'find-replace' | 'script-text';

export interface CellInputInterpretationContext {
  readonly sourceKind: CellInputSourceKind;
  readonly cultureId: string;
  readonly decimalSeparator: string;
  readonly groupSeparator: string;
  readonly dateSystem: ExcelDateSystem;
  readonly referenceDate: CanonicalExcelDateParts;
  readonly currentNumberFormat?: string;
  readonly currentCellType?: CellEditorKind;
  readonly inputOptions?: {
    readonly allowFormula?: boolean;
    readonly allowBoolean?: boolean;
    readonly allowDateTime?: boolean;
  };
}

export function isCellInputInterpretationContext(value: unknown): value is CellInputInterpretationContext {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const context = value as Record<string, unknown>;
  const referenceDate = context.referenceDate;
  const inputOptions = context.inputOptions;
  const optionsValid = inputOptions === undefined || (typeof inputOptions === 'object' && inputOptions !== null && !Array.isArray(inputOptions)
    && (typeof (inputOptions as Record<string, unknown>).allowFormula === 'undefined' || typeof (inputOptions as Record<string, unknown>).allowFormula === 'boolean')
    && (typeof (inputOptions as Record<string, unknown>).allowBoolean === 'undefined' || typeof (inputOptions as Record<string, unknown>).allowBoolean === 'boolean')
    && (typeof (inputOptions as Record<string, unknown>).allowDateTime === 'undefined' || typeof (inputOptions as Record<string, unknown>).allowDateTime === 'boolean'));
  return (context.sourceKind === 'direct-entry' || context.sourceKind === 'clipboard-text' || context.sourceKind === 'import-text' || context.sourceKind === 'find-replace' || context.sourceKind === 'script-text')
    && typeof context.cultureId === 'string' && typeof context.decimalSeparator === 'string' && typeof context.groupSeparator === 'string'
    && (context.currentNumberFormat === undefined || typeof context.currentNumberFormat === 'string')
    && (context.currentCellType === undefined || context.currentCellType === 'text' || context.currentCellType === 'number' || context.currentCellType === 'date' || context.currentCellType === 'list' || context.currentCellType === 'checkbox')
    && (context.dateSystem === '1900' || context.dateSystem === '1904')
    && optionsValid
    && Boolean(referenceDate && typeof referenceDate === 'object' && !Array.isArray(referenceDate)
      && Number.isInteger((referenceDate as Record<string, unknown>).year)
      && Number.isInteger((referenceDate as Record<string, unknown>).month)
      && Number.isInteger((referenceDate as Record<string, unknown>).day)
      && Number.isInteger((referenceDate as Record<string, unknown>).hour)
      && Number.isInteger((referenceDate as Record<string, unknown>).minute)
      && Number.isInteger((referenceDate as Record<string, unknown>).second)
      && Number.isInteger((referenceDate as Record<string, unknown>).millisecond));
}

export type NumberFormatIntent =
  | { readonly kind: 'preserve' }
  | { readonly kind: 'set'; readonly format: string };

/** Canonical result of the workbook input interpretation domain. */
export interface InterpretedCellInput {
  readonly value: CellData['value'];
  readonly valueType: 'empty' | 'text' | 'number' | 'boolean' | 'formula';
  readonly formula?: string;
  readonly numberFormatIntent: NumberFormatIntent;
  readonly lexicalSource: string;
}

/** Backward name retained for the public sheet-features API; semantics are now contextual. */
export type ParsedCellText = InterpretedCellInput;

/**
 * Build an explicit input context from workbook-owned culture and runtime date
 * state. This is the only production factory for string input contexts.
 */
export function createCellInputInterpretationContext(
  workbook: WorkbookModel,
  options: {
    sourceKind: CellInputSourceKind;
    dateSystem: ExcelDateSystem;
    referenceDate: CanonicalExcelDateParts;
    cell?: CellData;
  },
): CellInputInterpretationContext {
  const cultureId = workbook.collationContext.cultureId.trim();
  if (!cultureId) throw new Error('Workbook culture is required for cell input interpretation');
  const parts = cultureId === 'invariant' ? undefined : new Intl.NumberFormat(cultureId).formatToParts(12345.6);
  const decimalSeparator = cultureId === 'invariant' ? '.' : parts?.find((part) => part.type === 'decimal')?.value;
  const groupSeparator = cultureId === 'invariant' ? ',' : parts?.find((part) => part.type === 'group')?.value;
  if (!decimalSeparator || !groupSeparator || decimalSeparator === groupSeparator) {
    throw new Error(`Workbook culture has no stable numeric separators: ${cultureId}`);
  }
  return {
    sourceKind: options.sourceKind,
    cultureId,
    decimalSeparator,
    groupSeparator,
    dateSystem: options.dateSystem,
    referenceDate: structuredClone(options.referenceDate),
    currentNumberFormat: options.cell?.numberFormat ?? options.cell?.style?.numberFormat,
    currentCellType: options.cell?.editor?.kind,
  };
}

/** Interpret one raw lexical value using the complete workbook/input context. */
export function interpretCellInput(text: string, context: CellInputInterpretationContext): InterpretedCellInput {
  assertInputContext(context);
  const preserve = { kind: 'preserve' } as const;

  // Apostrophe escaping is a direct-entry lexical rule only. It is never
  // persisted in canonical CellData and is not silently applied to clipboard.
  if (context.sourceKind === 'direct-entry' && text.startsWith("'")) {
    return result(text.slice(1), 'text', text, preserve);
  }

  // Text format and text editors opt out before formula/number/date coercion.
  if (isTextFormat(context.currentNumberFormat) || context.currentCellType === 'text') {
    return result(text, text === '' ? 'empty' : 'text', text, preserve);
  }

  if (text.startsWith('=') && context.inputOptions?.allowFormula !== false) {
    parseFormula(text);
    return { value: null, valueType: 'formula', formula: text, numberFormatIntent: preserve, lexicalSource: text };
  }
  if (text === '') return result(null, 'empty', text, preserve);

  if (context.inputOptions?.allowDateTime !== false) {
    const dateTime = parseDateTimeLiteral(text, context);
    if (dateTime) {
    return {
        value: dateTime.serial,
        valueType: 'number',
        numberFormatIntent: isDateTimeFormat(context.currentNumberFormat)
          ? preserve
          : { kind: 'set', format: dateTime.format },
        lexicalSource: text,
      };
    }
  }

  const percent = text.endsWith('%');
  const numericText = percent ? text.slice(0, -1) : text;
  const currency = parseCurrencyLexeme(numericText);
  const normalizedNumericText = currency?.value ?? numericText;
  const number = parseCultureNumber(normalizedNumericText, context);
  if (number !== undefined) {
    const value = percent || isPercentFormat(context.currentNumberFormat) ? number / 100 : number;
    const format = percent && !isPercentFormat(context.currentNumberFormat)
      ? { kind: 'set' as const, format: '0%' }
      : currency?.format && (!context.currentNumberFormat || /^general$/i.test(context.currentNumberFormat.trim()))
        ? { kind: 'set' as const, format: currency.format }
        : preserve;
    return result(value, 'number', text, format);
  }

  const fraction = parseFraction(normalizedNumericText, context);
  if (fraction !== undefined) return result(fraction, 'number', text, preserve);

  if (context.inputOptions?.allowBoolean !== false) {
    if (/^true$/i.test(text)) return result(true, 'boolean', text, preserve);
    if (/^false$/i.test(text)) return result(false, 'boolean', text, preserve);
  }
  return result(text, 'text', text, preserve);
}

/** Parse one direct or host text input with an explicit interpretation context. */
export function parseCellText(text: string, context: CellInputInterpretationContext): ParsedCellText {
  return interpretCellInput(text, context);
}

/**
 * Construct the next cell while retaining presentation and auxiliary metadata.
 * The interpreter owns format intent so display cannot repair a wrong value.
 */
export function buildCellFromText(
  text: string,
  previous: CellData | undefined,
  context: CellInputInterpretationContext,
  style?: Partial<CellStyle>,
): CellData {
  const parsed = interpretCellInput(text, context);
  const next: CellData = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
  next.value = parsed.value;
  delete next.formula;
  delete next.formulaValue;
  delete next.displayValue;
  if (parsed.formula !== undefined) next.formula = parsed.formula;
  if (style !== undefined) next.style = { ...(next.style ?? {}), ...structuredClone(style) };
  if (parsed.numberFormatIntent.kind === 'set') {
    next.numberFormat = parsed.numberFormatIntent.format;
    next.style = { ...(next.style ?? {}), numberFormat: parsed.numberFormatIntent.format };
  }
  return next;
}

/** Return canonical formula text for callers that inspect a parsed formula. */
export function canonicalizeFormulaText(formula: string): string {
  return formatFormula(parseFormula(formula));
}

function result(
  value: CellData['value'],
  valueType: InterpretedCellInput['valueType'],
  lexicalSource: string,
  numberFormatIntent: NumberFormatIntent,
): InterpretedCellInput {
  return { value, valueType, lexicalSource, numberFormatIntent };
}

function assertInputContext(context: CellInputInterpretationContext): void {
  if (!isCellInputInterpretationContext(context)) throw new Error('Cell input interpretation context is invalid');
  if (!context.cultureId.trim()) throw new Error('Cell input culture is required');
  if (!context.decimalSeparator || !context.groupSeparator || context.decimalSeparator === context.groupSeparator) {
    throw new Error('Cell input numeric separators are invalid');
  }
  if (context.dateSystem !== '1900' && context.dateSystem !== '1904') throw new Error(`Unsupported cell input date system: ${String(context.dateSystem)}`);
  if (!Number.isInteger(context.referenceDate.year) || !Number.isInteger(context.referenceDate.month) || !Number.isInteger(context.referenceDate.day)) {
    throw new Error('Cell input reference date is invalid');
  }
}

function isTextFormat(format: string | undefined): boolean {
  if (!format) return false;
  return format.split(';')[0]?.replace(/"(?:[^"]|"")*"/g, '').replace(/\\./g, '').includes('@') ?? false;
}

function isPercentFormat(format: string | undefined): boolean {
  return Boolean(format && format.replace(/"(?:[^"]|"")*"/g, '').includes('%'));
}

function isDateTimeFormat(format: string | undefined): boolean {
  if (!format) return false;
  const unquoted = format.replace(/"(?:[^"]|"")*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[ydhms]/i.test(unquoted);
}

function parseCultureNumber(text: string, context: CellInputInterpretationContext): number | undefined {
  if (!text) return undefined;
  const decimal = escapeRegExp(context.decimalSeparator);
  const group = escapeRegExp(context.groupSeparator);
  const groupedInteger = `(?:\\d{1,3}(?:${group}\\d{3})+)`;
  const integer = `(?:\\d+)`;
  const fraction = `(?:${decimal}\\d+)?`;
  const exponent = `(?:[eE][+-]?\\d+)?`;
  const pattern = new RegExp(`^[+-]?(?:${groupedInteger}|${integer})${fraction}${exponent}$|^[+-]?${decimal}\\d+${exponent}$`);
  if (!pattern.test(text)) return undefined;
  const normalized = text.split(context.groupSeparator).join('').replace(context.decimalSeparator, '.');
  const value = Number(normalized);
  if (!Number.isFinite(value)) throw new Error(`Numeric input is not finite: ${text}`);
  return value;
}

function parseFraction(text: string, context: CellInputInterpretationContext): number | undefined {
  const fractionFormat = isFractionFormat(context.currentNumberFormat);
  const match = text.match(/^([+-]?\d+)(?:\s+(\d+))?\/(\d+)$/);
  if (!match || (!fractionFormat && match[2] === undefined)) return undefined;
  const whole = Number(match[1]);
  const numerator = Number(match[2] ?? match[1]);
  const denominator = Number(match[3]);
  if (!Number.isFinite(whole) || !Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return undefined;
  return match[2] === undefined ? numerator / denominator : whole + Math.sign(whole || 1) * (numerator / denominator);
}

function isFractionFormat(format: string | undefined): boolean {
  if (!format) return false;
  const unquoted = format.replace(/"(?:[^"]|"")*"/g, '').replace(/\\./g, '');
  return /[0#?][^;]*\/[0#?]/.test(unquoted) && !/[ydhms]/i.test(unquoted);
}

function parseCurrencyLexeme(text: string): { value: string; format?: string } | undefined {
  const symbols: Array<[string, string]> = [['$', '$#,##0.00'], ['€', '€#,##0.00'], ['£', '£#,##0.00'], ['¥', '¥#,##0.00'], ['￥', '¥#,##0.00']];
  for (const [symbol, format] of symbols) {
    if (text.startsWith(symbol)) return { value: text.slice(symbol.length), format };
    if (text.endsWith(symbol)) return { value: text.slice(0, -symbol.length), format };
    if (text.startsWith(`-${symbol}`) || text.startsWith(`+${symbol}`)) return { value: `${text[0]}${text.slice(symbol.length + 1)}`, format };
  }
  return undefined;
}

function parseDateTimeLiteral(text: string, context: CellInputInterpretationContext): { serial: number; format: string } | undefined {
  const timeOnly = parseTime(text);
  if (timeOnly) {
    const serial = canonicalExcelDateToSerial({ ...context.referenceDate, hour: timeOnly.hour, minute: timeOnly.minute, second: timeOnly.second, millisecond: timeOnly.millisecond }, context.dateSystem);
    return { serial, format: 'h:mm' };
  }
  const separator = text.includes('T') ? 'T' : text.includes(' ') ? ' ' : undefined;
  const dateText = separator ? text.split(separator)[0] ?? '' : text;
  const timeText = separator ? text.slice(dateText.length + 1) : '';
  const date = parseDate(dateText, context);
  if (!date) return undefined;
  const time = timeText ? parseTime(timeText) : undefined;
  if (timeText && !time) return undefined;
  const serial = canonicalExcelDateToSerial({
    year: date.year,
    month: date.month,
    day: date.day,
    hour: time?.hour ?? 0,
    minute: time?.minute ?? 0,
    second: time?.second ?? 0,
    millisecond: time?.millisecond ?? 0,
  }, context.dateSystem);
  const dateFormat = dateOrder(context.cultureId) === 'mdy'
    ? 'm/d/yy'
    : dateOrder(context.cultureId) === 'ymd'
      ? 'yy/m/d'
      : 'd/m/yy';
  return { serial, format: time ? `${dateFormat} h:mm` : dateFormat };
}

function parseDate(text: string, context: CellInputInterpretationContext): { year: number; month: number; day: number } | undefined {
  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (iso) return validDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  const parts = text.match(/^(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})$/);
  if (parts) {
    const first = Number(parts[1]);
    const second = Number(parts[2]);
    const third = Number(parts[3]);
    if (first >= 1000) return validDate(first, second, third);
    const order = dateOrder(context.cultureId);
    const values = order === 'mdy' ? { month: first, day: second, year: third } : { day: first, month: second, year: third };
    const year = values.year < 100 ? 2000 + values.year : values.year;
    return validDate(year, values.month, values.day);
  }
  const short = text.match(/^(\d{1,2})[/-](\d{1,2})$/);
  if (!short) return undefined;
  return validDate(context.referenceDate.year, Number(short[1]), Number(short[2]));
}

function parseTime(text: string): { hour: number; minute: number; second: number; millisecond: number } | undefined {
  const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(?:\s*(AM|PM|A|P))?$/i);
  if (!match) return undefined;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] ?? 0);
  const millisecond = Number((match[4] ?? '').padEnd(3, '0') || 0);
  const meridiem = match[5]?.toUpperCase();
  if (meridiem && (hour < 1 || hour > 12)) return undefined;
  if (!meridiem && hour > 23) return undefined;
  if (meridiem) hour = hour % 12 + (meridiem === 'PM' || meridiem === 'P' ? 12 : 0);
  if (minute > 59 || second > 59 || millisecond > 999) return undefined;
  return { hour, minute, second, millisecond };
}

function dateOrder(cultureId: string): 'mdy' | 'dmy' | 'ymd' {
  if (/^(zh|ja|ko|hu|lt|sv|fi|da|nb|no|tr)/i.test(cultureId)) return 'ymd';
  if (/^(en-US|en-CA)/i.test(cultureId)) return 'mdy';
  return 'dmy';
}

function validDate(year: number, month: number, day: number): { year: number; month: number; day: number } | undefined {
  try {
    const parts = { year, month, day, hour: 0, minute: 0, second: 0, millisecond: 0 };
    canonicalExcelDateToSerial(parts, '1900');
    return { year, month, day };
  } catch {
    return undefined;
  }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
