/**
 * A workbook date is a calendar identity, not a host-time-zone Date.
 *
 * The only place where a JavaScript Date is used in this module is as a UTC
 * calendar arithmetic primitive. Callers consume the typed parts below and
 * never inspect local Date fields. This is important for AutoFilter and
 * formula evaluation: the same workbook serial must produce the same parts
 * on every machine.
 */
export type ExcelDateSystem = '1900' | '1904';

export interface CanonicalExcelDateParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  millisecond: number;
}

export interface CanonicalExcelDate extends CanonicalExcelDateParts {
  system: ExcelDateSystem;
  serial: number;
}

export interface ExcelDateEvaluationContext {
  dateSystem?: ExcelDateSystem;
  canonicalReferenceDate?: CanonicalExcelDateParts;
}

export const EXCEL_MS_PER_DAY = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 31);
const EXCEL_EPOCH_1904 = Date.UTC(1904, 0, 1);

function assertSystem(system: ExcelDateSystem): void {
  if (system !== '1900' && system !== '1904') throw new Error(`Unsupported Excel date system: ${String(system)}`);
}

function assertSerial(serial: number): void {
  if (!Number.isFinite(serial)) throw new Error(`Excel serial date must be finite: ${String(serial)}`);
}

function utcDateFromParts(parts: CanonicalExcelDateParts): Date {
  if (!Number.isInteger(parts.year) || parts.year < 1 || parts.year > 9999
    || !Number.isInteger(parts.month) || parts.month < 1 || parts.month > 12
    || !Number.isInteger(parts.day) || parts.day < 1 || parts.day > 31
    || !Number.isInteger(parts.hour) || parts.hour < 0 || parts.hour > 23
    || !Number.isInteger(parts.minute) || parts.minute < 0 || parts.minute > 59
    || !Number.isInteger(parts.second) || parts.second < 0 || parts.second > 59
    || !Number.isInteger(parts.millisecond) || parts.millisecond < 0 || parts.millisecond > 999) {
    throw new Error('Canonical Excel date parts are out of range');
  }
  const date = new Date(Date.UTC(2000, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second, parts.millisecond));
  date.setUTCFullYear(parts.year);
  if (date.getUTCFullYear() !== parts.year || date.getUTCMonth() + 1 !== parts.month || date.getUTCDate() !== parts.day
    || date.getUTCHours() !== parts.hour || date.getUTCMinutes() !== parts.minute
    || date.getUTCSeconds() !== parts.second || date.getUTCMilliseconds() !== parts.millisecond) {
    throw new Error('Canonical Excel date parts describe an invalid calendar date');
  }
  return date;
}

function partsFromUtcDate(date: Date): CanonicalExcelDateParts {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) throw new Error('Canonical Excel date requires a valid UTC date');
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour: date.getUTCHours(),
    minute: date.getUTCMinutes(),
    second: date.getUTCSeconds(),
    millisecond: date.getUTCMilliseconds(),
  };
}

/** Resolve a serial into workbook-calendar parts without host timezone APIs. */
export function canonicalExcelDateFromSerial(serial: number, system: ExcelDateSystem = '1900'): CanonicalExcelDate {
  assertSystem(system);
  assertSerial(serial);
  if (system === '1900' && serial === 60) throw new Error('Excel serial 60 is the non-existent 1900-02-29 leap-day sentinel');
  const wholeDays = Math.floor(serial);
  const fraction = serial - wholeDays;
  const adjustedDays = system === '1900' && wholeDays > 60 ? wholeDays - 1 : wholeDays;
  let milliseconds = Math.round(fraction * EXCEL_MS_PER_DAY);
  let dayCarry = 0;
  if (milliseconds >= EXCEL_MS_PER_DAY) {
    milliseconds -= EXCEL_MS_PER_DAY;
    dayCarry = 1;
  }
  const epoch = system === '1904' ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_1900;
  const date = new Date(epoch + (adjustedDays + dayCarry) * EXCEL_MS_PER_DAY + milliseconds);
  if (Number.isNaN(date.getTime())) throw new Error(`Excel serial date is outside the supported UTC range: ${serial}`);
  return { system, serial, ...partsFromUtcDate(date) };
}

export function canonicalExcelDatePartsFromSerial(serial: number, system: ExcelDateSystem = '1900'): CanonicalExcelDateParts {
  const { system: _system, serial: _serial, ...parts } = canonicalExcelDateFromSerial(serial, system);
  return parts;
}

/** Convert canonical parts to a serial, applying the workbook's 1900 bug rule. */
export function canonicalExcelDateToSerial(parts: CanonicalExcelDateParts, system: ExcelDateSystem = '1900'): number {
  assertSystem(system);
  const date = utcDateFromParts(parts);
  const epoch = system === '1904' ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_1900;
  const serial = (date.getTime() - epoch) / EXCEL_MS_PER_DAY;
  if (system === '1900' && serial >= 60) return serial + 1;
  return serial;
}

export function canonicalExcelDateFromUtcDate(date: Date, system: ExcelDateSystem = '1900'): CanonicalExcelDate {
  const parts = partsFromUtcDate(date);
  return { system, serial: canonicalExcelDateToSerial(parts, system), ...parts };
}

export function canonicalExcelDateToUtcDate(value: CanonicalExcelDateParts): Date {
  return utcDateFromParts(value);
}

export function canonicalExcelDateToIso(value: CanonicalExcelDateParts): string {
  const date = utcDateFromParts(value);
  return date.toISOString();
}

/** Accept only the canonical UTC timestamp; display strings are not dates. */
export function canonicalExcelDateFromValue(value: unknown, system: ExcelDateSystem = '1900'): CanonicalExcelDate | null {
  if (typeof value === 'number' && Number.isFinite(value)) return canonicalExcelDateFromSerial(value, system);
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)) return null;
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?Z$/);
  if (!match) return null;
  const millisecond = match[7] === undefined ? 0 : Number(match[7].padEnd(3, '0'));
  const parts = { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]), hour: Number(match[4]), minute: Number(match[5]), second: Number(match[6]), millisecond };
  const serial = canonicalExcelDateToSerial(parts, system);
  if (system === '1900' && serial === 60) throw new Error('Canonical Excel date cannot represent the non-existent 1900-02-29 leap-day sentinel');
  return { system, serial, ...parts };
}

/** Build a canonical date from explicit calendar parts (used for clock/ranges). */
export function canonicalExcelDateFromParts(parts: CanonicalExcelDateParts, system: ExcelDateSystem = '1900'): CanonicalExcelDate {
  const serial = canonicalExcelDateToSerial(parts, system);
  return { system, serial, ...parts };
}

export function canonicalExcelDateDayOfWeek(value: CanonicalExcelDateParts): number {
  return utcDateFromParts(value).getUTCDay();
}

export function shiftCanonicalExcelDate(value: CanonicalExcelDateParts, days: number, system: ExcelDateSystem = '1900'): CanonicalExcelDate {
  const date = utcDateFromParts(value);
  const shifted = new Date(date.getTime() + days * EXCEL_MS_PER_DAY);
  return canonicalExcelDateFromUtcDate(shifted, system);
}

export function compareCanonicalExcelDates(left: CanonicalExcelDate, right: CanonicalExcelDate): number {
  if (left.system !== right.system) throw new Error('Cannot compare dates from different Excel date systems');
  return left.serial - right.serial;
}
