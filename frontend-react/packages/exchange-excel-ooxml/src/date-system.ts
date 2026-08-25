import type { DateSystem } from './types';

/** Excel serial date — 与 JS Date 分离 */
export interface ExcelSerialDate {
  serial: number;
  system: DateSystem;
}

const MS_PER_DAY = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 31);
const EXCEL_EPOCH_1904 = Date.UTC(1904, 0, 1);

function assertDateSystem(system: DateSystem): void {
  if (system !== '1900' && system !== '1904') throw new Error(`Unsupported Excel date system: ${String(system)}`);
}

function assertFiniteSerial(serial: number): void {
  if (!Number.isFinite(serial)) throw new Error(`Excel serial date must be finite: ${String(serial)}`);
}

/** 1900 系统闰年 bug: Excel 错误地认为 1900 是闰年 */
export function serialToExcelDate(serial: number, system: DateSystem = '1900'): Date {
  assertDateSystem(system);
  assertFiniteSerial(serial);
  if (system === '1900' && serial === 60) throw new Error('Excel serial 60 is the non-existent 1900-02-29 leap-day sentinel');
  const adjusted = system === '1900' && serial > 60 ? serial - 1 : serial;
  const result = new Date((system === '1904' ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_1900) + adjusted * MS_PER_DAY);
  if (Number.isNaN(result.getTime())) throw new Error(`Excel serial date is outside the supported UTC range: ${serial}`);
  return result;
}

export function dateToSerial(date: Date, system: DateSystem = '1900'): number {
  assertDateSystem(system);
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) throw new Error('Excel date must be a valid Date');
  const epoch = system === '1904' ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_1900;
  let serial = (date.getTime() - epoch) / MS_PER_DAY;
  if (system === '1900' && serial >= 60) serial += 1;
  return serial;
}

export function parseDateSystem(workbookXml?: string): DateSystem {
  const value = workbookXml?.match(/<(?:[A-Za-z_][\w.-]*:)?workbookPr\b[^>]*\bdate1904\s*=\s*["']([^"']+)["']/i)?.[1];
  if (value === undefined || value === '0' || value.toLowerCase() === 'false') return '1900';
  if (value === '1' || value.toLowerCase() === 'true') return '1904';
  throw new Error(`Invalid Excel date1904 value: ${value}`);
}

/** True when an OOXML number format denotes a calendar/time serial. */
export function isExcelDateFormat(format: string | undefined): boolean {
  if (!format) return false;
  const unquoted = format
    .replace(/"(?:[^"]|"")*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, (token) => /h|m|s/i.test(token) ? token : '');
  return /[ydhms]/i.test(unquoted);
}

/** Convert an Excel serial at the OOXML boundary into the canonical UTC value. */
export function serialToCanonicalDate(serial: number, system: DateSystem): string {
  return serialToExcelDate(serial, system).toISOString();
}

/** Convert a canonical UTC ISO date back to an Excel serial at export time. */
export function canonicalDateToSerial(value: string, system: DateSystem): number {
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)) {
    throw new Error(`Canonical Excel date must be a UTC ISO timestamp: ${value}`);
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) throw new Error(`Canonical Excel date is invalid: ${value}`);
  return dateToSerial(date, system);
}
