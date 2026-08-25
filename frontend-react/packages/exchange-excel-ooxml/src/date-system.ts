import type { DateSystem } from './types';
import {
  canonicalExcelDateFromSerial,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateFromValue,
  canonicalExcelDateToIso,
  canonicalExcelDateToUtcDate,
} from '@react-sheets/formula-engine';

export {
  canonicalExcelDateFromParts,
  canonicalExcelDateFromSerial,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateFromValue,
  canonicalExcelDatePartsFromSerial,
  canonicalExcelDateToIso,
  canonicalExcelDateToSerial,
  canonicalExcelDateToUtcDate,
  type CanonicalExcelDate,
  type CanonicalExcelDateParts,
  type ExcelDateSystem,
} from '@react-sheets/formula-engine';

/** Excel serial date — 与 JS Date 分离 */
export interface ExcelSerialDate {
  serial: number;
  system: DateSystem;
}

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
  return canonicalExcelDateToUtcDate(canonicalExcelDateFromSerial(serial, system));
}

export function dateToSerial(date: Date, system: DateSystem = '1900'): number {
  assertDateSystem(system);
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) throw new Error('Excel date must be a valid Date');
  return canonicalExcelDateFromUtcDate(date, system).serial;
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
  return canonicalExcelDateToIso(canonicalExcelDateFromSerial(serial, system));
}

/** Convert a canonical UTC ISO date back to an Excel serial at export time. */
export function canonicalDateToSerial(value: string, system: DateSystem): number {
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)) {
    throw new Error(`Canonical Excel date must be a UTC ISO timestamp: ${value}`);
  }
  const date = canonicalExcelDateFromValue(value, system);
  if (!date) throw new Error(`Canonical Excel date is invalid: ${value}`);
  return date.serial;
}
