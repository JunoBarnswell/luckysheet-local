import { createFormulaError, type FormulaValue } from '../values';
import {
  canonicalExcelDateFromParts,
  canonicalExcelDateFromSerial,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateFromValue,
  canonicalExcelDateToSerial,
  canonicalExcelDateDayOfWeek,
  shiftCanonicalExcelDate,
  type CanonicalExcelDate,
  type ExcelDateEvaluationContext,
} from '../excel-date';

function parseDateInput(value: FormulaValue | undefined, context?: ExcelDateEvaluationContext): CanonicalExcelDate | null {
  if (value === undefined || value === null) return null;
  const system = context?.dateSystem ?? '1900';
  if (value instanceof Date) return canonicalExcelDateFromUtcDate(value, system);
  if (typeof value === 'number') return canonicalExcelDateFromSerial(value, system);
  if (typeof value === 'string') return canonicalExcelDateFromValue(value, system);
  return null;
}

function error(error: unknown, fallback: string): ReturnType<typeof createFormulaError> {
  return createFormulaError('#VALUE!', error instanceof Error ? error.message : fallback);
}

export const datetimeFunctions: Record<string, (args: FormulaValue[], context?: ExcelDateEvaluationContext) => FormulaValue> = {
  DATE: (args, context) => {
    const year = Number(args[0]); const month = Number(args[1]); const day = Number(args[2]);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return createFormulaError('#VALUE!', 'Invalid date parameters');
    try { const system = context?.dateSystem ?? '1900'; return Math.round(canonicalExcelDateToSerial(canonicalExcelDateFromParts({ year, month, day, hour: 0, minute: 0, second: 0, millisecond: 0 }, system), system)); }
    catch (cause) { return error(cause, 'Invalid date parameters'); }
  },
  DATEVALUE: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? Math.floor(date.serial) : createFormulaError('#VALUE!', 'Cannot parse date in DATEVALUE'); } catch (cause) { return error(cause, 'Cannot parse date in DATEVALUE'); } },
  DAY: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.day : createFormulaError('#VALUE!', 'Invalid date in DAY'); } catch (cause) { return error(cause, 'Invalid date in DAY'); } },
  MONTH: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.month : createFormulaError('#VALUE!', 'Invalid date in MONTH'); } catch (cause) { return error(cause, 'Invalid date in MONTH'); } },
  YEAR: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.year : createFormulaError('#VALUE!', 'Invalid date in YEAR'); } catch (cause) { return error(cause, 'Invalid date in YEAR'); } },
  HOUR: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.hour : createFormulaError('#VALUE!', 'Invalid date in HOUR'); } catch (cause) { return error(cause, 'Invalid date in HOUR'); } },
  MINUTE: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.minute : createFormulaError('#VALUE!', 'Invalid date in MINUTE'); } catch (cause) { return error(cause, 'Invalid date in MINUTE'); } },
  SECOND: (args, context) => { try { const date = parseDateInput(args[0], context); return date ? date.second : createFormulaError('#VALUE!', 'Invalid date in SECOND'); } catch (cause) { return error(cause, 'Invalid date in SECOND'); } },
  TODAY: (_args, context) => {
    if (!context?.canonicalReferenceDate) return createFormulaError('#VALUE!', 'TODAY requires an explicit canonical workbook reference date');
    try { const system = context.dateSystem ?? '1900'; return Math.floor(canonicalExcelDateToSerial(canonicalExcelDateFromParts({ ...context.canonicalReferenceDate, hour: 0, minute: 0, second: 0, millisecond: 0 }, system), system)); }
    catch (cause) { return error(cause, 'Invalid canonical workbook reference date'); }
  },
  NOW: (_args, context) => {
    if (!context?.canonicalReferenceDate) return createFormulaError('#VALUE!', 'NOW requires an explicit canonical workbook reference date');
    try { const system = context.dateSystem ?? '1900'; return canonicalExcelDateToSerial(canonicalExcelDateFromParts(context.canonicalReferenceDate, system), system); }
    catch (cause) { return error(cause, 'Invalid canonical workbook reference date'); }
  },
  WEEKDAY: (args, context) => {
    try {
      const date = parseDateInput(args[0], context);
      if (!date) return createFormulaError('#VALUE!', 'Invalid date in WEEKDAY');
      const returnType = args[1] === undefined ? 1 : Number(args[1]);
      if (!Number.isInteger(returnType) || ![1, 2, 3, 11, 12, 13, 14, 15, 16, 17].includes(returnType)) return createFormulaError('#NUM!', 'Invalid WEEKDAY return_type');
      const sundayIndex = canonicalExcelDateDayOfWeek(date);
      if (returnType === 1) return sundayIndex + 1;
      const mondayIndex = (sundayIndex + 6) % 7;
      if (returnType === 2 || returnType === 11) return mondayIndex + 1;
      if (returnType === 3) return mondayIndex;
      return ((mondayIndex - (returnType - 11)) + 7) % 7 + 1;
    } catch (cause) { return error(cause, 'Invalid date in WEEKDAY'); }
  },
  EDATE: (args, context) => {
    try {
      const date = parseDateInput(args[0], context);
      const months = Number(args[1]);
      if (!date || !Number.isFinite(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EDATE');
      const offset = Math.trunc(months);
      const absoluteMonth = date.year * 12 + (date.month - 1) + offset;
      const year = Math.floor(absoluteMonth / 12);
      const month = absoluteMonth - year * 12 + 1;
      const firstOfNext = canonicalExcelDateFromParts({ year: month === 12 ? year + 1 : year, month: month === 12 ? 1 : month + 1, day: 1, hour: 0, minute: 0, second: 0, millisecond: 0 }, date.system);
      const lastDay = shiftCanonicalExcelDate(firstOfNext, -1, date.system).day;
      return Math.round(canonicalExcelDateToSerial(canonicalExcelDateFromParts({ ...date, year, month, day: Math.min(date.day, lastDay) }, date.system), date.system));
    } catch (cause) { return error(cause, 'Invalid arguments in EDATE'); }
  },
  EOMONTH: (args, context) => {
    try {
      const date = parseDateInput(args[0], context);
      const months = Number(args[1]);
      if (!date || !Number.isFinite(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EOMONTH');
      const offset = Math.trunc(months);
      const absoluteMonth = date.year * 12 + (date.month - 1) + offset;
      const year = Math.floor(absoluteMonth / 12);
      const month = absoluteMonth - year * 12 + 1;
      const firstOfNext = canonicalExcelDateFromParts({ year: month === 12 ? year + 1 : year, month: month === 12 ? 1 : month + 1, day: 1, hour: 0, minute: 0, second: 0, millisecond: 0 }, date.system);
      return Math.round(canonicalExcelDateToSerial(shiftCanonicalExcelDate(firstOfNext, -1, date.system), date.system));
    } catch (cause) { return error(cause, 'Invalid arguments in EOMONTH'); }
  },
  DAYS: (args, context) => { try { const end = parseDateInput(args[0], context); const start = parseDateInput(args[1], context); if (!end || !start || end.system !== start.system) return createFormulaError('#VALUE!', 'Invalid dates in DAYS'); return Math.round(end.serial - start.serial); } catch (cause) { return error(cause, 'Invalid dates in DAYS'); } },
};
