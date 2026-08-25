import { createFormulaError, type FormulaValue } from '../values';
import {
  canonicalExcelDateFromParts,
  canonicalExcelDateFromSerial,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateFromValue,
  canonicalExcelDateToSerial,
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
  WEEKDAY: (args, context) => { try { const date = parseDateInput(args[0], context); if (!date) return createFormulaError('#VALUE!', 'Invalid date in WEEKDAY'); const day = new Date(Date.UTC(2000, date.month - 1, date.day)).getUTCDay(); const returnType = args[1] === undefined ? 1 : Number(args[1]); if (returnType === 1) return day + 1; if (returnType === 2) return day === 0 ? 7 : day; if (returnType === 3) return day === 0 ? 6 : day - 1; return day + 1; } catch (cause) { return error(cause, 'Invalid date in WEEKDAY'); } },
  EDATE: (args, context) => { try { const date = parseDateInput(args[0], context); const months = Number(args[1]); if (!date || !Number.isFinite(months) || !Number.isInteger(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EDATE'); const utc = new Date(Date.UTC(2000, date.month - 1 + months, date.day, date.hour, date.minute, date.second, date.millisecond)); utc.setUTCFullYear(date.year); return Math.round(canonicalExcelDateFromUtcDate(utc, date.system).serial); } catch (cause) { return error(cause, 'Invalid arguments in EDATE'); } },
  EOMONTH: (args, context) => { try { const date = parseDateInput(args[0], context); const months = Number(args[1]); if (!date || !Number.isFinite(months) || !Number.isInteger(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EOMONTH'); const utc = new Date(Date.UTC(2000, date.month + months, 0)); utc.setUTCFullYear(date.year); return Math.round(canonicalExcelDateFromUtcDate(utc, date.system).serial); } catch (cause) { return error(cause, 'Invalid arguments in EOMONTH'); } },
  DAYS: (args, context) => { try { const end = parseDateInput(args[0], context); const start = parseDateInput(args[1], context); if (!end || !start || end.system !== start.system) return createFormulaError('#VALUE!', 'Invalid dates in DAYS'); return Math.round(end.serial - start.serial); } catch (cause) { return error(cause, 'Invalid dates in DAYS'); } },
};
