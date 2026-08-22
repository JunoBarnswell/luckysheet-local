import { createFormulaError, isFormulaError, type FormulaValue } from '../values';

function parseDateInput(val: FormulaValue | undefined): Date | null {
  if (val === undefined || val === null) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel serial date: 1 = Jan 1, 1900, 25569 = Jan 1, 1970 UTC
    const ms = Math.round((val - 25569) * 86400 * 1000);
    return new Date(ms);
  }
  if (typeof val === 'string') {
    const d = new Date(val);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

function toSerialDate(date: Date): number {
  const ms = date.getTime();
  return ms / (86400 * 1000) + 25569;
}

export const datetimeFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  DATE: (args) => {
    const year = Number(args[0]);
    const month = Number(args[1]);
    const day = Number(args[2]);
    if (Number.isNaN(year) || Number.isNaN(month) || Number.isNaN(day)) {
      return createFormulaError('#VALUE!', 'Invalid date parameters');
    }
    const d = new Date(Date.UTC(year, month - 1, day, 0, 0, 0, 0));
    return Math.round(toSerialDate(d));
  },

  DATEVALUE: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Cannot parse date in DATEVALUE');
    return Math.floor(toSerialDate(d));
  },

  DAY: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in DAY');
    return d.getUTCDate();
  },

  MONTH: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in MONTH');
    return d.getUTCMonth() + 1;
  },

  YEAR: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in YEAR');
    return d.getUTCFullYear();
  },

  HOUR: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in HOUR');
    return d.getUTCHours();
  },

  MINUTE: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in MINUTE');
    return d.getUTCMinutes();
  },

  SECOND: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in SECOND');
    return d.getUTCSeconds();
  },

  TODAY: () => {
    const now = new Date();
    const d = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    return Math.round(toSerialDate(d));
  },

  NOW: () => {
    return toSerialDate(new Date());
  },

  WEEKDAY: (args) => {
    const d = parseDateInput(args[0]);
    if (!d) return createFormulaError('#VALUE!', 'Invalid date in WEEKDAY');
    const returnType = args[1] !== undefined ? Number(args[1]) : 1;
    const day = d.getUTCDay(); // 0 = Sunday, 1 = Monday ... 6 = Saturday
    if (returnType === 1) return day + 1; // 1 = Sun, 7 = Sat
    if (returnType === 2) return day === 0 ? 7 : day; // 1 = Mon, 7 = Sun
    if (returnType === 3) return day === 0 ? 6 : day - 1; // 0 = Mon, 6 = Sun
    return day + 1;
  },

  EDATE: (args) => {
    const d = parseDateInput(args[0]);
    const months = Number(args[1]);
    if (!d || Number.isNaN(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EDATE');
    const next = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months, d.getUTCDate()));
    return Math.round(toSerialDate(next));
  },

  EOMONTH: (args) => {
    const d = parseDateInput(args[0]);
    const months = Number(args[1]);
    if (!d || Number.isNaN(months)) return createFormulaError('#VALUE!', 'Invalid arguments in EOMONTH');
    const next = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months + 1, 0));
    return Math.round(toSerialDate(next));
  },

  DAYS: (args) => {
    const end = parseDateInput(args[0]);
    const start = parseDateInput(args[1]);
    if (!end || !start) return createFormulaError('#VALUE!', 'Invalid dates in DAYS');
    const diff = end.getTime() - start.getTime();
    return Math.round(diff / (86400 * 1000));
  },
};
