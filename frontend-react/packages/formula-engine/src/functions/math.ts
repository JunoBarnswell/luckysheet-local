import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { coerceExcelNumber, normalizeExcelPrecision, roundExcel, roundExcelDown, roundExcelUp, truncateExcel } from '../numeric';
import type { FormulaEvaluationContext } from '../evaluator';

export function flattenNumericArgs(args: FormulaValue[]): number[] | ReturnType<typeof createFormulaError> {
  const numbers: number[] = [];
  for (const arg of args) {
    if (isFormulaError(arg)) return arg;
    if (Array.isArray(arg)) {
      for (const row of arg) {
        if (Array.isArray(row)) {
          for (const cell of row) {
            if (isFormulaError(cell)) return cell;
            if (typeof cell === 'number' && Number.isFinite(cell)) numbers.push(cell);
          }
        } else if (typeof row === 'number' && Number.isFinite(row)) {
          numbers.push(row);
        }
      }
    } else if (typeof arg === 'number' && Number.isFinite(arg)) {
      numbers.push(arg);
    } else if (typeof arg === 'string') {
      const parsed = coerceExcelNumber(arg);
      if (!isFormulaError(parsed)) numbers.push(parsed);
    } else if (typeof arg === 'boolean') {
      numbers.push(arg ? 1 : 0);
    }
  }
  return numbers;
}

export const mathFunctions: Record<string, (args: FormulaValue[], context?: FormulaEvaluationContext) => FormulaValue> = {
  SUM: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    return normalizeExcelPrecision(nums.reduce((acc, n) => acc + n, 0));
  },

  PRODUCT: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return 0;
    return normalizeExcelPrecision(nums.reduce((acc, n) => acc * n, 1));
  },

  ABS: (args) => {
    const arg = args[0];
    if (isFormulaError(arg)) return arg;
    const n = coerceExcelNumber(arg);
    if (isFormulaError(n)) return n;
    return normalizeExcelPrecision(Math.abs(n));
  },

  SQRT: (args) => {
    const arg = args[0];
    if (isFormulaError(arg)) return arg;
    const n = coerceExcelNumber(arg);
    if (isFormulaError(n) || n < 0) return isFormulaError(n) ? n : createFormulaError('#NUM!', 'Negative number in SQRT');
    return normalizeExcelPrecision(Math.sqrt(n));
  },

  POWER: (args) => {
    const base = coerceExcelNumber(args[0]);
    const exp = coerceExcelNumber(args[1]);
    if (isFormulaError(base) || isFormulaError(exp)) return isFormulaError(base) ? base : exp;
    return normalizeExcelPrecision(Math.pow(base, exp));
  },

  MOD: (args) => {
    const n = coerceExcelNumber(args[0]);
    const d = coerceExcelNumber(args[1]);
    if (isFormulaError(n) || isFormulaError(d)) return isFormulaError(n) ? n : d;
    if (d === 0) return createFormulaError('#DIV/0!', 'Division by zero in MOD');
    return normalizeExcelPrecision(((n % d) + d) % d);
  },

  ROUND: (args) => {
    const n = coerceExcelNumber(args[0]);
    const digits = args[1] !== undefined ? coerceExcelNumber(args[1]) : 0;
    if (isFormulaError(n) || isFormulaError(digits)) return isFormulaError(n) ? n : digits;
    return roundExcel(n, digits);
  },

  ROUNDUP: (args) => {
    const n = coerceExcelNumber(args[0]);
    const digits = args[1] !== undefined ? coerceExcelNumber(args[1]) : 0;
    if (isFormulaError(n) || isFormulaError(digits)) return isFormulaError(n) ? n : digits;
    return roundExcelUp(n, digits);
  },

  ROUNDDOWN: (args) => {
    const n = coerceExcelNumber(args[0]);
    const digits = args[1] !== undefined ? coerceExcelNumber(args[1]) : 0;
    if (isFormulaError(n) || isFormulaError(digits)) return isFormulaError(n) ? n : digits;
    return roundExcelDown(n, digits);
  },

  INT: (args) => {
    const n = coerceExcelNumber(args[0]);
    if (isFormulaError(n)) return n;
    return normalizeExcelPrecision(Math.floor(n));
  },

  TRUNC: (args) => {
    const n = coerceExcelNumber(args[0]);
    const digits = args[1] !== undefined ? coerceExcelNumber(args[1]) : 0;
    if (isFormulaError(n) || isFormulaError(digits)) return isFormulaError(n) ? n : digits;
    return truncateExcel(n, digits);
  },

  CEILING: (args) => {
    const n = coerceExcelNumber(args[0]);
    const sig = args[1] !== undefined ? coerceExcelNumber(args[1]) : 1;
    if (isFormulaError(n) || isFormulaError(sig)) return isFormulaError(n) ? n : sig;
    if (sig === 0) return 0;
    return normalizeExcelPrecision(Math.ceil(n / sig) * sig);
  },

  FLOOR: (args) => {
    const n = coerceExcelNumber(args[0]);
    const sig = args[1] !== undefined ? coerceExcelNumber(args[1]) : 1;
    if (isFormulaError(n) || isFormulaError(sig)) return isFormulaError(n) ? n : sig;
    if (sig === 0) return 0;
    return normalizeExcelPrecision(Math.floor(n / sig) * sig);
  },

  PI: () => Math.PI,

  RAND: (_args, context) => {
    const value = context?.random?.('RAND', context?.volatileOccurrence);
    return value ?? createFormulaError('#BLOCKED!', 'RAND requires a calculation entropy context');
  },

  RANDBETWEEN: (args, context) => {
    const minArg = coerceExcelNumber(args[0]);
    const maxArg = coerceExcelNumber(args[1]);
    if (isFormulaError(minArg) || isFormulaError(maxArg)) return isFormulaError(minArg) ? minArg : maxArg;
    const min = Math.ceil(minArg);
    const max = Math.floor(maxArg);
    if (min > max) return createFormulaError('#NUM!', 'Invalid range in RANDBETWEEN');
    const random = context?.random?.('RANDBETWEEN', context?.volatileOccurrence);
    if (random === undefined || isFormulaError(random)) return random ?? createFormulaError('#BLOCKED!', 'RANDBETWEEN requires a calculation entropy context');
    return Math.floor(random * (max - min + 1)) + min;
  },

  EXP: (args) => {
    const n = coerceExcelNumber(args[0]);
    if (isFormulaError(n)) return n;
    return normalizeExcelPrecision(Math.exp(n));
  },

  LN: (args) => {
    const n = coerceExcelNumber(args[0]);
    if (isFormulaError(n) || n <= 0) return isFormulaError(n) ? n : createFormulaError('#NUM!', 'Number must be positive in LN');
    return normalizeExcelPrecision(Math.log(n));
  },

  LOG: (args) => {
    const n = coerceExcelNumber(args[0]);
    const base = args[1] !== undefined ? coerceExcelNumber(args[1]) : 10;
    if (isFormulaError(n) || isFormulaError(base)) return isFormulaError(n) ? n : base;
    if (n <= 0 || base <= 0 || base === 1) {
      return createFormulaError('#NUM!', 'Invalid base or number in LOG');
    }
    return normalizeExcelPrecision(Math.log(n) / Math.log(base));
  },

  LOG10: (args) => {
    const n = coerceExcelNumber(args[0]);
    if (isFormulaError(n) || n <= 0) return isFormulaError(n) ? n : createFormulaError('#NUM!', 'Number must be positive in LOG10');
    return normalizeExcelPrecision(Math.log10(n));
  },

  SIGN: (args) => {
    const n = coerceExcelNumber(args[0]);
    if (isFormulaError(n)) return n;
    return Math.sign(n);
  },
};
