import { createFormulaError, isFormulaError, type FormulaValue } from '../values';

export function flattenNumericArgs(args: FormulaValue[]): number[] | ReturnType<typeof createFormulaError> {
  const numbers: number[] = [];
  for (const arg of args) {
    if (isFormulaError(arg)) return arg;
    if (Array.isArray(arg)) {
      for (const row of arg) {
        if (Array.isArray(row)) {
          for (const cell of row) {
            if (isFormulaError(cell)) return cell;
            if (typeof cell === 'number' && !Number.isNaN(cell)) numbers.push(cell);
          }
        } else if (typeof row === 'number' && !Number.isNaN(row)) {
          numbers.push(row);
        }
      }
    } else if (typeof arg === 'number' && !Number.isNaN(arg)) {
      numbers.push(arg);
    } else if (typeof arg === 'string') {
      const parsed = Number(arg);
      if (!Number.isNaN(parsed)) numbers.push(parsed);
    } else if (typeof arg === 'boolean') {
      numbers.push(arg ? 1 : 0);
    }
  }
  return numbers;
}

export const mathFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  SUM: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    return nums.reduce((acc, n) => acc + n, 0);
  },

  PRODUCT: (args) => {
    const nums = flattenNumericArgs(args);
    if (isFormulaError(nums)) return nums;
    if (nums.length === 0) return 0;
    return nums.reduce((acc, n) => acc * n, 1);
  },

  ABS: (args) => {
    const arg = args[0];
    if (isFormulaError(arg)) return arg;
    const n = Number(arg);
    if (Number.isNaN(n)) return createFormulaError('#VALUE!', 'Expected a number');
    return Math.abs(n);
  },

  SQRT: (args) => {
    const arg = args[0];
    if (isFormulaError(arg)) return arg;
    const n = Number(arg);
    if (Number.isNaN(n) || n < 0) return createFormulaError('#NUM!', 'Negative number or NaN in SQRT');
    return Math.sqrt(n);
  },

  POWER: (args) => {
    const base = Number(args[0]);
    const exp = Number(args[1]);
    if (Number.isNaN(base) || Number.isNaN(exp)) return createFormulaError('#VALUE!', 'Expected numbers');
    return Math.pow(base, exp);
  },

  MOD: (args) => {
    const n = Number(args[0]);
    const d = Number(args[1]);
    if (Number.isNaN(n) || Number.isNaN(d)) return createFormulaError('#VALUE!', 'Expected numbers');
    if (d === 0) return createFormulaError('#DIV/0!', 'Division by zero in MOD');
    return ((n % d) + d) % d;
  },

  ROUND: (args) => {
    const n = Number(args[0]);
    const digits = args[1] !== undefined ? Number(args[1]) : 0;
    if (Number.isNaN(n) || Number.isNaN(digits)) return createFormulaError('#VALUE!', 'Expected numbers');
    const factor = Math.pow(10, digits);
    return Math.round(n * factor) / factor;
  },

  ROUNDUP: (args) => {
    const n = Number(args[0]);
    const digits = args[1] !== undefined ? Number(args[1]) : 0;
    if (Number.isNaN(n) || Number.isNaN(digits)) return createFormulaError('#VALUE!', 'Expected numbers');
    const factor = Math.pow(10, digits);
    return (n >= 0 ? Math.ceil(n * factor) : Math.floor(n * factor)) / factor;
  },

  ROUNDDOWN: (args) => {
    const n = Number(args[0]);
    const digits = args[1] !== undefined ? Number(args[1]) : 0;
    if (Number.isNaN(n) || Number.isNaN(digits)) return createFormulaError('#VALUE!', 'Expected numbers');
    const factor = Math.pow(10, digits);
    return (n >= 0 ? Math.floor(n * factor) : Math.ceil(n * factor)) / factor;
  },

  INT: (args) => {
    const n = Number(args[0]);
    if (Number.isNaN(n)) return createFormulaError('#VALUE!', 'Expected a number');
    return Math.floor(n);
  },

  TRUNC: (args) => {
    const n = Number(args[0]);
    const digits = args[1] !== undefined ? Number(args[1]) : 0;
    if (Number.isNaN(n) || Number.isNaN(digits)) return createFormulaError('#VALUE!', 'Expected numbers');
    const factor = Math.pow(10, digits);
    return Math.trunc(n * factor) / factor;
  },

  CEILING: (args) => {
    const n = Number(args[0]);
    const sig = args[1] !== undefined ? Number(args[1]) : 1;
    if (Number.isNaN(n) || Number.isNaN(sig)) return createFormulaError('#VALUE!', 'Expected numbers');
    if (sig === 0) return 0;
    return Math.ceil(n / sig) * sig;
  },

  FLOOR: (args) => {
    const n = Number(args[0]);
    const sig = args[1] !== undefined ? Number(args[1]) : 1;
    if (Number.isNaN(n) || Number.isNaN(sig)) return createFormulaError('#VALUE!', 'Expected numbers');
    if (sig === 0) return 0;
    return Math.floor(n / sig) * sig;
  },

  PI: () => Math.PI,

  RAND: () => Math.random(),

  RANDBETWEEN: (args) => {
    const min = Math.ceil(Number(args[0]));
    const max = Math.floor(Number(args[1]));
    if (Number.isNaN(min) || Number.isNaN(max) || min > max) return createFormulaError('#NUM!', 'Invalid range in RANDBETWEEN');
    return Math.floor(Math.random() * (max - min + 1)) + min;
  },

  EXP: (args) => {
    const n = Number(args[0]);
    if (Number.isNaN(n)) return createFormulaError('#VALUE!', 'Expected a number');
    return Math.exp(n);
  },

  LN: (args) => {
    const n = Number(args[0]);
    if (Number.isNaN(n) || n <= 0) return createFormulaError('#NUM!', 'Number must be positive in LN');
    return Math.log(n);
  },

  LOG: (args) => {
    const n = Number(args[0]);
    const base = args[1] !== undefined ? Number(args[1]) : 10;
    if (Number.isNaN(n) || Number.isNaN(base) || n <= 0 || base <= 0 || base === 1) {
      return createFormulaError('#NUM!', 'Invalid base or number in LOG');
    }
    return Math.log(n) / Math.log(base);
  },

  LOG10: (args) => {
    const n = Number(args[0]);
    if (Number.isNaN(n) || n <= 0) return createFormulaError('#NUM!', 'Number must be positive in LOG10');
    return Math.log10(n);
  },

  SIGN: (args) => {
    const n = Number(args[0]);
    if (Number.isNaN(n)) return createFormulaError('#VALUE!', 'Expected a number');
    return Math.sign(n);
  },
};
