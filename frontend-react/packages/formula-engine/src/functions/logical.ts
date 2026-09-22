import { createFormulaError, isFormulaError, type FormulaValue } from '../values';

function toBoolean(val: FormulaValue | undefined): boolean {
  if (val === undefined || val === null) return false;
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (typeof val === 'string') {
    const s = val.trim().toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    return s.length > 0;
  }
  return false;
}

export const logicalFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  IF: (args) => {
    const condition = args[0] ?? null;
    if (isFormulaError(condition)) return condition;
    const isTrue = toBoolean(condition);
    if (isTrue) {
      return args[1] !== undefined ? args[1] : true;
    }
    return args[2] !== undefined ? args[2] : false;
  },

  IFS: (args) => {
    if (args.length % 2 !== 0) return createFormulaError('#VALUE!', 'IFS requires pairs of condition and value');
    for (let i = 0; i < args.length; i += 2) {
      const cond = args[i] ?? null;
      if (isFormulaError(cond)) return cond;
      if (toBoolean(cond)) {
        return args[i + 1] ?? null;
      }
    }
    return createFormulaError('#N/A', 'No condition matched in IFS');
  },

  IFERROR: (args) => {
    const val = args[0] ?? null;
    if (isFormulaError(val)) {
      return args[1] !== undefined ? args[1] : '';
    }
    return val;
  },

  IFNA: (args) => {
    const val = args[0] ?? null;
    if (isFormulaError(val) && val.code === '#N/A') {
      return args[1] !== undefined ? args[1] : '';
    }
    return val;
  },

  AND: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'AND requires arguments');
    for (const arg of args) {
      if (isFormulaError(arg)) return arg;
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) {
              if (isFormulaError(cell)) return cell;
              if (!toBoolean(cell)) return false;
            }
          } else if (!toBoolean(row)) return false;
        }
      } else if (!toBoolean(arg)) {
        return false;
      }
    }
    return true;
  },

  OR: (args) => {
    if (args.length === 0) return createFormulaError('#VALUE!', 'OR requires arguments');
    for (const arg of args) {
      if (isFormulaError(arg)) return arg;
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) {
              if (isFormulaError(cell)) return cell;
              if (toBoolean(cell)) return true;
            }
          } else if (toBoolean(row)) return true;
        }
      } else if (toBoolean(arg)) {
        return true;
      }
    }
    return false;
  },

  NOT: (args) => {
    const arg = args[0] ?? null;
    if (isFormulaError(arg)) return arg;
    return !toBoolean(arg);
  },

  XOR: (args) => {
    let trueCount = 0;
    for (const arg of args) {
      if (isFormulaError(arg)) return arg;
      if (toBoolean(arg)) trueCount += 1;
    }
    return trueCount % 2 === 1;
  },

  SWITCH: (args) => {
    if (args.length < 3) return createFormulaError('#VALUE!', 'SWITCH requires target, value, result');
    const target = args[0] ?? null;
    if (isFormulaError(target)) return target;

    for (let i = 1; i < args.length - 1; i += 2) {
      const matchVal = args[i];
      if (String(target) === String(matchVal)) {
        return args[i + 1] ?? null;
      }
    }
    // Default value if odd arguments count
    if (args.length % 2 === 0) {
      return args[args.length - 1] ?? null;
    }
    return createFormulaError('#N/A', 'No matching case in SWITCH');
  },

  TRUE: () => true,
  FALSE: () => false,
};
