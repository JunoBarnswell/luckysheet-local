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

function logicalAggregate(name: 'AND' | 'OR' | 'XOR', args: FormulaValue[]): FormulaValue {
  let found = false;
  let truthCount = 0;
  let falseCount = 0;
  for (const arg of args) {
    const reference = Array.isArray(arg);
    const values = reference ? arg.flat() : [arg];
    for (const value of values) {
      if (isFormulaError(value)) return value;
      if (reference && (value === null || typeof value === 'string')) continue;
      if (typeof value === 'string' && !['TRUE', 'FALSE'].includes(value.toUpperCase())) return createFormulaError('#VALUE!', 'Logical argument is not valid');
      found = true;
      if (toBoolean(value)) truthCount++; else falseCount++;
    }
  }
  if (!found) return createFormulaError('#VALUE!', 'Logical function has no logical values');
  return name === 'AND' ? falseCount === 0 : name === 'OR' ? truthCount > 0 : truthCount % 2 === 1;
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

  AND: (args) => logicalAggregate('AND', args),
  OR: (args) => logicalAggregate('OR', args),
  XOR: (args) => logicalAggregate('XOR', args),
  NOT: (args) => {
    const value = args[0];
    if (isFormulaError(value)) return value;
    if (typeof value === 'string' && !['TRUE', 'FALSE'].includes(value.toUpperCase())) return createFormulaError('#VALUE!', 'NOT requires a logical value');
    return !toBoolean(value);
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
