import { createFormulaError, isFormulaError, type FormulaValue } from '../values';
import { formatValue } from '@react-sheets/number-format';

function toStringVal(val: FormulaValue | undefined): string {
  if (val === null || val === undefined) return '';
  if (isFormulaError(val)) return val.code;
  return String(val);
}

export const textFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  CONCAT: (args) => {
    let result = '';
    for (const arg of args) {
      if (isFormulaError(arg)) return arg;
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) result += toStringVal(cell);
          } else {
            result += toStringVal(row);
          }
        }
      } else {
        result += toStringVal(arg);
      }
    }
    return result;
  },

  CONCATENATE: (args) => {
    let result = '';
    for (const arg of args) {
      if (isFormulaError(arg)) return arg;
      result += toStringVal(arg);
    }
    return result;
  },

  TEXTJOIN: (args) => {
    const delimiter = toStringVal(args[0]);
    const ignoreEmpty = Boolean(args[1]);
    const items: string[] = [];

    for (let i = 2; i < args.length; i++) {
      const arg = args[i];
      if (isFormulaError(arg)) return arg;
      if (Array.isArray(arg)) {
        for (const row of arg) {
          if (Array.isArray(row)) {
            for (const cell of row) {
              const s = toStringVal(cell);
              if (!ignoreEmpty || s.length > 0) items.push(s);
            }
          } else {
            const s = toStringVal(row);
            if (!ignoreEmpty || s.length > 0) items.push(s);
          }
        }
      } else {
        const s = toStringVal(arg);
        if (!ignoreEmpty || s.length > 0) items.push(s);
      }
    }
    return items.join(delimiter);
  },

  LEFT: (args) => {
    const str = toStringVal(args[0]);
    const numChars = args[1] !== undefined ? Number(args[1]) : 1;
    if (Number.isNaN(numChars) || numChars < 0) return createFormulaError('#VALUE!', 'Invalid length in LEFT');
    return str.slice(0, numChars);
  },

  RIGHT: (args) => {
    const str = toStringVal(args[0]);
    const numChars = args[1] !== undefined ? Number(args[1]) : 1;
    if (Number.isNaN(numChars) || numChars < 0) return createFormulaError('#VALUE!', 'Invalid length in RIGHT');
    return str.slice(Math.max(0, str.length - numChars));
  },

  MID: (args) => {
    const str = toStringVal(args[0]);
    const startNum = Number(args[1]);
    const numChars = Number(args[2]);
    if (Number.isNaN(startNum) || Number.isNaN(numChars) || startNum < 1 || numChars < 0) {
      return createFormulaError('#VALUE!', 'Invalid start or length in MID');
    }
    return str.slice(startNum - 1, startNum - 1 + numChars);
  },

  LEN: (args) => {
    const str = toStringVal(args[0]);
    return str.length;
  },

  LOWER: (args) => {
    return toStringVal(args[0]).toLowerCase();
  },

  UPPER: (args) => {
    return toStringVal(args[0]).toUpperCase();
  },

  PROPER: (args) => {
    const str = toStringVal(args[0]);
    return str.replace(/\b\w/g, (c) => c.toUpperCase());
  },

  TRIM: (args) => {
    return toStringVal(args[0]).trim().replace(/\s+/g, ' ');
  },

  CLEAN: (args) => {
    return toStringVal(args[0]).replace(/[\x00-\x1F\x7F]/g, '');
  },

  EXACT: (args) => {
    return toStringVal(args[0]) === toStringVal(args[1]);
  },

  FIND: (args) => {
    const findText = toStringVal(args[0]);
    const withinText = toStringVal(args[1]);
    const startNum = args[2] !== undefined ? Number(args[2]) : 1;
    if (Number.isNaN(startNum) || startNum < 1) return createFormulaError('#VALUE!', 'Invalid startNum in FIND');
    const idx = withinText.indexOf(findText, startNum - 1);
    return idx === -1 ? createFormulaError('#VALUE!', 'Text not found in FIND') : idx + 1;
  },

  SEARCH: (args) => {
    const findText = toStringVal(args[0]).toLowerCase();
    const withinText = toStringVal(args[1]).toLowerCase();
    const startNum = args[2] !== undefined ? Number(args[2]) : 1;
    if (Number.isNaN(startNum) || startNum < 1) return createFormulaError('#VALUE!', 'Invalid startNum in SEARCH');
    const idx = withinText.indexOf(findText, startNum - 1);
    return idx === -1 ? createFormulaError('#VALUE!', 'Text not found in SEARCH') : idx + 1;
  },

  REPLACE: (args) => {
    const oldText = toStringVal(args[0]);
    const startNum = Number(args[1]);
    const numChars = Number(args[2]);
    const newText = toStringVal(args[3]);
    if (Number.isNaN(startNum) || Number.isNaN(numChars) || startNum < 1 || numChars < 0) {
      return createFormulaError('#VALUE!', 'Invalid arguments in REPLACE');
    }
    return oldText.slice(0, startNum - 1) + newText + oldText.slice(startNum - 1 + numChars);
  },

  SUBSTITUTE: (args) => {
    const text = toStringVal(args[0]);
    const oldText = toStringVal(args[1]);
    const newText = toStringVal(args[2]);
    const instanceNum = args[3] !== undefined ? Number(args[3]) : undefined;

    if (!oldText) return text;
    if (instanceNum === undefined) {
      return text.split(oldText).join(newText);
    }
    let count = 0;
    return text.replace(new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), (match) => {
      count++;
      return count === instanceNum ? newText : match;
    });
  },

  REPT: (args) => {
    const text = toStringVal(args[0]);
    const count = Number(args[1]);
    if (Number.isNaN(count) || count < 0) return createFormulaError('#VALUE!', 'Invalid count in REPT');
    return text.repeat(Math.floor(count));
  },

  TEXT: (args) => {
    const val = args[0] ?? null;
    const format = toStringVal(args[1]);
    if (isFormulaError(val)) return val;
    if (typeof val === 'number') {
      return formatValue(val, format || undefined);
    }
    return toStringVal(val);
  },

  VALUE: (args) => {
    const str = toStringVal(args[0]).trim();
    const cleanStr = str.replace(/[$,%]/g, '');
    const num = Number(cleanStr);
    if (Number.isNaN(num)) return createFormulaError('#VALUE!', 'Cannot convert text to number in VALUE');
    if (str.endsWith('%')) return num / 100;
    return num;
  },

  CHAR: (args) => {
    const code = Number(args[0]);
    if (Number.isNaN(code) || code < 1 || code > 255) return createFormulaError('#VALUE!', 'Invalid char code');
    return String.fromCharCode(code);
  },

  CODE: (args) => {
    const str = toStringVal(args[0]);
    if (str.length === 0) return createFormulaError('#VALUE!', 'Empty string in CODE');
    return str.charCodeAt(0);
  },
};
