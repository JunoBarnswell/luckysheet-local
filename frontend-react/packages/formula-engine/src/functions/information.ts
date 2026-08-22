import { isFormulaError, type FormulaValue } from '../values';

export const informationFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  ISBLANK: (args) => {
    const val = args[0];
    return val === null || val === undefined || val === '';
  },

  ISNUMBER: (args) => {
    const val = args[0];
    return typeof val === 'number' && !Number.isNaN(val);
  },

  ISTEXT: (args) => {
    const val = args[0];
    return typeof val === 'string';
  },

  ISNONTEXT: (args) => {
    const val = args[0];
    return typeof val !== 'string';
  },

  ISLOGICAL: (args) => {
    const val = args[0];
    return typeof val === 'boolean';
  },

  ISERROR: (args) => {
    const val = args[0];
    return isFormulaError(val);
  },

  ISERR: (args) => {
    const val = args[0];
    return isFormulaError(val) && val.code !== '#N/A';
  },

  ISNA: (args) => {
    const val = args[0];
    return isFormulaError(val) && val.code === '#N/A';
  },

  N: (args) => {
    const val = args[0];
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    return 0;
  },

  T: (args) => {
    const val = args[0];
    return typeof val === 'string' ? val : '';
  },
};
