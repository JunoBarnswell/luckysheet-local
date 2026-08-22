import { createFormulaError, type FormulaValue } from '../values';

function to2DArray(val: FormulaValue | undefined): FormulaValue[][] {
  if (val === undefined || val === null) return [[]];
  if (Array.isArray(val)) {
    if (val.length === 0) return [[]];
    if (Array.isArray(val[0])) return val as FormulaValue[][];
    return [val as FormulaValue[]];
  }
  return [[val]];
}

function to1DArray(val: FormulaValue | undefined): FormulaValue[] {
  if (val === undefined || val === null) return [];
  if (Array.isArray(val)) {
    const list: FormulaValue[] = [];
    for (const row of val) {
      if (Array.isArray(row)) {
        for (const cell of row) list.push(cell);
      } else {
        list.push(row);
      }
    }
    return list;
  }
  return [val];
}

export const lookupFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  VLOOKUP: (args) => {
    const lookupValue = args[0];
    const table = to2DArray(args[1]);
    const colIndex = Number(args[2]);
    const exactMatch = args[3] !== undefined ? Boolean(args[3]) === false || String(args[3]) === '0' || String(args[3]).toLowerCase() === 'false' : true;

    if (Number.isNaN(colIndex) || colIndex < 1 || colIndex > (table[0]?.length ?? 0)) {
      return createFormulaError('#REF!', 'Column index out of bounds in VLOOKUP');
    }

    const lookupStr = String(lookupValue ?? '').toLowerCase();
    for (const row of table) {
      const firstCell = row[0];
      if (firstCell === undefined) continue;
      const cellStr = String(firstCell ?? '').toLowerCase();
      if (exactMatch) {
        if (cellStr === lookupStr) {
          return row[colIndex - 1] ?? null;
        }
      } else {
        if (cellStr === lookupStr) {
          return row[colIndex - 1] ?? null;
        }
      }
    }
    return createFormulaError('#N/A', 'Value not found in VLOOKUP');
  },

  HLOOKUP: (args) => {
    const lookupValue = args[0];
    const table = to2DArray(args[1]);
    const rowIndex = Number(args[2]);
    const exactMatch = args[3] !== undefined ? Boolean(args[3]) === false || String(args[3]) === '0' || String(args[3]).toLowerCase() === 'false' : true;

    if (Number.isNaN(rowIndex) || rowIndex < 1 || rowIndex > table.length) {
      return createFormulaError('#REF!', 'Row index out of bounds in HLOOKUP');
    }

    const firstRow = table[0] ?? [];
    const lookupStr = String(lookupValue ?? '').toLowerCase();
    for (let c = 0; c < firstRow.length; c++) {
      const cellStr = String(firstRow[c] ?? '').toLowerCase();
      if (exactMatch) {
        if (cellStr === lookupStr) {
          return table[rowIndex - 1]?.[c] ?? null;
        }
      } else {
        if (cellStr === lookupStr) {
          return table[rowIndex - 1]?.[c] ?? null;
        }
      }
    }
    return createFormulaError('#N/A', 'Value not found in HLOOKUP');
  },

  INDEX: (args) => {
    const table = to2DArray(args[0]);
    const rowNum = args[1] !== undefined ? Number(args[1]) : 1;
    const colNum = args[2] !== undefined ? Number(args[2]) : 1;

    if (Number.isNaN(rowNum) || Number.isNaN(colNum) || rowNum < 0 || colNum < 0) {
      return createFormulaError('#VALUE!', 'Invalid index in INDEX');
    }
    if (rowNum === 0 && colNum === 0) return table;
    if (rowNum === 0) {
      // Return whole column
      return table.map((r) => [r[colNum - 1] ?? null]);
    }
    if (colNum === 0) {
      // Return whole row
      return [table[rowNum - 1] ?? []];
    }
    const row = table[rowNum - 1];
    if (!row) return createFormulaError('#REF!', 'Row out of bounds in INDEX');
    const cell = row[colNum - 1];
    return cell !== undefined ? cell : createFormulaError('#REF!', 'Column out of bounds in INDEX');
  },

  MATCH: (args) => {
    const lookupValue = args[0];
    const array = to1DArray(args[1]);
    const matchType = args[2] !== undefined ? Number(args[2]) : 1;

    const lookupStr = String(lookupValue ?? '').toLowerCase();
    for (let i = 0; i < array.length; i++) {
      const cellStr = String(array[i] ?? '').toLowerCase();
      if (matchType === 0) {
        if (cellStr === lookupStr) return i + 1;
      } else {
        if (cellStr === lookupStr) return i + 1;
      }
    }
    return createFormulaError('#N/A', 'Value not found in MATCH');
  },

  XLOOKUP: (args) => {
    const lookupValue = args[0];
    const lookupArray = to1DArray(args[1]);
    const returnArray = to1DArray(args[2]);
    const ifNotFound = args[3] !== undefined ? args[3] : createFormulaError('#N/A', 'Value not found in XLOOKUP');

    const lookupStr = String(lookupValue ?? '').toLowerCase();
    for (let i = 0; i < lookupArray.length; i++) {
      if (String(lookupArray[i] ?? '').toLowerCase() === lookupStr) {
        return returnArray[i] ?? null;
      }
    }
    return ifNotFound;
  },

  CHOOSE: (args) => {
    const index = Number(args[0]);
    if (Number.isNaN(index) || index < 1 || index >= args.length) {
      return createFormulaError('#VALUE!', 'Index out of bounds in CHOOSE');
    }
    return args[index] ?? null;
  },

  COLUMNS: (args) => {
    const table = to2DArray(args[0]);
    return table[0]?.length ?? 0;
  },

  ROWS: (args) => {
    const table = to2DArray(args[0]);
    return table.length;
  },

  TRANSPOSE: (args) => {
    const table = to2DArray(args[0]);
    if (table.length === 0 || !table[0]) return [[]];
    const rowCount = table.length;
    const colCount = table[0].length;
    const result: FormulaValue[][] = [];

    for (let c = 0; c < colCount; c++) {
      const newRow: FormulaValue[] = [];
      for (let r = 0; r < rowCount; r++) {
        newRow.push(table[r]?.[c] ?? null);
      }
      result.push(newRow);
    }
    return result;
  },
};
