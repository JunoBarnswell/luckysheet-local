import type { CellData, RangeRef, WorkbookModel } from '@react-sheets/core-model';

export interface ClipboardData {
  range: RangeRef;
  values: CellData[][];
  isCut?: boolean;
}

export function copyRangeToClipboardData(workbook: WorkbookModel, range: RangeRef): ClipboardData {
  const sheet = workbook.getSheet(range.sheetId);
  const values: CellData[][] = [];

  for (let r = range.startRow; r <= range.endRow; r++) {
    const rowList: CellData[] = [];
    for (let c = range.startColumn; c <= range.endColumn; c++) {
      const cell = sheet.cells.get(r, c);
      rowList.push(cell ? structuredClone(cell) : { value: null });
    }
    values.push(rowList);
  }

  return { range: structuredClone(range), values };
}

export function formatTsv(values: CellData[][]): string {
  return values
    .map((row) =>
      row
        .map((cell) => {
          if (cell.value === null || cell.value === undefined) return '';
          return String(cell.value);
        })
        .join('\t'),
    )
    .join('\n');
}

export function parseTsv(text: string): CellData[][] {
  const lines = text.split(/\r?\n/);
  if (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();

  return lines.map((line) =>
    line.split('\t').map((val) => {
      const trimmed = val.trim();
      if (trimmed === '') return { value: null };
      if (trimmed.startsWith('=')) return { value: null, formula: trimmed };
      const num = Number(trimmed);
      if (!Number.isNaN(num)) return { value: num };
      if (trimmed.toUpperCase() === 'TRUE') return { value: true };
      if (trimmed.toUpperCase() === 'FALSE') return { value: false };
      return { value: val };
    }),
  );
}

export function shiftFormula(formula: string, rowOffset: number, colOffset: number): string {
  if (!formula.startsWith('=')) return formula;

  // Regex to match cell references like A1, $A$1, A$1, $A1, Sheet1!A1
  return formula.replace(/(\$?[A-Z]+)(\$?\d+)/g, (match, colPart: string, rowPart: string) => {
    const isAbsCol = colPart.startsWith('$');
    const isAbsRow = rowPart.startsWith('$');

    let colStr = isAbsCol ? colPart.slice(1) : colPart;
    let rowNum = parseInt(isAbsRow ? rowPart.slice(1) : rowPart, 10);

    if (!isAbsCol && colOffset !== 0) {
      let colIdx = 0;
      for (const char of colStr) colIdx = colIdx * 26 + char.charCodeAt(0) - 64;
      colIdx = Math.max(1, colIdx + colOffset);
      let newColStr = '';
      while (colIdx > 0) {
        const rem = (colIdx - 1) % 26;
        newColStr = String.fromCharCode(65 + rem) + newColStr;
        colIdx = Math.floor((colIdx - 1) / 26);
      }
      colStr = newColStr;
    }

    if (!isAbsRow && rowOffset !== 0) {
      rowNum = Math.max(1, rowNum + rowOffset);
    }

    return `${isAbsCol ? '$' : ''}${colStr}${isAbsRow ? '$' : ''}${rowNum}`;
  });
}
