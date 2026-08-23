import type { CellData, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import { formatFormula, offsetAst, parseFormula } from '@react-sheets/formula-engine';

/**
 * Host-neutral clipboard contract. Browser integrations may add MIME/HTML
 * representations, while the command layer always consumes the normalized
 * cell matrix. `source` identifies provenance for policy/audit decisions; it
 * is intentionally open-ended so native hosts do not need an adapter type.
 */
export interface ClipboardPayload {
  range: RangeRef;
  values: CellData[][];
  isCut?: boolean;
  mime?: string;
  html?: string;
  text?: string;
  source?: string;
}

/** @deprecated Use ClipboardPayload. Kept as a source-compatible type alias while hosts migrate. */
export type ClipboardData = ClipboardPayload;

export function copyRangeToClipboardData(workbook: WorkbookModel, range: RangeRef): ClipboardPayload {
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

  return {
    range: structuredClone(range),
    values,
    mime: 'application/x-react-sheets-cells',
    source: 'internal',
  };
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

/**
 * Shift relative references for copy/fill/paste using the formula AST. A
 * formula that is not understood by the parser is preserved verbatim; it is
 * never rewritten by a lossy regular expression.
 */
export function shiftFormula(formula: string, rowOffset: number, colOffset: number): string {
  if (!formula.trim().startsWith('=')) return formula;
  try {
    return formatFormula(offsetAst(parseFormula(formula), rowOffset, colOffset));
  } catch {
    return formula;
  }
}
