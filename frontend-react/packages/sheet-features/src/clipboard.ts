import type { CellData, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import { formatFormula, offsetAst, parseFormula } from '@react-sheets/formula-engine';

/** Stable MIME names understood by every host (browser, desktop and headless). */
export const CLIPBOARD_INTERNAL_MIME = 'application/x-react-sheets-cells';
export const CLIPBOARD_HTML_MIME = 'text/html';
export const CLIPBOARD_TEXT_MIME = 'text/plain';

export interface ClipboardRepresentation {
  mime: string;
  data: string;
}

/**
 * Host-neutral clipboard contract.  A host may expose any subset of the
 * representations, but the command layer always consumes the normalized
 * matrix.  `representations` is intentionally a data-only value so this
 * contract can be serialized for desktop and headless hosts without a DOM or
 * ClipboardItem dependency.
 */
export interface ClipboardPayload {
  range: RangeRef;
  values: CellData[][];
  isCut?: boolean;
  representations?: ClipboardRepresentation[];
  mime?: string;
  html?: string;
  text?: string;
  source?: string;
}

// Temporary type-only compatibility for existing host call sites. It has no
// runtime path or data conversion; the host migration can remove it once all
// consumers use ClipboardPayload directly.
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

  const internal = JSON.stringify({ range, values });
  const text = formatTsv(values);
  const html = formatHtml(values);
  return {
    range: structuredClone(range),
    values,
    representations: [
      { mime: CLIPBOARD_INTERNAL_MIME, data: internal },
      { mime: CLIPBOARD_HTML_MIME, data: html },
      { mime: CLIPBOARD_TEXT_MIME, data: text },
    ],
    mime: CLIPBOARD_INTERNAL_MIME,
    html,
    text,
    source: 'internal',
  };
}

export function formatTsv(values: CellData[][]): string {
  return values
    .map((row) =>
      row
        .map((cell) => {
          if (cell.value === null || cell.value === undefined) return '';
          return escapeTsvField(String(cell.value));
        })
        .join('\t'),
    )
    .join('\n');
}

export function parseTsv(text: string): CellData[][] {
  if (text.length === 0) return [];
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]!;
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"' && field.length === 0) {
      quoted = true;
    } else if (char === '\t') {
      row.push(field);
      field = '';
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && text[index + 1] === '\n') index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    if (row.length > 0 && !(row.length === 1 && row[0] === '' && rows.length > 0)) rows.push(row);
  }
  return rows.map((values) => values.map(parseClipboardScalar));
}

/** Parse a serialized clipboard payload supplied by a host ClipboardItem. */
export function parseClipboardPayload(payload: ClipboardPayload): CellData[][] {
  const internal = payload.representations?.find((entry) => entry.mime === CLIPBOARD_INTERNAL_MIME)?.data;
  if (internal) {
    try {
      const parsed = JSON.parse(internal) as { values?: CellData[][] };
      if (Array.isArray(parsed.values)) return parsed.values.map((row) => row.map((cell) => structuredClone(cell)));
    } catch {
      // Continue to the next lossless representation rather than failing a
      // paste because a host supplied malformed optional metadata.
    }
  }
  if (payload.values.length > 0) return payload.values.map((row) => row.map((cell) => structuredClone(cell)));
  const html = payload.representations?.find((entry) => entry.mime === CLIPBOARD_HTML_MIME)?.data ?? payload.html;
  if (html) return parseHtmlTable(html);
  const text = payload.representations?.find((entry) => entry.mime === CLIPBOARD_TEXT_MIME)?.data ?? payload.text;
  return text ? parseTsv(text) : [];
}

function escapeTsvField(value: string): string {
  return /[\t\r\n"]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
}

function parseClipboardScalar(value: string): CellData {
  if (value === '') return { value: null };
  const trimmed = value.trim();
  if (trimmed.startsWith('=')) return { value: null, formula: trimmed };
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    const number = Number(trimmed);
    if (Number.isFinite(number)) return { value: number };
  }
  if (trimmed.toUpperCase() === 'TRUE') return { value: true };
  if (trimmed.toUpperCase() === 'FALSE') return { value: false };
  return { value };
}

function formatHtml(values: CellData[][]): string {
  const rows = values.map((row) => `<tr>${row.map((cell) => {
    const value = cell.value == null ? '' : escapeHtml(String(cell.value));
    const formula = cell.formula ? ` data-formula="${escapeHtml(cell.formula)}"` : '';
    return `<td${formula}>${value}</td>`;
  }).join('')}</tr>`).join('');
  return `<table>${rows}</table>`;
}

function parseHtmlTable(html: string): CellData[][] {
  const rows: CellData[][] = [];
  const rowMatches = html.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) ?? [];
  for (const rowHtml of rowMatches) {
    const cells: CellData[] = [];
    const cellMatches = rowHtml.match(/<(?:td|th)\b([^>]*)>[\s\S]*?<\/(?:td|th)>/gi) ?? [];
    for (const cellHtml of cellMatches) {
      const attributes = cellHtml.match(/^<(?:td|th)\b([^>]*)>/i)?.[1] ?? '';
      const formula = attributes.match(/\bdata-formula\s*=\s*["']([^"']*)["']/i)?.[1];
      const body = cellHtml.replace(/^<(?:td|th)\b[^>]*>/i, '').replace(/<\/(?:td|th)>$/i, '');
      const text = decodeHtml(body.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ''));
      const parsed = parseClipboardScalar(text);
      if (formula) {
        parsed.value = null;
        parsed.formula = decodeHtml(formula);
      }
      cells.push(parsed);
    }
    rows.push(cells);
  }
  return rows;
}

function escapeHtml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

function decodeHtml(value: string): string {
  return value
    .replaceAll('&nbsp;', ' ')
    .replaceAll('&quot;', '"')
    .replaceAll('&#39;', "'")
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
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
