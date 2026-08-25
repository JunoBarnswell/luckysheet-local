import type {
  CellData,
  CellHyperlink,
  CellNote,
  CommentThread,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorkbookModel,
} from '@react-sheets/core-model';
import { formatFormula, offsetAst, parseFormula } from '@react-sheets/formula-engine';
import { parseCellText } from './text-input';

/** Stable MIME names understood by every host (browser, desktop and headless). */
export const CLIPBOARD_INTERNAL_MIME = 'application/x-react-sheets-cells';
export const CLIPBOARD_HTML_MIME = 'text/html';
export const CLIPBOARD_TEXT_MIME = 'text/plain';

export interface ClipboardRepresentation {
  mime: string;
  data: string;
}

export type ClipboardTransfer = 'copy' | 'move';

/**
 * Canonical Paste Special contract.  The fields are deliberately independent:
 * callers must not encode a combination as a string mode because content,
 * formatting, range-owned metadata and arithmetic are separate Excel
 * semantics.
 */
export type PasteContent = 'none' | 'all' | 'values' | 'formulas';
export type PasteFormatting =
  | 'all'
  | 'none'
  | 'number-format'
  | 'source-formatting'
  | 'all-except-borders'
  | 'source-theme';
export type PasteArithmetic = 'none' | 'add' | 'subtract' | 'multiply' | 'divide';

export interface PasteMetadataSpec {
  commentsNotes: boolean;
  validation: boolean;
  columnWidths: boolean;
  conditionalFormats: boolean;
  hyperlinks: boolean;
}

export interface PasteSpecialSpec {
  content: PasteContent;
  formatting: PasteFormatting;
  metadata: PasteMetadataSpec;
  operation: PasteArithmetic;
  skipBlanks: boolean;
  transpose: boolean;
  /** Paste Link is represented as formulas pointing at the source range. */
  link: boolean;
}

export const DEFAULT_PASTE_SPECIAL_SPEC: PasteSpecialSpec = {
  content: 'all',
  formatting: 'all',
  metadata: {
    commentsNotes: true,
    validation: true,
    columnWidths: false,
    conditionalFormats: true,
    hyperlinks: true,
  },
  operation: 'none',
  skipBlanks: false,
  transpose: false,
  link: false,
};

export function createPasteSpecialSpec(overrides: Partial<PasteSpecialSpec> = {}): PasteSpecialSpec {
  return {
    ...structuredClone(DEFAULT_PASTE_SPECIAL_SPEC),
    ...overrides,
    metadata: {
      ...structuredClone(DEFAULT_PASTE_SPECIAL_SPEC.metadata),
      ...(overrides.metadata ?? {}),
    },
  };
}

/** Canonical capability gate used by both the planner and Paste Special UI. */
export function isPasteSpecialSpecSupported(spec: PasteSpecialSpec, _payload?: ClipboardPayload): boolean {
  // The workbook model has no editable theme owner yet. Keep the contract
  // visible, but reject it before planning instead of reporting fake success.
  if (spec.formatting === 'source-theme') return false;
  if (spec.link && spec.operation !== 'none') return false;
  if (spec.content === 'none' && spec.formatting === 'none' && spec.operation === 'none' && !spec.link
    && !Object.values(spec.metadata).some(Boolean)) return false;
  return true;
}

export interface ClipboardColumnWidth {
  offset: number;
  widthPx: number;
}

export interface ClipboardCellMetadata<T> {
  rowOffset: number;
  columnOffset: number;
  value: T;
}

/** Range-owned data is carried once, outside the cell matrix. */
export interface ClipboardRangeMetadata {
  columnWidths: ClipboardColumnWidth[];
  validations: DataValidationRule[];
  conditionalFormats: ConditionalFormatRule[];
  notes: Array<ClipboardCellMetadata<CellNote>>;
  comments: Array<ClipboardCellMetadata<CommentThread>>;
  hyperlinks: Array<ClipboardCellMetadata<CellHyperlink>>;
  themeIdentity?: string;
}

export class FormulaRelocationError extends Error {
  readonly code = 'FORMULA_RELOCATION_FAILED';

  constructor(
    readonly formula: string,
    readonly rowOffset: number,
    readonly columnOffset: number,
    cause: unknown,
  ) {
    const detail = cause instanceof Error ? cause.message : String(cause);
    super(`Formula relocation failed for ${formula}: ${detail}`);
    this.name = 'FormulaRelocationError';
  }
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
  transfer: ClipboardTransfer;
  rangeMetadata: ClipboardRangeMetadata;
  representations?: ClipboardRepresentation[];
  mime?: string;
  html?: string;
  text?: string;
  source?: string;
}

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

  const rangeMetadata = captureRangeMetadata(sheet, range);
  const internal = JSON.stringify({ range, values, rangeMetadata });
  const text = formatTsv(values);
  const html = formatHtml(values);
  return {
    range: structuredClone(range),
    values,
    transfer: 'copy',
    rangeMetadata,
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

function captureRangeMetadata(sheet: ReturnType<WorkbookModel['getSheet']>, range: RangeRef): ClipboardRangeMetadata {
  const notes: Array<ClipboardCellMetadata<CellNote>> = [];
  for (const [key, note] of sheet.notes) {
    const row = Number(key.split(':')[0]);
    const column = Number(key.split(':')[1]);
    if (!Number.isInteger(row) || !Number.isInteger(column)) continue;
    if (row < range.startRow || row > range.endRow || column < range.startColumn || column > range.endColumn) continue;
    notes.push({ rowOffset: row - range.startRow, columnOffset: column - range.startColumn, value: structuredClone(note) });
  }
  const comments = sheet.commentThreads
    .filter((thread) => thread.row >= range.startRow && thread.row <= range.endRow && thread.column >= range.startColumn && thread.column <= range.endColumn)
    .map((thread) => ({ rowOffset: thread.row - range.startRow, columnOffset: thread.column - range.startColumn, value: structuredClone(thread) }));
  const hyperlinks: Array<ClipboardCellMetadata<CellHyperlink>> = [];
  for (const [key, hyperlink] of sheet.hyperlinks) {
    const row = Number(key.split(':')[0]);
    const column = Number(key.split(':')[1]);
    if (!Number.isInteger(row) || !Number.isInteger(column)) continue;
    if (row < range.startRow || row > range.endRow || column < range.startColumn || column > range.endColumn) continue;
    hyperlinks.push({ rowOffset: row - range.startRow, columnOffset: column - range.startColumn, value: structuredClone(hyperlink) });
  }
  const intersects = (candidate: RangeRef): boolean => candidate.sheetId === range.sheetId
    && candidate.startRow <= range.endRow && candidate.endRow >= range.startRow
    && candidate.startColumn <= range.endColumn && candidate.endColumn >= range.startColumn;
  const clip = (candidate: RangeRef): RangeRef => ({
    sheetId: range.sheetId,
    startRow: Math.max(candidate.startRow, range.startRow),
    endRow: Math.min(candidate.endRow, range.endRow),
    startColumn: Math.max(candidate.startColumn, range.startColumn),
    endColumn: Math.min(candidate.endColumn, range.endColumn),
  });
  return {
    columnWidths: Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, index) => {
      const column = range.startColumn + index;
      const widthPx = sheet.columnWidthsPx[column];
      return widthPx === undefined ? undefined : { offset: index, widthPx };
    }).filter((entry): entry is ClipboardColumnWidth => entry !== undefined),
    validations: sheet.dataValidations.filter((rule) => rule.ranges.some(intersects)).map((rule) => ({
      ...structuredClone(rule),
      ranges: rule.ranges.filter(intersects).map(clip),
    })),
    conditionalFormats: sheet.conditionalFormats.filter((rule) => rule.ranges.some(intersects)).map((rule) => ({
      ...structuredClone(rule),
      ranges: rule.ranges.filter(intersects).map(clip),
    })),
    notes,
    comments,
    hyperlinks,
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
      const parsed = JSON.parse(internal) as { values?: CellData[][]; rangeMetadata?: ClipboardRangeMetadata };
      if (parsed.rangeMetadata && isRangeMetadata(parsed.rangeMetadata)) payload.rangeMetadata = structuredClone(parsed.rangeMetadata);
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

function isRangeMetadata(value: unknown): value is ClipboardRangeMetadata {
  if (!value || typeof value !== 'object') return false;
  const record = value as Record<string, unknown>;
  return Array.isArray(record.columnWidths)
    && Array.isArray(record.validations)
    && Array.isArray(record.conditionalFormats)
    && Array.isArray(record.notes)
    && Array.isArray(record.comments)
    && Array.isArray(record.hyperlinks);
}

function escapeTsvField(value: string): string {
  return /[\t\r\n"]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
}

function parseClipboardScalar(value: string): CellData {
  return parseCellText(value);
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

/** Shift relative references for copy/fill/paste using the canonical formula AST. */
export function shiftFormula(formula: string, rowOffset: number, colOffset: number): string {
  if (!formula.trim().startsWith('=')) return formula;
  try {
    return formatFormula(offsetAst(parseFormula(formula), rowOffset, colOffset));
  } catch (error) {
    throw new FormulaRelocationError(formula, rowOffset, colOffset, error);
  }
}
