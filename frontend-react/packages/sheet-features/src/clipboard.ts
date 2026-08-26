import type {
  CellData,
  CellHyperlink,
  CellNote,
  CommentThread,
  ConditionalFormatRule,
  DataValidationRule,
  RangeRef,
  WorkbookModel,
  WorkbookTheme,
} from '@react-sheets/core-model';
import { formatFormula, offsetAst, parseFormula } from '@react-sheets/formula-engine';
import { parseCellText, type CellInputInterpretationContext } from './text-input';

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
export function isPasteSpecialSpecSupported(spec: PasteSpecialSpec, payload?: ClipboardPayload): boolean {
  if (spec.formatting === 'source-theme' && !payload?.rangeMetadata.sourceWorkbookThemeRef) return false;
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
  sourceWorkbookThemeRef?: WorkbookTheme;
}

export interface SparseClipboardCell {
  rowOffset: number;
  columnOffset: number;
  value: CellData;
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
 * Host-neutral clipboard contract. A host may expose any subset of the
 * representations, but the command layer always consumes this sparse payload.
 * `representations` is intentionally data-only so desktop and headless hosts
 * can serialize it without a DOM or ClipboardItem dependency.
 */
export interface SparseClipboardPayload {
  schema: 'SparseClipboardPayload';
  range: RangeRef;
  sourceExtent: { rows: number; columns: number };
  occupiedCells: SparseClipboardCell[];
  rangeMetadata: ClipboardRangeMetadata;
}

export interface ClipboardPayload extends SparseClipboardPayload {
  transfer: ClipboardTransfer;
  representations?: ClipboardRepresentation[];
  mime?: string;
  html?: string;
  text?: string;
  source?: string;
}

export function copyRangeToClipboardData(workbook: WorkbookModel, range: RangeRef): ClipboardPayload {
  const sheet = workbook.getSheet(range.sheetId);
  const occupiedCells: SparseClipboardCell[] = [];
  sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (cell, row, column) => {
    occupiedCells.push({ rowOffset: row - range.startRow, columnOffset: column - range.startColumn, value: structuredClone(cell) });
  });
  const rangeMetadata = captureRangeMetadata(workbook, sheet, range);
  return {
    schema: 'SparseClipboardPayload',
    range: structuredClone(range),
    sourceExtent: { rows: range.endRow - range.startRow + 1, columns: range.endColumn - range.startColumn + 1 },
    occupiedCells,
    rangeMetadata,
    transfer: 'copy',
    source: 'internal',
  };
}

function captureRangeMetadata(workbook: WorkbookModel, sheet: ReturnType<WorkbookModel['getSheet']>, range: RangeRef): ClipboardRangeMetadata {
  const notes: Array<ClipboardCellMetadata<CellNote>> = [];
  for (const { row, column, note } of sheet.review.noteEntries()) {
    if (row < range.startRow || row > range.endRow || column < range.startColumn || column > range.endColumn) continue;
    notes.push({ rowOffset: row - range.startRow, columnOffset: column - range.startColumn, value: note });
  }
  const comments = sheet.review.threadEntries()
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
    columnWidths: Object.entries(sheet.columnWidthsPx)
      .map(([columnText, widthPx]) => ({ column: Number(columnText), widthPx }))
      .filter((entry) => entry.column >= range.startColumn && entry.column <= range.endColumn)
      .map((entry) => ({ offset: entry.column - range.startColumn, widthPx: entry.widthPx })),
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
    sourceWorkbookThemeRef: structuredClone(workbook.theme),
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

export function parseTsv(text: string, inputContext: CellInputInterpretationContext): CellData[][] {
  parseCellText('', inputContext);
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
  return rows.map((values) => values.map((value) => parseClipboardScalar(value, inputContext)));
}

/** Parse a serialized clipboard payload supplied by a host ClipboardItem. */
export function sparseClipboardFromDense(range: RangeRef, values: CellData[][]): ClipboardPayload {
  const occupiedCells: SparseClipboardCell[] = [];
  for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < (values[rowOffset]?.length ?? 0); columnOffset += 1) {
      const value = values[rowOffset]?.[columnOffset];
      if (value !== undefined) occupiedCells.push({ rowOffset, columnOffset, value: structuredClone(value) });
    }
  }
  return {
    schema: 'SparseClipboardPayload',
    range: structuredClone(range),
    sourceExtent: { rows: values.length, columns: Math.max(0, ...values.map((row) => row.length)) },
    occupiedCells,
    rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
    transfer: 'copy',
    source: 'external',
  };
}

export function clipboardRepresentations(payload: ClipboardPayload): ClipboardRepresentation[] {
  const internal = JSON.stringify({
    schema: payload.schema,
    range: payload.range,
    sourceExtent: payload.sourceExtent,
    occupiedCells: payload.occupiedCells,
    rangeMetadata: payload.rangeMetadata,
  });
  const values: CellData[][] = Array.from({ length: payload.sourceExtent.rows }, () => []);
  for (const entry of payload.occupiedCells) {
    values[entry.rowOffset] ??= [];
    values[entry.rowOffset]![entry.columnOffset] = structuredClone(entry.value);
  }
  for (const row of values) for (let column = 0; column < payload.sourceExtent.columns; column += 1) row[column] ??= { value: null };
  const text = formatTsv(values);
  const html = formatHtml(values);
  return [
    { mime: CLIPBOARD_INTERNAL_MIME, data: internal },
    { mime: CLIPBOARD_HTML_MIME, data: html },
    { mime: CLIPBOARD_TEXT_MIME, data: text },
  ];
}

export function parseClipboardPayload(payload: ClipboardPayload, inputContext: CellInputInterpretationContext): ClipboardPayload {
  parseCellText('', inputContext);
  const internal = payload.representations?.find((entry) => entry.mime === CLIPBOARD_INTERNAL_MIME)?.data;
  if (internal) {
    try {
      const parsed = JSON.parse(internal) as Partial<ClipboardPayload>;
      if (parsed.schema === 'SparseClipboardPayload' && parsed.sourceExtent && Array.isArray(parsed.occupiedCells) && parsed.rangeMetadata && isRangeMetadata(parsed.rangeMetadata)) {
        return {
          ...structuredClone(payload),
          schema: 'SparseClipboardPayload',
          sourceExtent: structuredClone(parsed.sourceExtent),
          occupiedCells: structuredClone(parsed.occupiedCells),
          rangeMetadata: structuredClone(parsed.rangeMetadata),
        };
      }
    } catch {
      // Continue to the next representation; malformed host data is rejected
      // by the canonical payload validation below.
    }
  }
  if (payload.schema === 'SparseClipboardPayload' && !payload.representations?.length && payload.html === undefined && payload.text === undefined) return structuredClone(payload);
  const html = payload.representations?.find((entry) => entry.mime === CLIPBOARD_HTML_MIME)?.data ?? payload.html;
  if (html) return sparseClipboardFromDense(payload.range, parseHtmlTable(html, inputContext));
  const text = payload.representations?.find((entry) => entry.mime === CLIPBOARD_TEXT_MIME)?.data ?? payload.text;
  return text ? sparseClipboardFromDense(payload.range, parseTsv(text, inputContext)) : sparseClipboardFromDense(payload.range, []);
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

function parseClipboardScalar(value: string, inputContext: CellInputInterpretationContext): CellData {
  const parsed = parseCellText(value, inputContext);
  const cell: CellData = { value: parsed.value };
  if (parsed.formula !== undefined) cell.formula = parsed.formula;
  if (parsed.numberFormatIntent.kind === 'set') {
    cell.numberFormat = parsed.numberFormatIntent.format;
    cell.style = { numberFormat: parsed.numberFormatIntent.format };
  }
  return cell;
}

function formatHtml(values: CellData[][]): string {
  const rows = values.map((row) => `<tr>${row.map((cell) => {
    const value = cell.value == null ? '' : escapeHtml(String(cell.value));
    const formula = cell.formula ? ` data-formula="${escapeHtml(cell.formula)}"` : '';
    return `<td${formula}>${value}</td>`;
  }).join('')}</tr>`).join('');
  return `<table>${rows}</table>`;
}

function parseHtmlTable(html: string, inputContext: CellInputInterpretationContext): CellData[][] {
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
      const parsed = parseClipboardScalar(text, inputContext);
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
