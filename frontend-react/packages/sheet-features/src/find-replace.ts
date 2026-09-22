import type {
  CellData,
  CellNote,
  CommentThread,
  FormulaErrorCode,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { createFormulaError, cellKey } from '@react-sheets/core-model';
import { isFormulaError } from '@react-sheets/formula-engine';
import { parseCellText, type CellInputInterpretationContext, type NumberFormatIntent } from './text-input';
import type { DataRegionContext } from './data-region-context';

/** Canonical content families that Find/Replace may inspect. */
export type FindSearchTarget = 'values' | 'formulas' | 'notes' | 'comments';
export type FindScope = 'sheet' | 'workbook' | 'selection';
export type FindSearchOrder = 'rows' | 'columns';
export type FindDirection = 'next' | 'previous';

export interface FindSearchParams {
  sheetId: string;
  query: string;
  searchOrder: FindSearchOrder;
  scope?: FindScope;
  range?: RangeRef;
  targets?: readonly FindSearchTarget[];
  matchCase?: boolean;
  entireCell?: boolean;
  wildcard?: boolean;
  dataRegionContext?: DataRegionContext;
}

export interface FindMatch {
  readonly key: string;
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly target: FindSearchTarget;
  readonly sourceId?: string;
  readonly text: string;
  readonly range: RangeRef;
}

export interface FindCursor {
  readonly key: string;
}

export interface FindSearchResult {
  readonly matches: readonly FindMatch[];
  readonly total: number;
}

export type ReplacementValue =
  | { kind: 'empty'; value: null; numberFormatIntent?: NumberFormatIntent }
  | { kind: 'text'; value: string; numberFormatIntent?: NumberFormatIntent }
  | { kind: 'number'; value: number; numberFormatIntent?: NumberFormatIntent }
  | { kind: 'boolean'; value: boolean; numberFormatIntent?: NumberFormatIntent }
  | { kind: 'formula'; value: null; formula: string; numberFormatIntent?: NumberFormatIntent }
  | { kind: 'error'; value: null; code: FormulaErrorCode; numberFormatIntent?: NumberFormatIntent };

const REPLACEMENT_ERROR_CODES: ReadonlySet<string> = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
  '#CALC!', '#BLOCKED!', '#SPILL!', '#PARSE!',
]);

/** Parse replacement input through the same workbook interpreter as cell entry. */
export function parseReplacementValue(text: string, inputContext: CellInputInterpretationContext): ReplacementValue {
  if (text === '') return { kind: 'empty', value: null };
  let parsed: ReturnType<typeof parseCellText>;
  try { parsed = parseCellText(text, inputContext); }
  catch (error) {
    if (text.startsWith('=')) throw new Error(`Invalid replacement formula: ${error instanceof Error ? error.message : String(error)}`);
    throw error;
  }
  const formatIntent = parsed.numberFormatIntent.kind === 'set' ? { numberFormatIntent: parsed.numberFormatIntent } : {};
  if (parsed.formula !== undefined) return { kind: 'formula', value: null, formula: parsed.formula, ...formatIntent };
  const upper = text.toUpperCase();
  if (REPLACEMENT_ERROR_CODES.has(upper) && parsed.value === text && !isTextInputContext(inputContext)) return { kind: 'error', value: null, code: upper as FormulaErrorCode };
  if (typeof parsed.value === 'number') return { kind: 'number', value: parsed.value, ...formatIntent };
  if (typeof parsed.value === 'boolean') return { kind: 'boolean', value: parsed.value, ...formatIntent };
  if (parsed.value === null) return { kind: 'empty', value: null };
  return { kind: 'text', value: parsed.value, ...formatIntent };
}

function isTextInputContext(context: CellInputInterpretationContext): boolean {
  if (context.currentCellType === 'text') return true;
  const section = context.currentNumberFormat?.split(';')[0];
  return Boolean(section?.replace(/"(?:[^"]|"")*"/g, '').replace(/\\./g, '').includes('@'));
}

export function replacementCell(cell: CellData, replacement: ReplacementValue): CellData {
  if (replacement.kind === 'empty') throw new Error('Replacement text must not be empty');
  const next = structuredClone(cell);
  delete next.displayValue;
  delete next.formula;
  delete next.formulaMetadata;
  delete next.formulaValue;
  if (replacement.kind === 'formula') {
    next.value = null;
    next.formula = replacement.formula;
  } else if (replacement.kind === 'error') {
    next.value = null;
    next.formulaValue = createFormulaError(replacement.code);
  } else {
    next.value = replacement.value;
  }
  if (replacement.numberFormatIntent?.kind === 'set') {
    next.numberFormat = replacement.numberFormatIntent.format;
    next.style = { ...(next.style ?? {}), numberFormat: replacement.numberFormatIntent.format };
  }
  return next;
}

const TARGET_ORDER: readonly FindSearchTarget[] = ['values', 'formulas', 'notes', 'comments'];

function normalizeTargets(targets: readonly FindSearchTarget[] | undefined): FindSearchTarget[] {
  const requested = targets ?? ['values', 'formulas'];
  const result: FindSearchTarget[] = [];
  for (const target of TARGET_ORDER) {
    if (requested.includes(target)) result.push(target);
  }
  if (result.length === 0) throw new Error('Find requires at least one content target');
  return result;
}

function normalizeRange(range: RangeRef, sheet: WorksheetModel): RangeRef {
  if (range.sheetId !== sheet.id || !Number.isSafeInteger(range.startRow) || !Number.isSafeInteger(range.endRow)
    || !Number.isSafeInteger(range.startColumn) || !Number.isSafeInteger(range.endColumn)
    || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
    throw new Error(`Find range is outside ${sheet.id} bounds`);
  }
  return structuredClone(range);
}

function fullRange(sheet: WorksheetModel): RangeRef {
  return { sheetId: sheet.id, startRow: 0, endRow: sheet.rowCount - 1, startColumn: 0, endColumn: sheet.columnCount - 1 };
}

function inRange(row: number, column: number, range: RangeRef): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function wildcardPattern(query: string): string {
  let pattern = '';
  for (let index = 0; index < query.length; index += 1) {
    const character = query[index]!;
    if (character === '~') {
      const escaped = query[index + 1];
      if (escaped === undefined) pattern += escapeRegExp(character);
      else { pattern += escapeRegExp(escaped); index += 1; }
    } else if (character === '*') pattern += '.*';
    else if (character === '?') pattern += '.';
    else pattern += escapeRegExp(character);
  }
  return pattern;
}

function matcher(params: Pick<FindSearchParams, 'query' | 'matchCase' | 'entireCell' | 'wildcard'>): RegExp {
  if (!params.query) throw new Error('Find query must not be empty');
  const source = params.wildcard ? wildcardPattern(params.query) : escapeRegExp(params.query);
  const flags = params.matchCase ? 'g' : 'gi';
  return new RegExp(params.entireCell ? `^${source}$` : source, flags);
}

export function matchesFindText(text: string, params: Pick<FindSearchParams, 'query' | 'matchCase' | 'entireCell' | 'wildcard'>): boolean {
  const expression = matcher(params);
  return expression.test(text);
}

/** Replace every matched span with a literal replacement (never `$1` expansion). */
export function replaceFindText(text: string, params: Pick<FindSearchParams, 'query' | 'matchCase' | 'entireCell' | 'wildcard'>, replacement: string): string | undefined {
  const expression = matcher(params);
  if (!expression.test(text)) return undefined;
  expression.lastIndex = 0;
  return text.replace(expression, () => replacement);
}

function scalarText(value: unknown): string {
  if (isFormulaError(value)) return value.code;
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
  throw new Error(`Find resolved value has unsupported type: ${typeof value}`);
}

function cellValueText(sheet: WorksheetModel, row: number, column: number, resolveCellValue?: FindResolveCellValue): string {
  const resolved = resolveCellValue?.(sheet, row, column);
  const cell = sheet.cells.get(row, column);
  return scalarText(resolved === undefined ? (cell?.formulaValue ?? cell?.value) : resolved);
}

export type FindResolveCellValue = (sheet: WorksheetModel, row: number, column: number) => unknown;

function addMatch(matches: FindMatch[], sheet: WorksheetModel, row: number, column: number, target: FindSearchTarget, text: string, sourceId?: string): void {
  if (text === '') return;
  const key = `${sheet.id}!${row}:${column}:${target}:${sourceId ?? ''}`;
  matches.push({
    key,
    sheetId: sheet.id,
    row,
    column,
    target,
    ...(sourceId === undefined ? {} : { sourceId }),
    text,
    range: { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column },
  });
}

interface FindMetadataComment {
  id: string;
  text: string;
}

interface FindMetadataBucket {
  note?: CellNote;
  comments: FindMetadataComment[];
}

function buildFindMetadataIndex(sheet: WorksheetModel): Map<string, FindMetadataBucket> {
  const index = new Map<string, FindMetadataBucket>();
  const bucketFor = (key: string): FindMetadataBucket => {
    const existing = index.get(key);
    if (existing) return existing;
    const created: FindMetadataBucket = { comments: [] };
    index.set(key, created);
    return created;
  };
  for (const entry of sheet.review.noteEntries()) bucketFor(entry.key).note = entry.note;
  for (const thread of sheet.review.threadEntries()) {
    bucketFor(`${thread.row}:${thread.column}`).comments.push({ id: thread.id, text: thread.text });
  }
  return index;
}

function scanSheet(
  sheet: WorksheetModel,
  range: RangeRef,
  targets: readonly FindSearchTarget[],
  params: FindSearchParams,
  matches: FindMatch[],
  resolveCellValue?: FindResolveCellValue,
): void {
  const targetSet = new Set(targets);
  const metadataIndex = buildFindMetadataIndex(sheet);
  const cells = new Map<string, CellData>();
  sheet.cells.forEach((cell, row, column) => {
    if (inRange(row, column, range)) cells.set(`${row}:${column}`, cell);
  });
  const coordinates = [...cells.keys()].map((key) => key.split(':').map(Number) as [number, number]);
  for (const [row, column] of coordinates.sort(([ar, ac], [br, bc]) => ar - br || ac - bc)) {
    const cell = cells.get(`${row}:${column}`)!;
    // In the canonical "values + formulas" family a formula cell is one
    // logical search item: formula text takes precedence over its result.
    // A values-only search still searches the resolved formula result.
    if (targetSet.has('values') && (cell.formula === undefined || !targetSet.has('formulas'))) {
      const text = cellValueText(sheet, row, column, resolveCellValue);
      if (matchesFindText(text, params)) addMatch(matches, sheet, row, column, 'values', text);
    }
    if (targetSet.has('formulas') && cell.formula !== undefined && matchesFindText(cell.formula, params)) {
      addMatch(matches, sheet, row, column, 'formulas', cell.formula);
    }
    if (targetSet.has('notes')) {
      const note = metadataIndex.get(`${row}:${column}`)?.note;
      if (note && matchesFindText(note.text, params)) addMatch(matches, sheet, row, column, 'notes', note.text, note.id);
    }
    if (targetSet.has('comments')) {
      const comments = metadataIndex.get(`${row}:${column}`)?.comments ?? [];
      for (const comment of comments) if (matchesFindText(comment.text, params)) addMatch(matches, sheet, row, column, 'comments', comment.text, comment.id);
    }
  }
  // Notes/comments may be attached to an otherwise empty cell and therefore
  // are not present in CellMatrix. Include them in the same row/column order.
  const metadataCoordinates = new Set(coordinates.map(([row, column]) => `${row}:${column}`));
  const extras: Array<{ row: number; column: number }> = [];
  for (const [key] of metadataIndex) {
    const parts = key.split(':').map(Number);
    const row = parts[0];
    const column = parts[1];
    if (row === undefined || column === undefined) continue;
    if (!metadataCoordinates.has(key) && inRange(row, column, range)) extras.push({ row, column });
  }
  for (const { row, column } of extras.sort((a, b) => a.row - b.row || a.column - b.column)) {
    const metadata = metadataIndex.get(`${row}:${column}`);
    if (targetSet.has('notes')) {
      const note = metadata?.note;
      if (note && matchesFindText(note.text, params)) addMatch(matches, sheet, row, column, 'notes', note.text, note.id);
    }
    if (targetSet.has('comments')) {
      for (const comment of metadata?.comments ?? []) {
        if (matchesFindText(comment.text, params)) addMatch(matches, sheet, row, column, 'comments', comment.text, comment.id);
      }
    }
  }
}

/** Build a stable, workbook-order result set. Hidden rows remain searchable. */
export function planFind(workbook: WorkbookModel, params: FindSearchParams, resolveCellValue?: FindResolveCellValue): FindSearchResult {
  if (!params.sheetId || !params.query) throw new Error('Find requires a sheet and non-empty query');
  if (params.searchOrder !== 'rows' && params.searchOrder !== 'columns') throw new Error('Find requires a valid search order');
  if (params.scope === 'selection' && params.range === undefined && params.dataRegionContext === undefined) throw new Error('Selection Find requires an explicit range');
  if (params.dataRegionContext && params.dataRegionContext.sheetId !== params.sheetId) throw new Error('Find DataRegionContext targets another sheet');
  const targets = normalizeTargets(params.targets);
  const selectedSheet = workbook.getSheet(params.sheetId);
  const sheets = params.scope === 'workbook' ? workbook.getSheets() : [selectedSheet];
  const matches: FindMatch[] = [];
  for (const sheet of sheets) {
    const range = params.scope === 'selection' || (params.range && sheet.id === params.sheetId)
      ? normalizeRange(params.range ?? params.dataRegionContext?.range ?? fullRange(sheet), sheet)
      : fullRange(sheet);
    scanSheet(sheet, range, targets, params, matches, resolveCellValue);
  }
  // The scan is target-major within each cell; normalize to one observable
  // order independent of map insertion and metadata storage order.
  const targetRank = new Map(TARGET_ORDER.map((target, index) => [target, index]));
  const compareCoordinates = params.searchOrder === 'rows'
    ? (a: FindMatch, b: FindMatch) => a.row - b.row || a.column - b.column
    : (a: FindMatch, b: FindMatch) => a.column - b.column || a.row - b.row;
  matches.sort((a, b) => {
    const sheetOrder = sheets.indexOf(workbook.getSheet(a.sheetId)) - sheets.indexOf(workbook.getSheet(b.sheetId));
    return sheetOrder || compareCoordinates(a, b) || (targetRank.get(a.target)! - targetRank.get(b.target)!) || a.key.localeCompare(b.key);
  });
  return { matches, total: matches.length };
}

export function findAtCursor(matches: readonly FindMatch[], cursor: FindCursor | null, direction: FindDirection): FindMatch | undefined {
  if (matches.length === 0) return undefined;
  if (!cursor) return direction === 'next' ? matches[0] : matches[matches.length - 1];
  const current = matches.findIndex((match) => match.key === cursor.key);
  if (current < 0) return direction === 'next' ? matches[0] : matches[matches.length - 1];
  const offset = direction === 'next' ? 1 : -1;
  return matches[(current + offset + matches.length) % matches.length];
}

export function findCursorFor(match: FindMatch): FindCursor {
  return { key: match.key };
}

export function noteAt(sheet: WorksheetModel, row: number, column: number): CellNote | undefined {
  return sheet.review.getNoteAt(row, column);
}

export function noteKey(row: number, column: number): string {
  return cellKey(row, column);
}
