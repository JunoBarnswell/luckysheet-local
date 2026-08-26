import type { CellData, RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRuntime, CommandResult } from '@react-sheets/command-runtime';
import {
  matchesFindText,
  noteAt,
  parseReplacementValue,
  planFind,
  replacementCell,
  replaceFindText,
  type FindMatch,
  type FindResolveCellValue,
  type FindSearchOrder,
  type FindScope,
  type FindSearchTarget,
  type CellInputInterpretationContext,
  isCellInputInterpretationContext,
  type DataRegionContext,
} from '@react-sheets/sheet-features';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';

export interface FindReplaceParams {
  sheetId: string;
  query: string;
  replace: string;
  mode: 'one' | 'all';
  searchOrder: FindSearchOrder;
  matchKey?: string;
  scope?: FindScope;
  range?: RangeRef;
  targets?: readonly FindSearchTarget[];
  matchCase?: boolean;
  entireCell?: boolean;
  wildcard?: boolean;
  dataRegionContext?: DataRegionContext;
  inputContext: CellInputInterpretationContext;
}

export interface FindCellReplacementMutationPatch {
  kind: 'cell';
  match: FindMatch;
  previous?: CellData;
  next: CellData;
}

export interface FindNoteReplacementMutationPatch {
  kind: 'note';
  match: FindMatch;
  previous?: import('@react-sheets/core-model').CellNote;
  next: import('@react-sheets/core-model').CellNote;
}

export interface FindCommentReplacementMutationPatch {
  kind: 'comment';
  match: FindMatch;
  previousText: string;
  nextText: string;
}

export type FindReplacementMutationPatch = FindCellReplacementMutationPatch | FindNoteReplacementMutationPatch | FindCommentReplacementMutationPatch;

export interface FindReplacementMutationParams {
  direction: 'forward' | 'reverse';
  patches: readonly FindReplacementMutationPatch[];
  affectedRanges: readonly RangeRef[];
  dataRegionContext?: DataRegionContext;
}

interface CellPatch {
  kind: 'cell';
  match: FindMatch;
  previous?: CellData;
  next: CellData;
}

interface NotePatch {
  kind: 'note';
  match: FindMatch;
  previous?: import('@react-sheets/core-model').CellNote;
  next: import('@react-sheets/core-model').CellNote;
  text: string;
}

interface CommentPatch {
  kind: 'comment';
  match: FindMatch;
  previousText: string;
  nextText: string;
  text: string;
}

type ReplacementPatch = CellPatch | NotePatch | CommentPatch;

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);
const isRange = (value: unknown): value is RangeRef => isRecord(value)
  && typeof value.sheetId === 'string'
  && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
  && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn)
  && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
  && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);

function isFindTarget(value: unknown): value is FindSearchTarget {
  return value === 'values' || value === 'formulas' || value === 'notes' || value === 'comments';
}

function isValidFindReplace(value: unknown): value is FindReplaceParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string' || value.sheetId.length === 0
    || typeof value.query !== 'string' || value.query.length === 0 || typeof value.replace !== 'string'
    || (value.mode !== 'one' && value.mode !== 'all')
    || (value.searchOrder !== 'rows' && value.searchOrder !== 'columns')) return false;
  if (!isCellInputInterpretationContext(value.inputContext)) return false;
  if (value.mode === 'one' && (typeof value.matchKey !== 'string' || value.matchKey.length === 0)) return false;
  if (value.scope !== undefined && value.scope !== 'sheet' && value.scope !== 'workbook' && value.scope !== 'selection') return false;
  if (value.scope === 'selection' && value.range === undefined) return false;
  if (value.range !== undefined && !isRange(value.range)) return false;
  if (value.targets !== undefined && (!Array.isArray(value.targets) || value.targets.length === 0 || !value.targets.every(isFindTarget))) return false;
  return (value.matchCase === undefined || typeof value.matchCase === 'boolean')
    && (value.entireCell === undefined || typeof value.entireCell === 'boolean')
    && (value.wildcard === undefined || typeof value.wildcard === 'boolean');
}

function isFindMatch(value: unknown): value is FindMatch {
  return isRecord(value) && typeof value.key === 'string' && typeof value.sheetId === 'string'
    && Number.isSafeInteger(value.row) && Number.isSafeInteger(value.column)
    && (value.target === 'values' || value.target === 'formulas' || value.target === 'notes' || value.target === 'comments')
    && isRange(value.range) && typeof value.text === 'string'
    && (value.sourceId === undefined || typeof value.sourceId === 'string');
}

function isCellData(value: unknown): value is CellData {
  return isRecord(value) && 'value' in value;
}

function isFindReplacementPatch(value: unknown): value is FindReplacementMutationPatch {
  if (!isRecord(value) || !isFindMatch(value.match)) return false;
  if (value.kind === 'cell') return isCellData(value.next) && (value.previous === undefined || isCellData(value.previous));
  if (value.kind === 'note') return isRecord(value.next) && typeof value.next.id === 'string' && typeof value.next.text === 'string'
    && (value.previous === undefined || isRecord(value.previous));
  return value.kind === 'comment' && typeof value.previousText === 'string' && typeof value.nextText === 'string';
}

function isFindReplacementMutation(value: unknown): value is FindReplacementMutationParams {
  return isRecord(value) && (value.direction === 'forward' || value.direction === 'reverse')
    && Array.isArray(value.patches) && value.patches.length > 0 && value.patches.every(isFindReplacementPatch)
    && Array.isArray(value.affectedRanges) && value.affectedRanges.every(isRange);
}

function equalValue(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function applyFindReplacementMutation(params: FindReplacementMutationParams, context: CommandContext): void {
  if (!isFindReplacementMutation(params)) throw new Error('Invalid find.replaced mutation parameters');
  const forward = params.direction === 'forward';
  // Validate every current source before touching any target. This keeps
  // local undo/redo and committed replay atomic on stale operation payloads.
  for (const patch of params.patches) {
    const sheet = context.workbook.getSheet(patch.match.sheetId);
    if (patch.kind === 'cell') {
      const current = sheet.cells.get(patch.match.row, patch.match.column);
      const expected = forward ? patch.previous : patch.next;
      if (!equalValue(current, expected)) throw new Error(`Find replacement source changed at ${patch.match.key}`);
    } else if (patch.kind === 'note') {
      const current = sheet.notes.get(`${patch.match.row}:${patch.match.column}`);
      const expected = forward ? patch.previous : patch.next;
      if (!equalValue(current, expected)) throw new Error(`Find note source changed at ${patch.match.key}`);
    } else {
      const thread = sheet.commentThreads.find((entry) => entry.id === patch.match.sourceId);
      const expected = forward ? patch.previousText : patch.nextText;
      if (!thread || thread.row !== patch.match.row || thread.column !== patch.match.column || thread.text !== expected) throw new Error(`Find comment source changed at ${patch.match.key}`);
    }
  }
  for (const patch of params.patches) {
    const sheet = context.workbook.getSheet(patch.match.sheetId);
    if (patch.kind === 'cell') {
      const next = forward ? patch.next : patch.previous;
      if (next === undefined) sheet.cells.delete(patch.match.row, patch.match.column);
      else sheet.cells.set(patch.match.row, patch.match.column, structuredClone(next));
    } else if (patch.kind === 'note') {
      const next = forward ? patch.next : patch.previous;
      const key = `${patch.match.row}:${patch.match.column}`;
      if (next === undefined) sheet.notes.delete(key);
      else sheet.notes.set(key, structuredClone(next));
    } else {
      const thread = sheet.commentThreads.find((entry) => entry.id === patch.match.sourceId);
      if (!thread) throw new Error(`Find comment ${patch.match.sourceId} disappeared during commit`);
      thread.text = forward ? patch.nextText : patch.previousText;
    }
  }
}

function assertCurrentMatch(match: FindMatch, params: FindReplaceParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(match.sheetId);
  const cell = sheet.cells.get(match.row, match.column);
  let text: string | undefined;
  if (match.target === 'values') {
    const resolved = context.resolveCellValue?.(sheet, match.row, match.column);
    const value = resolved === undefined ? (cell?.formulaValue ?? cell?.value) : resolved;
    if (value !== null && value !== undefined && (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')) text = String(value);
    else if (value && typeof value === 'object' && 'code' in value && typeof (value as { code?: unknown }).code === 'string') text = (value as { code: string }).code;
    else text = value == null ? '' : undefined;
  } else if (match.target === 'formulas') text = cell?.formula;
  else if (match.target === 'notes') text = noteAt(sheet, match.row, match.column)?.text;
  else text = sheet.commentThreads.find((thread) => thread.id === match.sourceId)?.text ?? cell?.comment?.text;
  if (text === undefined || !matchesFindText(text, params)) throw new Error(`Find match ${match.key} changed before replacement`);
}

function buildPatches(params: FindReplaceParams, context: CommandContext): ReplacementPatch[] {
  const result = planFind(context.workbook, {
    sheetId: params.sheetId,
    query: params.query,
    searchOrder: params.searchOrder,
    scope: params.scope,
    range: params.range,
    targets: params.targets,
    matchCase: params.matchCase,
    entireCell: params.entireCell,
    wildcard: params.wildcard,
    dataRegionContext: params.dataRegionContext,
  }, context.resolveCellValue as FindResolveCellValue | undefined);
  const matches = params.mode === 'one'
    ? [result.matches.find((match) => match.key === params.matchKey)].filter((match): match is FindMatch => match !== undefined)
    : [...result.matches];
  if (params.mode === 'one' && matches.length === 0) throw new Error(`Find match ${params.matchKey} is no longer available`);
  const patches: ReplacementPatch[] = [];
  const touchedCells = new Set<string>();
  for (const match of matches) {
    assertCurrentMatch(match, params, context);
    const key = `${match.sheetId}:${match.row}:${match.column}`;
    if (match.target === 'values' || match.target === 'formulas') {
      if (touchedCells.has(key)) throw new Error(`Find result has conflicting content targets at ${match.sheetId}!${match.row}:${match.column}`);
      const sheet = context.workbook.getSheet(match.sheetId);
      const cell = sheet.cells.get(match.row, match.column);
      if (!cell) throw new Error(`Find result cell disappeared at ${match.key}`);
      const replaced = replaceFindText(match.text, params, params.replace);
      if (replaced === undefined) throw new Error(`Find result no longer matches at ${match.key}`);
      const replacement = parseReplacementValue(replaced, {
        ...params.inputContext,
        currentNumberFormat: cell.numberFormat ?? cell.style?.numberFormat,
        currentCellType: cell.editor?.kind,
      });
      if (replacement.kind === 'empty') throw new Error('Replacement text must not be empty');
      if (match.target === 'formulas' && replacement.kind !== 'formula') throw new Error(`Formula replacement at ${match.key} must produce a formula`);
      patches.push({ kind: 'cell', match, previous: structuredClone(cell), next: replacementCell(cell, replacement) });
      touchedCells.add(key);
    } else if (match.target === 'notes') {
      const note = context.workbook.getSheet(match.sheetId).notes.get(`${match.row}:${match.column}`);
      if (!note || note.id !== match.sourceId) throw new Error(`Note ${match.sourceId} changed before replacement`);
      const replaced = replaceFindText(note.text, params, params.replace);
      if (replaced === undefined) throw new Error(`Note ${match.sourceId} no longer matches`);
      patches.push({ kind: 'note', match, previous: structuredClone(note), next: { ...structuredClone(note), text: replaced }, text: replaced });
    } else {
      const thread = context.workbook.getSheet(match.sheetId).commentThreads.find((entry) => entry.id === match.sourceId);
      if (!thread) throw new Error(`Comment ${match.sourceId} changed before replacement`);
      const replaced = replaceFindText(thread.text, params, params.replace);
      if (replaced === undefined) throw new Error(`Comment ${match.sourceId} no longer matches`);
      patches.push({ kind: 'comment', match, previousText: thread.text, nextText: replaced, text: replaced });
    }
  }
  return patches;
}

export function registerFindReplaceCommands(runtime: CommandRuntime): string[] {
  runtime.registry.registerMutation<FindReplacementMutationParams>({
    id: 'find.replaced',
    handler: (item, context) => applyFindReplacementMutation(item.params, context),
    metadata: {
      schema: { name: 'FindReplacementMutationParams', validate: isFindReplacementMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [...params.affectedRanges], mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['find.replaced'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerCommand<FindReplaceParams>({
    id: 'find.replace',
    execute: (params, context): CommandResult => {
      if (!isValidFindReplace(params)) throw new Error('Invalid find.replace parameters');
      const patches = buildPatches(params, context);
      const affectedRanges = patches.map((patch) => patch.match.range);
      const forward: FindReplacementMutationParams = { direction: 'forward', patches, affectedRanges, dataRegionContext: params.dataRegionContext };
      const inverse: FindReplacementMutationParams = { direction: 'reverse', patches, affectedRanges, dataRegionContext: params.dataRegionContext };
      context.applyMutation({ id: 'find.replaced', unitId: context.workbook.unitId, sheetId: params.sheetId, params: forward, affectedRanges: [...affectedRanges], inverse: [{ id: 'find.replaced', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges: [...affectedRanges] }], apply: () => applyFindReplacementMutation(forward, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges, event: { type: 'find.replaced', payload: { count: patches.length } } };
    },
  });
  return ['find.replace'];
}

export function registerFindReplaceFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return { id: 'find-replace', version: '1.0.0', commandIds: registerFindReplaceCommands(runtime), permissions: ['sheet.cell.write', 'review.note', 'review.comment'] };
}
