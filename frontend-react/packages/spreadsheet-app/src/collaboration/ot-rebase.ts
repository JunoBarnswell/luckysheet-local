import { mapAxisCoordinate, type RangeRef } from '@react-sheets/core-model';
import { formatFormula, mapAstStructuralReferences, parseFormula, transformReferenceInterval } from '@react-sheets/formula-engine';
import type { ClassifiedMutation, CollaborationOperationKind } from './operation-types';

export interface StructuralDelta {
  kind: 'insert-rows' | 'delete-rows' | 'insert-columns' | 'delete-columns';
  sheetId: string;
  at: number;
  count: number;
}

export interface RebaseResult {
  rebased: ClassifiedMutation;
  transformed: boolean;
}

export interface StructuralRebaseContext {
  readonly sheetOrder: readonly { readonly id: string; readonly name: string }[];
}

type StructuralAxis = 'row' | 'column';
const MAX_ROW_INDEX = 1_048_575;
const MAX_COLUMN_INDEX = 16_383;

function deltaAxis(delta: StructuralDelta): StructuralAxis {
  return delta.kind.endsWith('rows') ? 'row' : 'column';
}

function deltaOperation(delta: StructuralDelta): 'insert' | 'delete' {
  return delta.kind.startsWith('insert-') ? 'insert' : 'delete';
}

function rebaseConflict(message: string): never {
  throw new Error(`STRUCTURAL_REBASE_CONFLICT: ${message}`);
}

function shiftPoint(value: number, delta: StructuralDelta): number {
  const maximum = deltaAxis(delta) === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (!Number.isSafeInteger(value) || value < 0 || value > maximum) rebaseConflict('coordinate is outside worksheet bounds');
  const mapped = mapAxisCoordinate(value, delta.at, delta.count, deltaOperation(delta) === 'insert' ? 1 : -1);
  if (mapped === null) rebaseConflict(`coordinate ${value} was deleted by the committed operation`);
  if (!Number.isSafeInteger(mapped) || mapped < 0 || mapped > maximum) rebaseConflict('coordinate exceeds worksheet bounds');
  return mapped;
}

function shiftRange(range: RangeRef, delta: StructuralDelta): RangeRef {
  if (range.sheetId !== delta.sheetId) return range;
  const axis = deltaAxis(delta);
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  const start = range[startKey];
  const end = range[endKey];
  if (!Number.isSafeInteger(start + delta.count) || !Number.isSafeInteger(end + delta.count)) {
    rebaseConflict('range exceeds the safe integer range after transformation');
  }
  if (deltaOperation(delta) === 'insert' && delta.at > start && delta.at <= end) {
    rebaseConflict('insertion splits the pending range into non-contiguous coordinates');
  }
  const interval = transformReferenceInterval(range[startKey], range[endKey], {
    axis,
    at: delta.at,
    count: delta.count,
    op: deltaOperation(delta),
  });
  if (!interval) rebaseConflict('the committed deletion removed the entire pending range');
  const maximum = axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (interval.start < 0 || interval.end > maximum) rebaseConflict('range exceeds worksheet bounds after transformation');
  return { ...range, [startKey]: interval.start, [endKey]: interval.end };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function isRange(value: Record<string, unknown>): value is Record<string, unknown> & RangeRef {
  return typeof value.sheetId === 'string'
    && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
    && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn);
}

const ADDRESS_FIELDS = new Set(['address', 'anchor', 'cellAddress', 'formulaAnchor', 'from', 'origin', 'source', 'target', 'to']);
const ADDRESS_RANGE_FIELDS = new Set(['sourceOrigin', 'targetOrigin']);
const UNSUPPORTED_STRUCTURAL_KINDS = new Set<CollaborationOperationKind>([
  'move-range', 'sort', 'table-resize', 'sheet-identity',
]);

function transformParams(
  value: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  pendingKind: CollaborationOperationKind,
  context: StructuralRebaseContext,
  field = '',
  root = false,
): unknown {
  if (typeof value === 'string' && isFormulaField(field, value) && value.trim() !== '') {
    return transformFormulaValue(value, delta, ownerSheetId, context);
  }
  if (Array.isArray(value)) {
    if ((field === 'rowIndices' && deltaAxis(delta) === 'row')
      || (field === 'columnIndices' && deltaAxis(delta) === 'column')) {
      return value.map((entry) => {
        if (typeof entry !== 'number') rebaseConflict(`${field} contains a non-numeric coordinate`);
        return shiftPoint(entry, delta);
      });
    }
    return value.map((entry) => transformParams(entry, delta, ownerSheetId, pendingKind, context, field));
  }
  if (!isRecord(value)) return value;

  const sheetId = typeof value.sheetId === 'string' ? value.sheetId : ownerSheetId;
  const formulaOwnerSheetId = isRecord(value.formulaAnchor) && typeof value.formulaAnchor.sheetId === 'string'
    ? value.formulaAnchor.sheetId
    : sheetId;
  if (isRange(value)) return shiftRange(value, delta);

  const isAddress = ADDRESS_FIELDS.has(field) || ADDRESS_RANGE_FIELDS.has(field) || root;
  const pendingStructural = isStructuralKind(pendingKind, deltaAxis(delta));
  const next: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (root && pendingStructural && key === (deltaAxis(delta) === 'row' ? 'row' : 'column')) {
      next[key] = entry;
      continue;
    }
    if (sheetId === delta.sheetId && isAddress && key === (deltaAxis(delta) === 'row' ? 'row' : 'column')
      && typeof entry === 'number') {
      next[key] = shiftPoint(entry, delta);
      continue;
    }
    if (key === 'at' && root && sheetId === delta.sheetId && isStructuralKind(pendingKind, deltaAxis(delta))) {
      next[key] = entry;
      continue;
    }
    if (typeof entry === 'string' && (isFormulaField(key, entry) || isRuleFormulaField(key, entry, value))) {
      next[key] = transformFormulaValue(entry, delta, formulaOwnerSheetId, context);
    } else {
      next[key] = transformParams(entry, delta, sheetId, pendingKind, context, key);
    }
  }

  if (root && sheetId === delta.sheetId && isStructuralKind(pendingKind, deltaAxis(delta))) {
    const atKey = typeof value.at === 'number' ? 'at' : deltaAxis(delta) === 'row' ? 'row' : 'column';
    const at = value[atKey];
    if (typeof at === 'number') next[atKey] = shiftPoint(at, delta);
  }
  return next;
}

function isFormulaField(field: string, formula: string): boolean {
  const normalized = field.toLowerCase();
  return normalized === 'formula'
    || ((normalized === 'formula1' || normalized === 'formula2') && formula.trimStart().startsWith('='));
}

function isRuleFormulaField(field: string, formula: string, rule: Record<string, unknown>): boolean {
  const startsWithEquals = formula.trimStart().startsWith('=');
  if ((field === 'value1' || field === 'value2') && startsWithEquals) return true;
  if (field === 'value1' && rule.operator === 'formula') return true;
  if (field === 'formula1' && (rule.operator === 'formula' || rule.type === 'custom')) return true;
  return field === 'formula2' && rule.type === 'custom';
}

function transformFormulaValue(
  formula: string,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): string {
  const targetSheet = context.sheetOrder.find((sheet) => sheet.id === delta.sheetId);
  if (!targetSheet) rebaseConflict(`worksheet identity is missing for formula references on ${delta.sheetId}`);
  try {
    const hasPrefix = formula.trimStart().startsWith('=');
    const source = hasPrefix ? formula : `=${formula}`;
    const transformed = formatFormula(mapAstStructuralReferences(parseFormula(source), {
      shift: { axis: deltaAxis(delta), at: delta.at, count: delta.count, op: deltaOperation(delta) },
      ownerSheetId,
      targetSheetId: delta.sheetId,
      targetSheetName: targetSheet.name,
      sheetOrder: context.sheetOrder,
    }));
    return hasPrefix ? transformed : transformed.replace(/^=/, '');
  } catch (error) {
    rebaseConflict(`pending formula cannot be structurally transformed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function shiftSnapshotCellKey(value: unknown, sheetId: string, delta: StructuralDelta): string {
  if (typeof value !== 'string') rebaseConflict('paste snapshot contains a non-string cell metadata key');
  const match = value.match(/^(\d+):(\d+)$/);
  if (!match) rebaseConflict('paste snapshot contains an invalid cell metadata key');
  if (sheetId !== delta.sheetId) return value;
  const row = Number(match[1]);
  const column = Number(match[2]);
  return `${deltaAxis(delta) === 'row' ? shiftPoint(row, delta) : row}:${deltaAxis(delta) === 'column' ? shiftPoint(column, delta) : column}`;
}

function shiftSnapshotAddress(value: unknown, sheetId: string, delta: StructuralDelta): void {
  if (!isRecord(value)) rebaseConflict('paste snapshot contains an invalid cell coordinate');
  const ownerSheetId = typeof value.sheetId === 'string' ? value.sheetId : sheetId;
  if (ownerSheetId !== delta.sheetId) return;
  const coordinateKey = deltaAxis(delta) === 'row' ? 'row' : 'column';
  const coordinate = value[coordinateKey];
  if (typeof coordinate !== 'number') rebaseConflict(`paste snapshot ${coordinateKey} coordinate is invalid`);
  value[coordinateKey] = shiftPoint(coordinate, delta);
}

function shiftPasteSnapshot(value: unknown, sheetId: string, delta: StructuralDelta): unknown {
  if (!isRecord(value)) rebaseConflict('pending range.paste has an invalid snapshot');
  const snapshot = value;
  if (!Array.isArray(snapshot.cells)) rebaseConflict('pending range.paste snapshot cells are invalid');
  for (const cell of snapshot.cells) shiftSnapshotAddress(cell, sheetId, delta);
  for (const field of ['notes', 'hyperlinks'] as const) {
    const entries = snapshot[field];
    if (entries === undefined) continue;
    if (!Array.isArray(entries)) rebaseConflict(`pending range.paste snapshot ${field} are invalid`);
    for (const entry of entries) {
      if (!isRecord(entry)) rebaseConflict(`pending range.paste snapshot ${field} entry is invalid`);
      entry.key = shiftSnapshotCellKey(entry.key, sheetId, delta);
    }
  }
  if (snapshot.commentCells !== undefined) {
    if (!Array.isArray(snapshot.commentCells)) rebaseConflict('pending range.paste snapshot commentCells are invalid');
    snapshot.commentCells = snapshot.commentCells.map((key) => shiftSnapshotCellKey(key, sheetId, delta));
  }
  if (snapshot.comments !== undefined) {
    if (!Array.isArray(snapshot.comments)) rebaseConflict('pending range.paste snapshot comments are invalid');
    for (const comment of snapshot.comments) shiftSnapshotAddress(comment, sheetId, delta);
  }
  if (snapshot.columnWidths !== undefined) {
    if (!Array.isArray(snapshot.columnWidths)) rebaseConflict('pending range.paste snapshot columnWidths are invalid');
    if (deltaAxis(delta) === 'column' && sheetId === delta.sheetId) {
      for (const entry of snapshot.columnWidths) {
        if (!isRecord(entry) || typeof entry.column !== 'number') rebaseConflict('pending range.paste snapshot column width is invalid');
        entry.column = shiftPoint(entry.column, delta);
      }
    }
  }
  return snapshot;
}

function transformPasteSnapshots(params: unknown, transformedParams: unknown, ownerSheetId: string, delta: StructuralDelta): unknown {
  if (!isRecord(params) || !isRecord(transformedParams)) rebaseConflict('pending range.paste parameters are invalid');
  if (!Object.prototype.hasOwnProperty.call(params, 'snapshot')) return transformedParams;
  const result = transformedParams;
  result.snapshot = shiftPasteSnapshot(result.snapshot, ownerSheetId, delta);
  if (Object.prototype.hasOwnProperty.call(params, 'sourceSnapshot')) {
    const sourceRange = params.sourceRange;
    if (!isRecord(sourceRange) || typeof sourceRange.sheetId !== 'string') {
      rebaseConflict('pending cross-sheet range.paste source range is invalid');
    }
    result.sourceSnapshot = shiftPasteSnapshot(result.sourceSnapshot, sourceRange.sheetId, delta);
  }
  return result;
}

function isStructuralKind(kind: CollaborationOperationKind, axis: StructuralAxis): boolean {
  return kind === `${kind.startsWith('insert-') ? 'insert' : 'delete'}-${axis}s`;
}

function extractStructuralDelta(mutation: ClassifiedMutation): StructuralDelta | undefined {
  const p = mutation.params as { at?: number; count?: number; row?: number; column?: number } | null;
  if (!p) return undefined;
  const at = p.at ?? p.row ?? p.column;
  const count = p.count ?? 1;
  if (at == null) return undefined;

  if (!Number.isSafeInteger(at) || at < 0 || !Number.isSafeInteger(count) || count < 1) {
    rebaseConflict(`committed ${mutation.mutationId} has invalid structural bounds`);
  }
  if (!Number.isSafeInteger(at + count)) rebaseConflict(`committed ${mutation.mutationId} exceeds the safe integer range`);
  const maximum = mutation.kind.endsWith('rows') ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (at > maximum || at + count > maximum + 1) {
    rebaseConflict(`committed ${mutation.mutationId} exceeds worksheet bounds`);
  }
  switch (mutation.kind) {
    case 'insert-rows': return { kind: 'insert-rows', sheetId: mutation.sheetId, at, count };
    case 'delete-rows': return { kind: 'delete-rows', sheetId: mutation.sheetId, at, count };
    case 'insert-columns': return { kind: 'insert-columns', sheetId: mutation.sheetId, at, count };
    case 'delete-columns': return { kind: 'delete-columns', sheetId: mutation.sheetId, at, count };
    default: return undefined;
  }
}

const STRUCTURAL_KINDS = new Set<CollaborationOperationKind>([
  'insert-rows', 'delete-rows', 'insert-columns', 'delete-columns',
]);

/** 将 pending 操作按已提交的结构变更 rebase — 例: A 插第 5 行，B 改 A10 → A11 */
export function rebaseMutation(
  pending: ClassifiedMutation,
  committed: ClassifiedMutation,
  context: StructuralRebaseContext = { sheetOrder: [] },
): RebaseResult {
  if (committed.kind === 'unknown') {
    rebaseConflict(`committed ${committed.mutationId} has no registered structural transform`);
  }
  const delta = extractStructuralDelta(committed);
  if (!delta) {
    if (STRUCTURAL_KINDS.has(committed.kind)) rebaseConflict(`committed ${committed.mutationId} has no structural bounds`);
    if (UNSUPPORTED_STRUCTURAL_KINDS.has(committed.kind)) {
      rebaseConflict(`committed ${committed.mutationId} has no canonical structural patch`);
    }
    return { rebased: pending, transformed: false };
  }
  if (!STRUCTURAL_KINDS.has(committed.kind)) {
    return { rebased: pending, transformed: false };
  }
  if (pending.kind === 'unknown') {
    rebaseConflict(`Cannot rebase unknown mutation ${pending.mutationId} across ${committed.mutationId}`);
  }
  if (UNSUPPORTED_STRUCTURAL_KINDS.has(pending.kind)) {
    rebaseConflict(`pending ${pending.mutationId} has no canonical structural patch for ${committed.mutationId}`);
  }

  const rebasedRanges = pending.affectedRanges.map((range) => {
    return shiftRange(range, delta);
  });

  const transformedParams = transformParams(pending.params, delta, pending.sheetId, pending.kind, context, '', true);
  const rebasedParams = pending.mutationId === 'range.paste'
    ? transformPasteSnapshots(pending.params, transformedParams, pending.sheetId, delta)
    : transformedParams;

  return {
    rebased: { ...pending, affectedRanges: rebasedRanges, params: rebasedParams },
    transformed: true,
  };
}

/** 按 revision 顺序依次 rebase 一批 pending 操作 */
export function rebaseAgainstHistory(
  pending: ClassifiedMutation,
  committedHistory: ClassifiedMutation[],
  context: StructuralRebaseContext = { sheetOrder: [] },
): RebaseResult {
  let current = pending;
  let transformed = false;
  for (const committed of committedHistory) {
    const result = rebaseMutation(current, committed, context);
    current = result.rebased;
    transformed = transformed || result.transformed;
  }
  return { rebased: current, transformed };
}
