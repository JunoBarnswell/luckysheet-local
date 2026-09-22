import type { RangeRef } from '@react-sheets/core-model';
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

function shiftRow(range: RangeRef, delta: StructuralDelta): RangeRef {
  if (range.sheetId !== delta.sheetId) return range;
  let { startRow, endRow } = range;
  if (delta.kind === 'insert-rows' && delta.at <= startRow) {
    startRow += delta.count;
    endRow += delta.count;
  } else if (delta.kind === 'delete-rows') {
    const deletedEnd = delta.at + delta.count - 1;
    if (delta.at <= startRow) {
      const shift = Math.min(delta.count, startRow - delta.at);
      startRow -= shift;
      endRow -= shift;
    }
    if (startRow <= deletedEnd && endRow >= delta.at) {
      endRow = Math.max(delta.at - 1, startRow);
    }
  }
  return { ...range, startRow, endRow };
}

function shiftColumn(range: RangeRef, delta: StructuralDelta): RangeRef {
  if (range.sheetId !== delta.sheetId) return range;
  let { startColumn, endColumn } = range;
  if (delta.kind === 'insert-columns' && delta.at <= startColumn) {
    startColumn += delta.count;
    endColumn += delta.count;
  } else if (delta.kind === 'delete-columns') {
    const deletedEnd = delta.at + delta.count - 1;
    if (delta.at <= startColumn) {
      const shift = Math.min(delta.count, startColumn - delta.at);
      startColumn -= shift;
      endColumn -= shift;
    }
    if (startColumn <= deletedEnd && endColumn >= delta.at) {
      endColumn = Math.max(delta.at - 1, startColumn);
    }
  }
  return { ...range, startColumn, endColumn };
}

function extractStructuralDelta(mutation: ClassifiedMutation): StructuralDelta | undefined {
  const p = mutation.params as { at?: number; count?: number; row?: number; column?: number } | null;
  if (!p) return undefined;
  const at = p.at ?? p.row ?? p.column;
  const count = p.count ?? 1;
  if (at == null) return undefined;

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
export function rebaseMutation(pending: ClassifiedMutation, committed: ClassifiedMutation): RebaseResult {
  const delta = extractStructuralDelta(committed);
  if (!delta || !STRUCTURAL_KINDS.has(committed.kind)) {
    return { rebased: pending, transformed: false };
  }
  if (pending.kind === 'unknown') {
    throw new Error(`Cannot rebase unknown mutation ${pending.mutationId} across ${committed.mutationId}`);
  }

  const rebasedRanges = pending.affectedRanges.map((range) => {
    if (delta.kind === 'insert-rows' || delta.kind === 'delete-rows') return shiftRow(range, delta);
    return shiftColumn(range, delta);
  });

  const rebasedParams = rebaseParams(pending, delta);

  return {
    rebased: { ...pending, affectedRanges: rebasedRanges, params: rebasedParams },
    transformed: true,
  };
}

function rebaseParams(mutation: ClassifiedMutation, delta: StructuralDelta): unknown {
  const p = mutation.params as Record<string, unknown> | null;
  if (!p) return mutation.params;

  const next = { ...p };
  if (typeof p.row === 'number' && (delta.kind === 'insert-rows' || delta.kind === 'delete-rows')) {
    const shifted = shiftRow({
      sheetId: mutation.sheetId,
      startRow: p.row as number,
      endRow: p.row as number,
      startColumn: 0,
      endColumn: 0,
    }, delta);
    next.row = shifted.startRow;
  }
  if (typeof p.column === 'number' && (delta.kind === 'insert-columns' || delta.kind === 'delete-columns')) {
    const shifted = shiftColumn({
      sheetId: mutation.sheetId,
      startRow: 0,
      endRow: 0,
      startColumn: p.column as number,
      endColumn: p.column as number,
    }, delta);
    next.column = shifted.startColumn;
  }
  return next;
}

/** 按 revision 顺序依次 rebase 一批 pending 操作 */
export function rebaseAgainstHistory(
  pending: ClassifiedMutation,
  committedHistory: ClassifiedMutation[],
): RebaseResult {
  let current = pending;
  let transformed = false;
  for (const committed of committedHistory) {
    const result = rebaseMutation(current, committed);
    current = result.rebased;
    transformed = transformed || result.transformed;
  }
  return { rebased: current, transformed };
}
