import type { RangeRef } from '@react-sheets/core-model';
import type { CommittedOperationEnvelope, OperationEnvelope, OperationMutation } from '@react-sheets/protocol';

/** 协同 OT 操作分类 — 按类型做 transform/rebase */
export type CollaborationOperationKind =
  | 'cell-value'
  | 'cell-style'
  | 'clear'
  | 'insert-rows'
  | 'delete-rows'
  | 'insert-columns'
  | 'delete-columns'
  | 'move-range'
  | 'sort'
  | 'merge'
  | 'table-resize'
  | 'drawing'
  | 'comment'
  | 'pivot-config'
  | 'unknown';

export interface ClassifiedMutation {
  mutationId: string;
  kind: CollaborationOperationKind;
  sheetId: string;
  affectedRanges: RangeRef[];
  params: unknown;
}

export function operationMutationToClassified(
  mutation: OperationMutation,
  sheetId = mutation.sheetId,
  affectedRanges: RangeRef[] = [],
): ClassifiedMutation {
  return classifyMutation(mutation.id, mutation.params, sheetId, affectedRanges);
}

export function committedMutationToClassified(
  mutation: CommittedOperationEnvelope['mutations'][number],
): ClassifiedMutation {
  return operationMutationToClassified(mutation, mutation.sheetId, [...mutation.affectedRanges]);
}

const MUTATION_KIND_MAP: Readonly<Record<string, CollaborationOperationKind>> = {
  'cell.set': 'cell-value',
  'cell.restore': 'cell-value',
  'range.set': 'cell-value',
  'fill.applied': 'cell-value',
  'fill.restored': 'cell-value',
  'range.paste': 'cell-value',
  'range.clear': 'clear',
  'cells.inserted': 'move-range',
  'cells.deleted': 'move-range',
  'cells.inserted.restore': 'move-range',
  'cells.deleted.restore': 'move-range',
  'style.set': 'cell-style',
  'style.preset.set': 'cell-style',
  'format.painter.applied': 'cell-style',
  'row.insert': 'insert-rows',
  'row.delete': 'delete-rows',
  'column.insert': 'insert-columns',
  'column.delete': 'delete-columns',
  'rows.inserted': 'insert-rows',
  'rows.deleted': 'delete-rows',
  'columns.inserted': 'insert-columns',
  'columns.deleted': 'delete-columns',
  'range.move': 'move-range',
  'dataRegion.materialize.commit': 'cell-value',
  'dataRegion.materialize.restore': 'cell-value',
  'sort.apply': 'sort',
  'merge.set': 'merge',
  'table.resize': 'table-resize',
  'drawing.update': 'drawing',
  'drawing.visibility.set': 'drawing',
  'drawing.rename': 'drawing',
  'comment.add': 'comment',
  'comment.update': 'comment',
  'find.replaced': 'cell-value',
  'pivot.layout.set': 'pivot-config',
};

export function classifyMutation(mutationId: string, params: unknown, sheetId: string, affectedRanges: RangeRef[]): ClassifiedMutation {
  return {
    mutationId,
    kind: MUTATION_KIND_MAP[mutationId] ?? 'unknown',
    sheetId,
    affectedRanges,
    params,
  };
}
