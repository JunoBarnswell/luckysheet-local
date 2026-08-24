import type { PivotHitTest, RangeRef } from '@react-sheets/core-model';

export type ContextTargetKind = 'editor' | 'drag-handle' | 'floating' | 'pivot' | 'table' | 'filter' | 'grid' | 'row-header' | 'column-header' | 'none';

export interface ResolvedContextHit {
  kind: ContextTargetKind;
  priority: number;
  sheetId: string;
  row?: number;
  column?: number;
  range?: RangeRef;
  objectId?: string;
  pivot?: PivotHitTest;
}

export interface ContextHitInput {
  sheetId: string;
  editor?: boolean;
  dragHandle?: boolean;
  floatingId?: string;
  pivot?: PivotHitTest;
  tableId?: string;
  filterColumn?: number;
  header?: 'row' | 'column';
  cell?: { row: number; column: number; range?: RangeRef };
}

/**
 * One deterministic context resolver for pointer, right-click and contextual
 * Ribbon selection. Higher-priority objects always win over ordinary cells.
 */
export function resolveContextHit(input: ContextHitInput): ResolvedContextHit {
  if (input.editor) return { kind: 'editor', priority: 100, sheetId: input.sheetId };
  if (input.dragHandle) return { kind: 'drag-handle', priority: 90, sheetId: input.sheetId };
  if (input.floatingId) return { kind: 'floating', priority: 80, sheetId: input.sheetId, objectId: input.floatingId };
  if (input.pivot && input.pivot.kind !== 'none') {
    return {
      kind: 'pivot',
      priority: 70,
      sheetId: input.sheetId,
      row: input.pivot.row,
      column: input.pivot.column,
      objectId: input.pivot.pivotId,
      pivot: input.pivot,
    };
  }
  if (input.header === 'row') return { kind: 'row-header', priority: 40, sheetId: input.sheetId, ...input.cell };
  if (input.header === 'column') return { kind: 'column-header', priority: 40, sheetId: input.sheetId, ...input.cell };
  if (input.tableId) return { kind: 'table', priority: 60, sheetId: input.sheetId, objectId: input.tableId, ...input.cell };
  if (input.filterColumn !== undefined) return { kind: 'filter', priority: 50, sheetId: input.sheetId, column: input.filterColumn, ...input.cell };
  if (input.cell) return { kind: 'grid', priority: 10, sheetId: input.sheetId, ...input.cell };
  return { kind: 'none', priority: 0, sheetId: input.sheetId };
}
