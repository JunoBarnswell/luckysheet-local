import type { MutationInfo } from '@react-sheets/command-runtime';
import type { PivotModel, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import { getPivotSourceRanges } from './engine';

export type PivotRefreshTrigger =
  | { kind: 'open' }
  | { kind: 'explicit'; pivotId: string }
  | { kind: 'explicit-all' }
  | { kind: 'layout-change'; pivotId: string }
  | { kind: 'source-change'; mutations: readonly MutationInfo[] }
  | { kind: 'source-content-change'; sourceId: string };

function intersects(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow
    && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn
    && right.startColumn <= left.endColumn;
}

function dependsOnMutation(workbook: WorkbookModel, pivot: PivotModel, mutation: MutationInfo): boolean {
  if (mutation.id.startsWith('pivot.') || mutation.id.startsWith('drawing.')) return false;
  try {
    return getPivotSourceRanges(workbook, pivot).some((sourceRange) =>
      mutation.affectedRanges.some((affectedRange) => intersects(sourceRange, affectedRange)));
  } catch {
    // An invalid source cannot be guessed into a refresh target. The normal
    // projection/command validation will expose the source error instead.
    return false;
  }
}

/** Pure policy gate shared by local, remote, and replay-triggered refreshes. */
export function pivotIdsToRefresh(
  workbook: WorkbookModel,
  pivots: readonly PivotModel[],
  trigger: PivotRefreshTrigger,
): string[] {
  switch (trigger.kind) {
    case 'explicit':
      return pivots.some((pivot) => pivot.id === trigger.pivotId) ? [trigger.pivotId] : [];
    case 'explicit-all':
      return pivots.map((pivot) => pivot.id);
    case 'layout-change':
      return pivots.some((pivot) => pivot.id === trigger.pivotId) ? [trigger.pivotId] : [];
    case 'open':
      return pivots.filter((pivot) => pivot.refreshPolicy.mode === 'on-open').map((pivot) => pivot.id);
    case 'source-change':
      return pivots
        .filter((pivot) => pivot.refreshPolicy.mode === 'on-change')
        .filter((pivot) => trigger.mutations.some((mutation) => dependsOnMutation(workbook, pivot, mutation)))
        .map((pivot) => pivot.id);
    case 'source-content-change':
      return pivots
        .filter((pivot) => pivot.refreshPolicy.mode === 'on-change')
        .filter((pivot) => pivot.source.kind === 'data-source' && pivot.source.dataSourceId === trigger.sourceId)
        .map((pivot) => pivot.id);
  }
}
