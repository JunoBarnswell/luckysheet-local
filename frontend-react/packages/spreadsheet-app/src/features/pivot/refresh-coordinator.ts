import type { MutationInfo } from '@react-sheets/command-runtime';
import type { PivotModel, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import { getPivotSourceRanges } from './engine';

export type PivotRefreshTrigger =
  | { kind: 'open'; sheetId?: string }
  | { kind: 'explicit'; pivotId: string }
  | { kind: 'explicit-all' }
  | { kind: 'layout-change'; pivotId: string }
  | { kind: 'source-change'; mutations: readonly MutationInfo[]; sheetId?: string }
  | { kind: 'source-content-change'; sourceId: string; sheetId?: string }
  | {
    kind: 'control-change';
    drawings: readonly { sheetId: string; drawingId: string }[];
    /** Links captured on a remove mutation, after the canonical drawing is gone. */
    pivotIds?: readonly string[];
  };

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

function linkedPivotIdsForDrawing(
  workbook: WorkbookModel,
  reference: { sheetId: string; drawingId: string },
): string[] {
  const sheet = workbook.getSheets().find((candidate) => candidate.id === reference.sheetId);
  const drawing = sheet?.drawings.find((candidate) => candidate.id === reference.drawingId);
  if (!sheet || !drawing || (drawing.kind !== 'slicer' && drawing.kind !== 'timeline')) return [];
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) return [];
  return [...new Set([payload.pivotId, ...(payload.connections ?? []).map((connection) => connection.pivotId)])];
}

/** Pure policy gate shared by local, remote, and replay-triggered refreshes. */
export function pivotIdsToRefresh(
  workbook: WorkbookModel,
  pivots: readonly PivotModel[],
  trigger: PivotRefreshTrigger,
): string[] {
  const targetSheetMatches = (pivot: PivotModel, sheetId?: string): boolean => !sheetId || pivot.target.sheetId === sheetId;
  switch (trigger.kind) {
    case 'explicit':
      return pivots.some((pivot) => pivot.id === trigger.pivotId) ? [trigger.pivotId] : [];
    case 'explicit-all':
      return pivots.map((pivot) => pivot.id);
    case 'layout-change':
      return pivots.some((pivot) => pivot.id === trigger.pivotId) ? [trigger.pivotId] : [];
    case 'open':
      return pivots
        .filter((pivot) => targetSheetMatches(pivot, trigger.sheetId))
        .filter((pivot) => pivot.refreshPolicy.mode === 'on-open')
        .map((pivot) => pivot.id);
    case 'source-change':
      return pivots
        .filter((pivot) => targetSheetMatches(pivot, trigger.sheetId))
        .filter((pivot) => pivot.refreshPolicy.mode === 'on-change')
        .filter((pivot) => trigger.mutations.some((mutation) => dependsOnMutation(workbook, pivot, mutation)))
        .map((pivot) => pivot.id);
    case 'source-content-change':
      return pivots
        .filter((pivot) => targetSheetMatches(pivot, trigger.sheetId))
        .filter((pivot) => pivot.refreshPolicy.mode === 'on-change')
        .filter((pivot) => pivot.source.kind === 'data-source' && pivot.source.dataSourceId === trigger.sourceId)
        .map((pivot) => pivot.id);
    case 'control-change': {
      const linkedPivotIds = new Set(trigger.pivotIds ?? []);
      for (const drawing of trigger.drawings) {
        for (const pivotId of linkedPivotIdsForDrawing(workbook, drawing)) linkedPivotIds.add(pivotId);
      }
      return pivots.filter((pivot) => linkedPivotIds.has(pivot.id)).map((pivot) => pivot.id);
    }
  }
}
