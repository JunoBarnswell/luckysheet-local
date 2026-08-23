import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { RevisionRecord } from '@react-sheets/protocol';
import { DrawingRuntime } from './features/drawing';
import { registerSpreadsheetFeatures } from './feature-registry';
import type { HistoryEntryMeta, RestoreCommandParams } from './features/history';

export function describeRevisionMutations(record: RevisionRecord): string {
  const labels = record.payload.mutations.map((mutation) => mutation.id);
  if (labels.length === 0) return 'Workbook metadata';
  const preview = labels.slice(0, 3).join(' · ');
  return labels.length > 3 ? `${preview} · +${labels.length - 3} more` : preview;
}

export function revisionToHistoryMeta(record: RevisionRecord): HistoryEntryMeta {
  return {
    revision: record.revision,
    operationId: record.operationId,
    actorId: record.payload.actorId,
    category: 'collaboration',
    description: describeRevisionMutations(record),
    createdAt: record.createdAt,
  };
}

export function buildRestoreParams(
  snapshot: WorkbookSnapshotV1,
  targetRevision: number,
  reason?: string,
): RestoreCommandParams {
  return {
    targetRevision,
    snapshot: structuredClone(snapshot),
    reason,
  };
}

export function replayRevisionsToSnapshot(
  baseSnapshot: WorkbookSnapshotV1,
  revisions: readonly RevisionRecord[],
  targetRevision: number,
): WorkbookSnapshotV1 {
  const workbook = WorkbookModel.fromSnapshot(structuredClone(baseSnapshot));
  const runtime = new CommandRuntime(workbook);
  registerSpreadsheetFeatures(runtime, new DrawingRuntime());

  const ordered = revisions
    .filter((record) => record.revision > 0 && record.revision <= targetRevision)
    .sort((left, right) => left.revision - right.revision);

  for (const record of ordered) {
    runtime.applyRemoteMutations(record.payload.mutations.map((mutation) => ({
      id: mutation.id,
      unitId: record.payload.unitId,
      sheetId: mutation.sheetId,
      params: mutation.params,
      affectedRanges: mutation.affectedRanges,
    })));
  }

  return workbook.snapshot();
}
