import type { RevisionRecord } from '@react-sheets/protocol';
import type { HistoryEntryMeta, RestoreCommandParams } from './index';

export function describeRevisionMutations(record: RevisionRecord): string {
  const labels = record.payload.mutations.map((mutation) => mutation.id);
  if (labels.length === 0) return 'Workbook metadata';
  const preview = labels.slice(0, 3).join(' · ');
  return labels.length > 3 ? `${preview} · +${labels.length - 3} more` : preview;
}

export function revisionToHistoryMeta(record: RevisionRecord): HistoryEntryMeta {
  return { revision: record.revision, operationId: record.operationId, actorId: record.payload.actorId, category: 'collaboration', description: describeRevisionMutations(record), createdAt: record.createdAt };
}

export function buildRestoreParams(targetRevision: number, reason?: string): RestoreCommandParams {
  if (!Number.isSafeInteger(targetRevision) || targetRevision < 0) throw new Error('targetRevision must be a non-negative integer');
  return { targetRevision, reason };
}

/** Server materializes history; the browser never replays revisions into a local snapshot. */
export function assertCommittedRevision(records: readonly RevisionRecord[], targetRevision: number): void {
  if (!Number.isSafeInteger(targetRevision) || targetRevision < 0) throw new Error('targetRevision must be a non-negative integer');
  const ordered = records.filter((record) => record.revision <= targetRevision).sort((a, b) => a.revision - b.revision);
  for (let index = 0; index < ordered.length; index += 1) {
    if (ordered[index]!.revision !== index + 1) throw new Error(`HISTORY_GAP: missing committed revision at ${index + 1}`);
  }
}
