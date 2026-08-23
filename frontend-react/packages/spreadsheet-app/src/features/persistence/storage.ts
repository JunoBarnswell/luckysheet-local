import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { OperationEnvelopeV2 } from '@react-sheets/protocol';
import { computeSnapshotChecksum, verifySnapshotChecksum } from '../../../../storage/src/checksum';

const DRAFT_KEY_PREFIX = 'react-sheets:draft:';
const OPERATION_KEY_PREFIX = 'react-sheets:pending-operations:';
const memoryDrafts = new Map<string, string>();
const memoryOperationJournals = new Map<string, string>();

export interface LocalDraftRecord {
  unitId: string;
  revision: number;
  checksum: string;
  snapshot: WorkbookSnapshotV1;
  updatedAt: string;
}

/** Durable offline journal. It stores operation intent, never a workbook
 * snapshot. Sent entries are restored as pending so a lost ACK is retried by
 * operationId and cannot silently lose a local edit. */
export interface PendingOperationJournal {
  schema: 'PendingOperationJournalV1';
  unitId: string;
  nextClientSequence: number;
  operations: OperationEnvelopeV2[];
  checksum: string;
  updatedAt: string;
}

function journalPayload(
  unitId: string,
  nextClientSequence: number,
  operations: readonly OperationEnvelopeV2[],
): Omit<PendingOperationJournal, 'checksum' | 'updatedAt'> {
  return {
    schema: 'PendingOperationJournalV1',
    unitId,
    nextClientSequence,
    operations: operations.map((operation) => structuredClone(operation)),
  };
}

function journalChecksum(payload: Omit<PendingOperationJournal, 'checksum' | 'updatedAt'>): string {
  return computeSnapshotChecksum(JSON.stringify(payload));
}

export class LocalOperationStore {
  write(unitId: string, operations: readonly OperationEnvelopeV2[], nextClientSequence: number): void {
    if (!unitId.trim()) throw new Error('unitId is required');
    if (!Number.isSafeInteger(nextClientSequence) || nextClientSequence < 0) {
      throw new Error('nextClientSequence must be a non-negative safe integer');
    }
    const payload = journalPayload(unitId, nextClientSequence, operations);
    const record: PendingOperationJournal = {
      ...payload,
      checksum: journalChecksum(payload),
      updatedAt: new Date().toISOString(),
    };
    const encoded = JSON.stringify(record);
    if (typeof window !== 'undefined') {
      window.localStorage.setItem(`${OPERATION_KEY_PREFIX}${unitId}`, encoded);
    } else {
      memoryOperationJournals.set(unitId, encoded);
    }
  }

  read(unitId: string): PendingOperationJournal | null {
    const raw = typeof window !== 'undefined'
      ? window.localStorage.getItem(`${OPERATION_KEY_PREFIX}${unitId}`)
      : memoryOperationJournals.get(unitId);
    if (!raw) return null;
    try {
      const record = JSON.parse(raw) as PendingOperationJournal;
      if (record.schema !== 'PendingOperationJournalV1' || record.unitId !== unitId) return null;
      if (!Number.isSafeInteger(record.nextClientSequence) || record.nextClientSequence < 0) return null;
      if (!Array.isArray(record.operations) || !record.checksum || !record.updatedAt) return null;
      const payload = journalPayload(record.unitId, record.nextClientSequence, record.operations);
      if (journalChecksum(payload) !== record.checksum) return null;
      const seen = new Set<string>();
      let previousSequence = 0;
      for (const operation of record.operations) {
        if (operation.schema !== 'OperationEnvelopeV2' || operation.unitId !== unitId) return null;
        if (!operation.operationId || seen.has(operation.operationId)) return null;
        if (!Number.isSafeInteger(operation.clientSequence) || operation.clientSequence <= previousSequence) return null;
        if (operation.clientSequence > record.nextClientSequence) return null;
        seen.add(operation.operationId);
        previousSequence = operation.clientSequence;
      }
      return {
        schema: 'PendingOperationJournalV1',
        unitId,
        nextClientSequence: record.nextClientSequence,
        operations: record.operations.map((operation) => structuredClone(operation)),
        checksum: record.checksum,
        updatedAt: record.updatedAt,
      };
    } catch {
      return null;
    }
  }

  clear(unitId: string): void {
    if (typeof window !== 'undefined') {
      window.localStorage.removeItem(`${OPERATION_KEY_PREFIX}${unitId}`);
    } else {
      memoryOperationJournals.delete(unitId);
    }
  }
}

export interface PersistenceSnapshotMeta {
  unitId: string;
  revision: number;
  checksum: string;
  updatedAt: string;
  hasLocalDraft: boolean;
  draftUpdatedAt?: string;
}

export function buildPersistenceMeta(
  snapshot: WorkbookSnapshotV1,
  revision: number,
  draft?: LocalDraftRecord | null,
): PersistenceSnapshotMeta {
  const snapshotJson = JSON.stringify(snapshot);
  return {
    unitId: snapshot.unitId,
    revision,
    checksum: computeSnapshotChecksum(snapshotJson),
    updatedAt: new Date().toISOString(),
    hasLocalDraft: Boolean(draft),
    draftUpdatedAt: draft?.updatedAt,
  };
}

export function buildLocalDraftRecord(
  snapshot: WorkbookSnapshotV1,
  revision: number,
): LocalDraftRecord {
  const snapshotJson = JSON.stringify(snapshot);
  return {
    unitId: snapshot.unitId,
    revision,
    checksum: computeSnapshotChecksum(snapshotJson),
    snapshot: structuredClone(snapshot),
    updatedAt: new Date().toISOString(),
  };
}

export function verifyLocalDraft(record: LocalDraftRecord): boolean {
  return verifySnapshotChecksum(JSON.stringify(record.snapshot), record.checksum);
}

export function isDraftNewerThanServer(draft: LocalDraftRecord, serverRevision: number): boolean {
  return draft.revision > serverRevision;
}

export class LocalDraftStore {
  write(record: LocalDraftRecord): void {
    const payload = JSON.stringify(record);
    if (typeof window !== 'undefined') {
      window.localStorage.setItem(`${DRAFT_KEY_PREFIX}${record.unitId}`, payload);
      return;
    }
    memoryDrafts.set(record.unitId, payload);
  }

  read(unitId: string): LocalDraftRecord | null {
    const raw = typeof window !== 'undefined'
      ? window.localStorage.getItem(`${DRAFT_KEY_PREFIX}${unitId}`)
      : memoryDrafts.get(unitId);
    if (!raw) return null;
    try {
      const record = JSON.parse(raw) as LocalDraftRecord;
      if (record.unitId !== unitId || !record.snapshot || !record.checksum) return null;
      if (!verifyLocalDraft(record)) return null;
      return record;
    } catch {
      return null;
    }
  }

  clear(unitId: string): void {
    if (typeof window !== 'undefined') {
      window.localStorage.removeItem(`${DRAFT_KEY_PREFIX}${unitId}`);
      return;
    }
    memoryDrafts.delete(unitId);
  }
}

export function scheduleDebounced<T extends (...args: never[]) => void>(fn: T, delayMs: number): T {
  let timer: ReturnType<typeof setTimeout> | null = null;
  return ((...args: Parameters<T>) => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      fn(...args);
    }, delayMs);
  }) as T;
}
