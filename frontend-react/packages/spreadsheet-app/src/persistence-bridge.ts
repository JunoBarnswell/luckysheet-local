import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import { computeSnapshotChecksum, verifySnapshotChecksum } from '../../storage/src/checksum';

const DRAFT_KEY_PREFIX = 'react-sheets:draft:';
const memoryDrafts = new Map<string, string>();

export interface LocalDraftRecord {
  unitId: string;
  revision: number;
  checksum: string;
  snapshot: WorkbookSnapshotV1;
  updatedAt: string;
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
