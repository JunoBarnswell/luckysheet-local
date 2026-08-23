import { deflateRawSync, inflateRawSync } from 'node:zlib';
import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { CollaborationChangeSet } from '@react-sheets/protocol';
import { computeSnapshotChecksum } from './checksum';

export interface PersistenceSnapshot {
  revision: number;
  snapshot: WorkbookSnapshotV1;
  checksum: string;
  compressed: boolean;
  createdAt: string;
}

export interface PersistenceSessionOptions {
  /** 累积 changeset 条数阈值 */
  changesetThreshold?: number;
  /** 累积 payload 字节阈值 */
  byteThreshold?: number;
  /** 时间阈值 ms */
  timeThresholdMs?: number;
  persistSnapshot?: (record: PersistenceSnapshot) => Promise<void> | void;
  appendChangeSet?: (changeSet: CollaborationChangeSet) => Promise<void> | void;
}

/** append-only changeset + 周期 snapshot + checksum */
export class PersistenceSession {
  private revision = 0;
  private pendingChangesets = 0;
  private pendingBytes = 0;
  private lastSnapshotAt = Date.now();
  private lastSnapshot: WorkbookSnapshotV1 | null = null;

  private readonly changesetThreshold: number;
  private readonly byteThreshold: number;
  private readonly timeThresholdMs: number;
  private readonly persistSnapshot?: PersistenceSessionOptions['persistSnapshot'];
  private readonly appendChangeSet?: PersistenceSessionOptions['appendChangeSet'];

  constructor(options: PersistenceSessionOptions = {}) {
    this.changesetThreshold = options.changesetThreshold ?? 50;
    this.byteThreshold = options.byteThreshold ?? 512_000;
    this.timeThresholdMs = options.timeThresholdMs ?? 60_000;
    this.persistSnapshot = options.persistSnapshot;
    this.appendChangeSet = options.appendChangeSet;
  }

  getRevision(): number {
    return this.revision;
  }

  async recordChangeSet(changeSet: CollaborationChangeSet): Promise<void> {
    const payloadSize = JSON.stringify(changeSet).length;
    this.pendingChangesets += 1;
    this.pendingBytes += payloadSize;
    this.revision = Math.max(this.revision, changeSet.baseRevision + 1);
    await this.appendChangeSet?.(changeSet);
    await this.maybeSnapshot();
  }

  async writeSnapshot(snapshot: WorkbookSnapshotV1, revision?: number): Promise<PersistenceSnapshot> {
    if (revision != null) this.revision = revision;
    this.lastSnapshot = snapshot;
    const snapshotJson = JSON.stringify(snapshot);
    const checksum = computeSnapshotChecksum(snapshotJson);
    const record: PersistenceSnapshot = {
      revision: this.revision,
      snapshot,
      checksum,
      compressed: true,
      createdAt: new Date().toISOString(),
    };
    await this.persistSnapshot?.(record);
    this.pendingChangesets = 0;
    this.pendingBytes = 0;
    this.lastSnapshotAt = Date.now();
    return record;
  }

  encodeSnapshot(record: PersistenceSnapshot): Buffer {
    const json = JSON.stringify(record.snapshot);
    return deflateRawSync(Buffer.from(json, 'utf8'));
  }

  decodeSnapshot(payload: Buffer, meta: Omit<PersistenceSnapshot, 'snapshot'>): WorkbookSnapshotV1 {
    const json = inflateRawSync(payload).toString('utf8');
    const checksum = computeSnapshotChecksum(json);
    if (checksum !== meta.checksum) {
      throw new Error(`Snapshot checksum mismatch at revision ${meta.revision}`);
    }
    return JSON.parse(json) as WorkbookSnapshotV1;
  }

  shouldSnapshot(): boolean {
    return this.pendingChangesets >= this.changesetThreshold
      || this.pendingBytes >= this.byteThreshold
      || Date.now() - this.lastSnapshotAt >= this.timeThresholdMs;
  }

  getLastSnapshot(): WorkbookSnapshotV1 | null {
    return this.lastSnapshot;
  }

  private async maybeSnapshot(): Promise<void> {
    if (!this.shouldSnapshot() || !this.lastSnapshot) return;
    await this.writeSnapshot(this.lastSnapshot, this.revision);
  }
}

export { computeSnapshotChecksum, verifySnapshotChecksum } from './checksum';
