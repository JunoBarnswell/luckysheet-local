import { deflateRawSync, inflateRawSync } from 'node:zlib';
import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { CommittedOperationEnvelopeV2 } from '@react-sheets/protocol';
import { computeSnapshotChecksum } from './checksum';

export interface PersistenceSnapshot {
  schema: 'SnapshotV2';
  schemaVersion: 2;
  unitId: string;
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
  appendOperation?: (operation: CommittedOperationEnvelopeV2) => Promise<void> | void;
}

interface SessionState {
  revision: number;
  pendingChangesets: number;
  pendingBytes: number;
  lastSnapshotAt: number;
  lastSnapshot: WorkbookSnapshotV1 | null;
}

/** append-only changeset + 周期 snapshot + checksum */
export class PersistenceSession {
  private readonly states = new Map<string, SessionState>();

  private readonly changesetThreshold: number;
  private readonly byteThreshold: number;
  private readonly timeThresholdMs: number;
  private readonly persistSnapshot?: PersistenceSessionOptions['persistSnapshot'];
  private readonly appendOperation?: PersistenceSessionOptions['appendOperation'];

  constructor(options: PersistenceSessionOptions = {}) {
    this.changesetThreshold = options.changesetThreshold ?? 50;
    this.byteThreshold = options.byteThreshold ?? 512_000;
    this.timeThresholdMs = options.timeThresholdMs ?? 60_000;
    this.persistSnapshot = options.persistSnapshot;
    this.appendOperation = options.appendOperation;
  }

  getRevision(unitId = 'default'): number {
    return this.stateFor(unitId).revision;
  }

  async recordOperation(operation: CommittedOperationEnvelopeV2): Promise<void> {
    const state = this.stateFor(operation.unitId);
    state.pendingChangesets += 1;
    state.pendingBytes += JSON.stringify(operation).length;
    state.revision = Math.max(state.revision, operation.revision);
    await this.appendOperation?.(operation);
    await this.maybeSnapshot(operation.unitId);
  }

  async writeSnapshot(snapshot: WorkbookSnapshotV1, revision?: number): Promise<PersistenceSnapshot> {
    const state = this.stateFor(snapshot.unitId);
    if (revision != null) state.revision = revision;
    state.lastSnapshot = structuredClone(snapshot);
    const snapshotJson = JSON.stringify(snapshot);
    const checksum = computeSnapshotChecksum(snapshotJson);
    const record: PersistenceSnapshot = {
      schema: 'SnapshotV2',
      schemaVersion: 2,
      unitId: snapshot.unitId,
      revision: state.revision,
      snapshot,
      checksum,
      compressed: true,
      createdAt: new Date().toISOString(),
    };
    await this.persistSnapshot?.(record);
    state.pendingChangesets = 0;
    state.pendingBytes = 0;
    state.lastSnapshotAt = Date.now();
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
    return this.shouldSnapshotFor('default');
  }

  getLastSnapshot(unitId = 'default'): WorkbookSnapshotV1 | null {
    return this.stateFor(unitId).lastSnapshot;
  }

  private shouldSnapshotFor(unitId: string): boolean {
    const state = this.stateFor(unitId);
    return state.pendingChangesets >= this.changesetThreshold
      || state.pendingBytes >= this.byteThreshold
      || Date.now() - state.lastSnapshotAt >= this.timeThresholdMs;
  }

  private async maybeSnapshot(unitId: string): Promise<void> {
    const state = this.stateFor(unitId);
    if (!this.shouldSnapshotFor(unitId) || !state.lastSnapshot) return;
    await this.writeSnapshot(state.lastSnapshot, state.revision);
  }

  private stateFor(unitId: string): SessionState {
    const existing = this.states.get(unitId);
    if (existing) return existing;
    const created: SessionState = {
      revision: 0,
      pendingChangesets: 0,
      pendingBytes: 0,
      lastSnapshotAt: Date.now(),
      lastSnapshot: null,
    };
    this.states.set(unitId, created);
    return created;
  }
}

export { computeSnapshotChecksum, verifySnapshotChecksum } from './checksum';
