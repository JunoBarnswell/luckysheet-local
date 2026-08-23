import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import type {
  ApiError,
  CommittedOperationEnvelopeV2,
  OperationEnvelopeV2,
} from '@react-sheets/protocol';
import { classifyMutation, committedMutationToClassified } from './operation-types';
import { rebaseAgainstHistory } from './ot-rebase';
import { OfflineQueue } from './offline-queue';
import { CollaborativeUndoStack } from './collaborative-undo';
import { PresenceStore } from './presence';

export interface CollaborationSessionOptions {
  /** Sends an envelope over the authenticated V2 transport. */
  send?: (operation: OperationEnvelopeV2) => boolean;
  createOperationId?: () => string;
  /** Durable operation journal; workbook snapshots are not stored here. */
  loadPending?: () => { operations: readonly OperationEnvelopeV2[]; nextClientSequence: number } | null;
  persistPending?: (operations: readonly OperationEnvelopeV2[], nextClientSequence: number) => void;
}

interface AckWaiter {
  resolve: (revision: number) => void;
  reject: (cause: unknown) => void;
}

const LOCAL_UNDO_KEY = 'local';

/** 协同会话 — V2 operation envelope + OT rebase + ACK-gated offline queue. */
export class CollaborationSession {
  readonly presence = new PresenceStore();
  readonly offlineQueue: OfflineQueue;
  readonly collaborativeUndo = new CollaborativeUndoStack();

  private runtime: CommandRuntime;
  private send?: (operation: OperationEnvelopeV2) => boolean;
  private readonly createOperationId: () => string;
  private clientSequence = 0;
  private baseRevision = 0;
  private readonly committedMutations: ReturnType<typeof classifyMutation>[] = [];
  private readonly committedOperationIds = new Set<string>();
  private readonly localClassified = new Map<string, ReturnType<typeof classifyMutation>[]>();
  private readonly ackWaiters = new Map<string, AckWaiter>();

  constructor(runtime: CommandRuntime, options: CollaborationSessionOptions = {}) {
    this.runtime = runtime;
    this.send = options.send;
    this.createOperationId = options.createOperationId ?? (() => {
      if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
      return `op-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    });
    const restored = options.loadPending?.() ?? null;
    this.clientSequence = Math.max(
      restored?.nextClientSequence ?? 0,
      ...(restored?.operations.map((operation) => operation.clientSequence) ?? []),
    );
    this.offlineQueue = new OfflineQueue({
      load: () => restored?.operations ?? [],
      persist: (operations) => options.persistPending?.(operations, this.clientSequence),
      flush: (operation) => this.flushOperation(operation),
    });
    for (const operation of restored?.operations ?? []) {
      this.localClassified.set(operation.operationId, this.classifyEnvelope(operation));
    }
    this.offlineQueue.setOnline(false);
  }

  rebindCommands(runtime: CommandRuntime): void {
    this.runtime = runtime;
  }

  attachTransport(send?: (operation: OperationEnvelopeV2) => boolean): void {
    this.send = send;
  }

  setRevision(revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('Revision must be a non-negative safe integer');
    this.baseRevision = Math.max(this.baseRevision, revision);
  }

  getRevision(): number {
    return this.baseRevision;
  }

  /** 本地命令执行后 enqueue V2 协同变更；ranges 只用于本地 OT，不进入 wire。 */
  enqueueLocalMutations(mutations: MutationInfo[], unitId: string, operationId = this.createOperationId()): OperationEnvelopeV2 {
    if (!unitId.trim()) throw new Error('unitId is required');
    if (mutations.length === 0) throw new Error('At least one mutation is required');
    this.clientSequence += 1;
    const operation: OperationEnvelopeV2 = {
      schema: 'OperationEnvelopeV2',
      operationId,
      unitId,
      clientSequence: this.clientSequence,
      baseRevision: this.baseRevision,
      mutations: mutations.map(({ id, sheetId, params }) => ({ id, sheetId, params })),
      createdAt: new Date().toISOString(),
    };
    this.localClassified.set(operationId, mutations.map((mutation) => classifyMutation(
      mutation.id,
      mutation.params,
      mutation.sheetId,
      [...mutation.affectedRanges],
    )));
    this.offlineQueue.enqueue(operation);
    return operation;
  }

  /** 应用远端已提交 V2 operation — 不进本地撤销栈。 */
  applyRemote(operation: CommittedOperationEnvelopeV2): void {
    if (operation.unitId !== this.runtime.workbook.unitId) throw new Error('Remote operation belongs to another workbook');
    if (!Number.isSafeInteger(operation.revision) || operation.revision < 1) throw new Error('Remote operation revision is invalid');
    this.runtime.applyRemoteMutations(operation.mutations.map((mutation) => ({
      id: mutation.id,
      unitId: operation.unitId,
      sheetId: mutation.sheetId,
      params: mutation.params,
      affectedRanges: [...mutation.affectedRanges],
    })));
    for (const mutation of operation.mutations) this.committedMutations.push(committedMutationToClassified(mutation));
    this.baseRevision = Math.max(this.baseRevision, operation.revision);
  }

  /** ACK is the only normal path that removes an operation from the queue. */
  acknowledge(operationId: string, revision: number): boolean {
    const local = this.localClassified.get(operationId);
    if (local) this.committedMutations.push(...local);
    this.localClassified.delete(operationId);
    this.baseRevision = Math.max(this.baseRevision, revision);
    const removed = this.offlineQueue.acknowledge(operationId);
    this.ackWaiters.get(operationId)?.resolve(revision);
    this.ackWaiters.delete(operationId);
    return removed;
  }

  /** Rejection preserves the queued operation and makes the failure visible. */
  reject(operationId: string, error: ApiError | Error): boolean {
    const cause = error instanceof Error ? error : new Error(error.message);
    const kept = this.offlineQueue.reject(operationId, cause);
    this.ackWaiters.get(operationId)?.reject(cause);
    this.ackWaiters.delete(operationId);
    return kept;
  }

  /** Reject in-flight waits on a transport close; queue items remain pending. */
  transportClosed(cause = new Error('Collaboration socket unavailable')): void {
    for (const [operationId, waiter] of this.ackWaiters) {
      this.ackWaiters.delete(operationId);
      waiter.reject(cause);
    }
  }

  /** rebase pending local operation against committed structural history. */
  rebasePending(mutationId: string, params: unknown, sheetId: string, affectedRanges: MutationInfo['affectedRanges']) {
    const pending = classifyMutation(mutationId, params, sheetId, [...affectedRanges]);
    return rebaseAgainstHistory(pending, this.committedMutations);
  }

  recordCommittedMutations(mutations: Array<{ id: string; params: unknown; sheetId: string; affectedRanges: MutationInfo['affectedRanges'] }>): void {
    for (const mutation of mutations) {
      this.committedMutations.push(classifyMutation(mutation.id, mutation.params, mutation.sheetId, [...mutation.affectedRanges]));
    }
  }

  recordLocalUndo(entry: { operationId: string; undoMutations: MutationInfo[] }): void {
    this.collaborativeUndo.push(LOCAL_UNDO_KEY, {
      operationId: entry.operationId,
      actorId: LOCAL_UNDO_KEY,
      undoMutations: entry.undoMutations,
      timestamp: Date.now(),
    });
  }

  undoOwnLast(): MutationInfo[] | undefined {
    const entry = this.collaborativeUndo.pop(LOCAL_UNDO_KEY);
    return entry ? this.collaborativeUndo.createCompensatingCommand(entry) : undefined;
  }

  private flushOperation(operation: OperationEnvelopeV2): Promise<number> {
    return new Promise<number>((resolve, reject) => {
      this.ackWaiters.set(operation.operationId, { resolve, reject });
      if (!this.send?.(operation)) {
        this.ackWaiters.delete(operation.operationId);
        reject(new Error('Collaboration socket unavailable'));
      }
    });
  }
}
