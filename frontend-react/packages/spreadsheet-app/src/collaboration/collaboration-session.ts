import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { structuralRangeOwnerAffectedRanges } from '@react-sheets/core-model';
import type {
  ApiError,
  CommittedOperationEnvelope,
  OperationEnvelope,
  OperationIntent,
} from '@react-sheets/protocol';
import { requiresStructuralPatch, validateOperationEnvelope, validateStructuralPatch } from '@react-sheets/protocol';
import { classifyMutation, committedMutationToClassified } from './operation-types';
import { rebaseAgainstHistory } from './ot-rebase';
import { OfflineQueue } from './offline-queue';
import { PresenceStore } from './presence';

export interface CollaborationSessionOptions {
  clientSessionId?: string;
  /** Sends an operation over the authenticated REST transport. */
  send?: (operation: OperationEnvelope) => boolean | Promise<boolean | number>;
  createOperationId?: () => string;
  /** Durable operation journal; workbook snapshots are not stored here. */
  loadPending?: () => { operations: readonly OperationEnvelope[]; nextClientSequence: number } | null;
  persistPending?: (operations: readonly OperationEnvelope[], nextClientSequence: number) => void;
}

interface AckWaiter {
  resolve: (revision: number) => void;
  reject: (cause: unknown) => void;
}

interface CommittedOperationFact {
  readonly baseRevision: number;
  readonly revision: number;
}

interface QueueRebasePlan {
  readonly operations: readonly OperationEnvelope[];
  readonly classifications: ReadonlyMap<string, ReturnType<typeof classifyMutation>[]>;
  readonly remoteCount: number;
}

function assertStructuralPatch(value: unknown, mutationId: string): void {
  validateStructuralPatch(value, mutationId);
}

type StructuralImpactRange = NonNullable<CommittedOperationEnvelope['mutations'][number]['structuralImpactRanges']>[number];

function structuralPatchImpactRanges(
  patch: NonNullable<CommittedOperationEnvelope['mutations'][number]['structuralPatch']>,
): StructuralImpactRange[] {
  const ranges = new Map<string, StructuralImpactRange>();
  for (const delta of patch.formulaOwnerDeltas) {
    const deltaRanges: StructuralImpactRange[] = delta.kind === 'formula-cell' || delta.kind === 'formula-rule-anchor'
      ? [delta.beforeAddress, delta.afterAddress].flatMap((address) => address ? [{
        sheetId: address.sheetId,
        startRow: address.row,
        endRow: address.row,
        startColumn: address.column,
        endColumn: address.column,
      }] : [])
      : delta.kind === 'formula-rule' ? [...delta.beforeRanges, ...delta.afterRanges] : [];
    for (const range of deltaRanges) {
      ranges.set(JSON.stringify([range.sheetId, range.startRow, range.endRow, range.startColumn, range.endColumn]), range);
    }
  }
  for (const delta of patch.rangeOwnerDeltas) {
    for (const range of structuralRangeOwnerAffectedRanges(delta)) {
      ranges.set(JSON.stringify([range.sheetId, range.startRow, range.endRow, range.startColumn, range.endColumn]), { ...range });
    }
  }
  return [...ranges.values()];
}

function sameRange(left: StructuralImpactRange, right: StructuralImpactRange | undefined): boolean {
  return right !== undefined
    && left.sheetId === right.sheetId
    && left.startRow === right.startRow && left.endRow === right.endRow
    && left.startColumn === right.startColumn && left.endColumn === right.endColumn;
}

function committedMutationInfos(operation: CommittedOperationEnvelope): MutationInfo[] {
  return operation.mutations.map((mutation) => ({
    id: mutation.id,
    unitId: operation.unitId,
    sheetId: mutation.sheetId,
    params: mutation.params,
    affectedRanges: [...mutation.affectedRanges],
    ...(mutation.structuralImpactRanges?.length ? { structuralImpactRanges: [...mutation.structuralImpactRanges] } : {}),
    ...(mutation.structuralPatch ? {
      structuralFormulaOwnerDeltas: structuredClone(mutation.structuralPatch.formulaOwnerDeltas),
      structuralDefinedNameOwnerDeltas: structuredClone(mutation.structuralPatch.definedNameOwnerDeltas),
      structuralRangeOwnerDeltas: structuredClone(mutation.structuralPatch.rangeOwnerDeltas),
    } : {}),
  }));
}

/** 协同会话 — single operation envelope + OT rebase + ACK-gated offline queue. */
export class CollaborationSession {
  readonly presence = new PresenceStore();
  readonly offlineQueue: OfflineQueue;

  private runtime: CommandRuntime;
  private send?: (operation: OperationEnvelope) => boolean | Promise<boolean | number>;
  private readonly createOperationId: () => string;
  private clientSequence = 0;
  private readonly clientSessionId: string;
  private baseRevision = 0;
  private readonly committedMutations: ReturnType<typeof classifyMutation>[] = [];
  private readonly remoteMutations: ReturnType<typeof classifyMutation>[] = [];
  private rebasedRemoteCount = 0;
  private readonly committedOperations = new Map<string, CommittedOperationFact>();
  private readonly localClassified = new Map<string, ReturnType<typeof classifyMutation>[]>();
  private readonly ackWaiters = new Map<string, AckWaiter>();

  constructor(runtime: CommandRuntime, options: CollaborationSessionOptions = {}) {
    this.runtime = runtime;
    this.clientSessionId = options.clientSessionId ?? crypto.randomUUID();
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

  attachTransport(send?: (operation: OperationEnvelope) => boolean | Promise<boolean | number>): void {
    this.send = send;
  }

  setRevision(revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('Revision must be a non-negative safe integer');
    this.baseRevision = Math.max(this.baseRevision, revision);
    this.runtime.setRevision(this.baseRevision);
  }

  getRevision(): number {
    return this.baseRevision;
  }

  assertNextRevision(revision: number): void {
    const expected = this.baseRevision + 1;
    if (!Number.isSafeInteger(revision) || revision !== expected) {
      throw new Error(`COLLABORATION_REVISION_GAP: expected revision ${expected}, received ${revision}`);
    }
  }

  /** 本地命令执行后 enqueue operation；ranges 只用于本地 OT，不进入 wire。 */
  enqueueLocalMutations(
    mutations: MutationInfo[],
    unitId: string,
    operationId = this.createOperationId(),
    intent?: OperationIntent,
  ): OperationEnvelope {
    if (!unitId.trim()) throw new Error('unitId is required');
    if (mutations.length === 0) throw new Error('At least one mutation is required');
    this.assertMutationContracts(mutations);
    this.clientSequence += 1;
    const operation: OperationEnvelope = {
      schema: 'OperationEnvelope',
      clientSessionId: this.clientSessionId,
      operationId,
      unitId,
      clientSequence: this.clientSequence,
      baseRevision: this.baseRevision,
      mutations: mutations.map(({ id, sheetId, params }) => ({ id, sheetId, params })),
      createdAt: new Date().toISOString(),
      ...(intent ? { intent } : {}),
    };
    this.localClassified.set(operationId, mutations.map((mutation) => classifyMutation(
      mutation.id,
      mutation.params,
      mutation.sheetId,
      [...mutation.affectedRanges, ...(mutation.structuralImpactRanges ?? [])],
    )));
    this.offlineQueue.enqueue(operation);
    return operation;
  }

  /** Submit a compensating operation generated by a local collaborative undo. */
  enqueueCompensatingMutations(
    mutations: MutationInfo[],
    unitId: string,
    targetOperationId: string,
  ): OperationEnvelope {
    if (mutations.length === 0) throw new Error('At least one compensating mutation is required');
    if (!targetOperationId.trim()) throw new Error('Undo target operationId is required');
    const pendingTarget = this.offlineQueue.getPendingOperation(targetOperationId);
    const target = pendingTarget ?? this.committedOperations.get(targetOperationId);
    if (!target || (pendingTarget && pendingTarget.unitId !== unitId)) {
      throw new Error('UNDO_TARGET_UNRESOLVED: undo requires a queued or acknowledged operation in this workbook');
    }
    return this.enqueueLocalMutations(mutations, unitId, this.createOperationId(), {
      type: 'undo',
      targetOperationId,
      targetBaseRevision: target.baseRevision,
    });
  }

  /** 应用远端已提交 operation — 不进本地撤销栈。 */
  applyRemote(operation: CommittedOperationEnvelope): void {
    this.assertCommittedOperation(operation);
    if (operation.unitId !== this.runtime.workbook.unitId) throw new Error('Remote operation belongs to another workbook');
    if (!Number.isSafeInteger(operation.revision) || operation.revision < 1) throw new Error('Remote operation revision is invalid');
    const pendingLocal = this.offlineQueue.getPendingOperation(operation.operationId);
    if (pendingLocal && !this.committedOperations.has(operation.operationId)) {
      const requested = pendingLocal;
      const committedMutations = operation.mutations.map(({ id, sheetId, params }) => ({ id, sheetId, params }));
      if (operation.origin !== 'client'
        || operation.clientSessionId !== requested.clientSessionId
        || operation.clientSequence !== requested.clientSequence
        || operation.baseRevision !== requested.baseRevision
        || operation.unitId !== requested.unitId
        || JSON.stringify(committedMutations) !== JSON.stringify(requested.mutations)
        || JSON.stringify(operation.intent ?? null) !== JSON.stringify(requested.intent ?? null)) {
        throw new Error('Committed operation does not match the pending local operation');
      }
    }
    const committedFact = this.committedOperations.get(operation.operationId);
    if (committedFact) {
      if (committedFact.baseRevision !== operation.baseRevision || committedFact.revision !== operation.revision) {
        throw new Error('COMMITTED_OPERATION_IDENTITY_MISMATCH: committed revision facts cannot change');
      }
      this.applyValidatedStructuralPatches(operation);
      this.baseRevision = Math.max(this.baseRevision, operation.revision);
      return;
    }
    this.assertNextRevision(operation.revision);
    if (pendingLocal) {
      this.applyValidatedStructuralPatches(operation);
      this.acknowledge(operation.operationId, operation.revision);
      return;
    }
    const incoming = operation.mutations.map((mutation) => committedMutationToClassified(mutation));
    this.assertPendingCanRebase(incoming);
    this.runtime.applyRemoteMutations(operation.mutations.map((mutation) => ({
      id: mutation.id,
      unitId: operation.unitId,
      sheetId: mutation.sheetId,
      params: mutation.params,
      affectedRanges: [...mutation.affectedRanges],
      ...(mutation.structuralImpactRanges?.length ? { structuralImpactRanges: [...mutation.structuralImpactRanges] } : {}),
      ...(mutation.structuralPatch ? {
        structuralFormulaOwnerDeltas: structuredClone(mutation.structuralPatch.formulaOwnerDeltas),
        structuralDefinedNameOwnerDeltas: structuredClone(mutation.structuralPatch.definedNameOwnerDeltas),
        structuralRangeOwnerDeltas: structuredClone(mutation.structuralPatch.rangeOwnerDeltas),
      } : {}),
    })), { operationId: operation.operationId, baseRevision: operation.baseRevision, revision: operation.revision });
    this.committedOperations.set(operation.operationId, { baseRevision: operation.baseRevision, revision: operation.revision });
    for (const classified of incoming) {
      this.committedMutations.push(classified);
      this.remoteMutations.push(classified);
    }
    this.baseRevision = Math.max(this.baseRevision, operation.revision);
    this.rebaseQueuedOperations(this.baseRevision);
  }

  /** Apply server-owned structural facts before making a local operation terminal. */
  applyCommittedStructuralPatches(operation: CommittedOperationEnvelope): void {
    this.assertCommittedOperation(operation);
    if (operation.unitId !== this.runtime.workbook.unitId) throw new Error('Committed operation belongs to another workbook');
    this.applyValidatedStructuralPatches(operation);
  }

  private applyValidatedStructuralPatches(operation: CommittedOperationEnvelope): void {
    this.runtime.applyCommittedStructuralPatches(
      operation.operationId,
      committedMutationInfos(operation),
      operation.revision,
    );
  }

  /** ACK is the only normal path that removes an operation from the queue. */
  acknowledge(operationId: string, revision: number): boolean {
    if (!Number.isSafeInteger(revision) || revision < 1) throw new Error('ACK revision is invalid');
    const committed = this.committedOperations.get(operationId);
    if (committed) {
      if (committed.revision !== revision) throw new Error('OPERATION_ACK_MISMATCH: an acknowledged operation cannot change revision');
      return false;
    }
    if (!this.localClassified.has(operationId) && !this.offlineQueue.hasPendingOperation(operationId)) {
      return false;
    }
    const local = this.localClassified.get(operationId);
    const pending = this.offlineQueue.getPendingOperation(operationId);
    if (!pending) throw new Error('ACK operation lacks its durable base revision');
    if (pending.baseRevision >= revision) throw new Error('ACK revision must follow its durable base revision');
    const nextRevision = Math.max(this.baseRevision, revision);
    const plan = this.planQueuedOperations(nextRevision,
      this.getPendingOperations().filter((operation) => operation.operationId !== operationId),
      { operationId, baseRevision: pending.baseRevision, revision });
    // No history fact or waiter becomes terminal until the single journal image
    // contains both the ACK removal and all unsent dependent undo requests.
    const removed = this.offlineQueue.acknowledgeMany([operationId], plan.operations).length > 0;
    this.runtime.markOperationCommitted(operationId, revision, pending.baseRevision);
    if (local) this.committedMutations.push(...local);
    this.committedOperations.set(operationId, { baseRevision: pending.baseRevision, revision });
    this.localClassified.delete(operationId);
    this.baseRevision = nextRevision;
    this.publishQueueRebase(plan);
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
  rebasePending(
    mutationId: string,
    params: unknown,
    sheetId: string,
    affectedRanges: MutationInfo['affectedRanges'],
    history: readonly ReturnType<typeof classifyMutation>[] = this.committedMutations,
  ) {
    const pending = classifyMutation(mutationId, params, sheetId, [...affectedRanges]);
    return rebaseAgainstHistory(pending, [...history], { sheetOrder: this.currentSheetOrder() });
  }

  private currentSheetOrder(): readonly { readonly id: string; readonly name: string }[] {
    return this.runtime.workbook.sheetOrder.map((id) => ({ id, name: this.runtime.workbook.getSheet(id).name }));
  }

  recordCommittedMutations(mutations: Array<{ id: string; params: unknown; sheetId: string; affectedRanges: MutationInfo['affectedRanges'] }>): void {
    const normalized = mutations.map((mutation) => ({
      ...mutation,
      unitId: this.runtime.workbook.unitId,
    }));
    this.assertMutationContracts(normalized);
    for (const mutation of normalized) {
      this.committedMutations.push(classifyMutation(mutation.id, mutation.params, mutation.sheetId, [...mutation.affectedRanges]));
    }
  }

  /** Load server history before replaying restored offline operations. */
  loadCommittedHistory(operations: readonly CommittedOperationEnvelope[], appliedThroughRevision: number): void {
    if (!Number.isSafeInteger(appliedThroughRevision) || appliedThroughRevision < 0) {
      throw new Error('COLLABORATION_HISTORY_REVISION_INVALID: applied-through revision must be a non-negative safe integer');
    }
    const ordered = [...operations].sort((left, right) => left.revision - right.revision);
    for (const operation of ordered) this.assertCommittedOperation(operation);
    if (ordered.some((operation) => operation.revision > appliedThroughRevision)) {
      throw new Error(`COLLABORATION_HISTORY_AHEAD_OF_MODEL: history exceeds applied revision ${appliedThroughRevision}`);
    }
    const pendingOperationIds = new Set(this.offlineQueue.getPendingOperationIds());
    const newOperations = ordered.filter((operation) => operation.unitId === this.runtime.workbook.unitId
      && !this.committedOperations.has(operation.operationId));
    const acknowledgedPending = newOperations.filter((operation) => pendingOperationIds.has(operation.operationId)).map((operation) => operation.operationId);
    this.offlineQueue.acknowledgeMany(acknowledgedPending);
    for (const operation of newOperations) {
      if (pendingOperationIds.delete(operation.operationId)) {
        this.localClassified.delete(operation.operationId);
      }
      this.committedOperations.set(operation.operationId, { baseRevision: operation.baseRevision, revision: operation.revision });
      for (const mutation of operation.mutations) {
        const classified = committedMutationToClassified(mutation);
        this.committedMutations.push(classified);
        this.remoteMutations.push(classified);
      }
      this.baseRevision = Math.max(this.baseRevision, operation.revision);
    }
    // Hydration already contains these revisions. Replaying their transforms
    // would move pending addresses a second time and create false conflicts.
    this.rebasedRemoteCount = this.remoteMutations.length;
  }

  /** Return pending intent for runtime hydration without exposing queue state. */
  getPendingOperations(): readonly OperationEnvelope[] {
    return this.offlineQueue.getPending().map((entry) => entry.operation);
  }

  /** Explicitly discard the local journal; rejected operations are never
   * removed implicitly by transport failures. */
  clearPending(): void {
    const operationIds = this.offlineQueue.getPendingOperationIds();
    this.offlineQueue.discardMany(operationIds);
    for (const operationId of operationIds) this.localClassified.delete(operationId);
  }

  private async flushOperation(operation: OperationEnvelope): Promise<number> {
    if (!this.send) throw new Error('Collaboration transport unavailable');
    let resolveAck!: (revision: number) => void;
    let rejectAck!: (cause: unknown) => void;
    const ack = new Promise<number>((resolve, reject) => {
      resolveAck = resolve;
      rejectAck = reject;
    });
    void ack.catch(() => undefined);
    this.ackWaiters.set(operation.operationId, { resolve: resolveAck, reject: rejectAck });
    let result: boolean | number;
    try {
      result = await this.send(operation);
    } catch (error) {
      this.ackWaiters.delete(operation.operationId);
      throw error;
    }
    if (typeof result === 'number') {
      this.ackWaiters.delete(operation.operationId);
      this.acknowledge(operation.operationId, result);
      return result;
    }
    if (!result) {
      this.ackWaiters.delete(operation.operationId);
      throw new Error('Collaboration transport unavailable');
    }
    return ack;
  }

  private classifyEnvelope(operation: OperationEnvelope): ReturnType<typeof classifyMutation>[] {
    return operation.mutations.map((mutation) => {
      if (!this.runtime.registry.hasMutation(mutation.id)) throw new Error(`Unknown mutation in pending operation: ${mutation.id}`);
      const metadata = this.runtime.registry.getMutationMetadata(mutation.id);
      if (!metadata?.affectedRanges) throw new Error(`Mutation ${mutation.id} cannot be rebased without affected-range metadata`);
      let affectedRanges: MutationInfo['affectedRanges'];
      try {
        const resolved = metadata.affectedRanges.resolve(mutation.params as never);
        if (!Array.isArray(resolved)) throw new Error('affected-range resolver did not return an array');
        affectedRanges = [...resolved];
      } catch (error) {
        throw new Error(`Cannot classify pending mutation ${mutation.id}: ${error instanceof Error ? error.message : String(error)}`);
      }
      return classifyMutation(mutation.id, mutation.params, mutation.sheetId, affectedRanges);
    });
  }

  private assertMutationContracts(mutations: readonly MutationInfo[]): void {
    for (const mutation of mutations) {
      if (mutation.unitId !== this.runtime.workbook.unitId) throw new Error(`Mutation unit mismatch: ${mutation.id}`);
      if (!this.runtime.registry.hasMutation(mutation.id)) throw new Error(`Unknown mutation: ${mutation.id}`);
      const issues = this.runtime.registry.validateMutationInfo(mutation);
      if (issues.length > 0) throw new Error(`Invalid mutation ${mutation.id}: ${issues.map((entry) => entry.message).join('; ')}`);
    }
  }

  private assertCommittedOperation(operation: CommittedOperationEnvelope): void {
    if (operation.schema !== 'OperationEnvelope') throw new Error('Unsupported committed operation contract');
    if (!operation.operationId.trim() || !operation.unitId.trim() || !operation.actorId.trim()) {
      throw new Error('Committed operation requires operationId, unitId, and actorId');
    }
    if (!Number.isSafeInteger(operation.revision) || operation.revision < 1) throw new Error('Committed operation revision is invalid');
    if (Number.isNaN(Date.parse(operation.createdAt)) || Number.isNaN(Date.parse(operation.committedAt))) {
      throw new Error('Committed operation timestamps are invalid');
    }
    for (const mutation of operation.mutations) {
      if (mutation.structuralPatch === undefined && requiresStructuralPatch(mutation.id)) {
        throw new Error(`Committed mutation ${mutation.id} requires a server-derived StructuralPatch`);
      }
      if (mutation.structuralPatch !== undefined) {
        assertStructuralPatch(mutation.structuralPatch, mutation.id);
        const expectedImpact = structuralPatchImpactRanges(mutation.structuralPatch);
        const actualImpact = mutation.structuralImpactRanges ?? [];
        if (expectedImpact.length !== actualImpact.length || expectedImpact.some((range, index) => !sameRange(range, actualImpact[index]))) {
          throw new Error('Committed structural patch impact ranges are inconsistent');
        }
      } else if ((mutation.structuralImpactRanges?.length ?? 0) > 0) {
        throw new Error('Committed structural impact ranges require a structural patch');
      }
    }
    const request = validateOperationEnvelope({
      schema: operation.schema,
      operationId: operation.operationId,
      unitId: operation.unitId,
      clientSessionId: operation.clientSessionId,
      clientSequence: operation.clientSequence,
      baseRevision: operation.baseRevision,
      mutations: operation.mutations.map(({ id, sheetId, params }) => ({ id, sheetId, params })),
      createdAt: operation.createdAt,
      ...(operation.intent ? { intent: operation.intent } : {}),
    });
    this.assertMutationContracts(operation.mutations.map((mutation) => ({
      id: mutation.id,
      unitId: operation.unitId,
      sheetId: mutation.sheetId,
      params: mutation.params,
      affectedRanges: [...mutation.affectedRanges],
    })));
    if (request.mutations.length !== operation.mutations.length) throw new Error('Committed operation mutation count mismatch');
  }

  private assertPendingCanRebase(committed: readonly ReturnType<typeof classifyMutation>[]): void {
    for (const operationId of this.offlineQueue.getPendingOperationIds()) {
      const operation = this.localClassified.has(operationId)
        ? undefined
        : this.offlineQueue.getPendingOperation(operationId);
      const current = this.localClassified.get(operationId) ?? (operation ? this.classifyEnvelope(operation) : undefined);
      if (!current) throw new Error(`PENDING_CLASSIFICATION_MISSING: ${operationId}`);
      for (const mutation of current) {
        for (const other of committed) for (const left of mutation.affectedRanges) for (const right of other.affectedRanges) {
          if (left.sheetId === right.sheetId && left.startRow <= right.endRow && left.endRow >= right.startRow
            && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn) {
            throw new Error(`COLLABORATION_CONFLICT: ${operationId} 的目标已被其他用户修改，草稿已保留`);
          }
        }
        rebaseAgainstHistory(mutation, [...committed], { sheetOrder: this.currentSheetOrder() });
      }
    }
  }

  private rebaseQueuedOperations(revision: number): void {
    const plan = this.planQueuedOperations(revision, this.getPendingOperations());
    if (plan.operations.length > 0) this.offlineQueue.rewrite(plan.operations);
    this.publishQueueRebase(plan);
  }

  private planQueuedOperations(
    revision: number,
    queued: readonly OperationEnvelope[],
    acknowledged?: CommittedOperationFact & { readonly operationId: string },
  ): QueueRebasePlan {
    const rewritten: OperationEnvelope[] = [];
    const rebasedClassifications = new Map<string, ReturnType<typeof classifyMutation>[]>();
    const targetBases = new Map([...this.committedOperations].map(([operationId, fact]) => [operationId, fact.baseRevision]));
    if (acknowledged) targetBases.set(acknowledged.operationId, acknowledged.baseRevision);
    const newRemoteHistory = this.remoteMutations.slice(this.rebasedRemoteCount);
    for (let index = 0; index < queued.length; index += 1) {
      const queuedOperation = queued[index]!;
      if (!this.offlineQueue.canRewrite(queuedOperation.operationId)) {
        rewritten.push(queuedOperation);
        targetBases.set(queuedOperation.operationId, queuedOperation.baseRevision);
        continue;
      }
      const current = this.localClassified.get(queuedOperation.operationId) ?? this.classifyEnvelope(queuedOperation);
      const rebased = current.map((mutation) => this.rebasePending(
        mutation.mutationId,
        mutation.params,
        mutation.sheetId,
        mutation.affectedRanges,
        newRemoteHistory,
      ).rebased);
      rebasedClassifications.set(queuedOperation.operationId, rebased);
      const operation: OperationEnvelope = {
        ...structuredClone(queuedOperation),
        baseRevision: revision + index,
        mutations: queuedOperation.mutations.map((mutation, mutationIndex) => ({
          ...mutation,
          params: rebased[mutationIndex]!.params,
        })),
      };
      if (operation.intent) {
        const targetBaseRevision = targetBases.get(operation.intent.targetOperationId);
        if (targetBaseRevision === undefined) throw new Error('UNDO_TARGET_UNRESOLVED: queued undo must follow its queued or acknowledged target');
        operation.intent = { ...operation.intent, targetBaseRevision };
      }
      targetBases.set(operation.operationId, operation.baseRevision);
      rewritten.push(operation);
    }
    return { operations: rewritten, classifications: rebasedClassifications, remoteCount: this.remoteMutations.length };
  }

  private publishQueueRebase(plan: QueueRebasePlan): void {
    for (const [operationId, mutations] of plan.classifications) {
      this.localClassified.set(operationId, mutations);
    }
    this.rebasedRemoteCount = plan.remoteCount;
  }
}
