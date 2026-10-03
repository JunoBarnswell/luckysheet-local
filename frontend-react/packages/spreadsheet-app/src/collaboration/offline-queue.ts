import { ApiRequestError, validateOperationEnvelope, type OperationEnvelope } from '@react-sheets/protocol';

export type OfflineQueueState = 'idle' | 'syncing' | 'offline' | 'error';
export type QueuedOperationStatus = 'pending' | 'sent' | 'acked' | 'rejected';

export interface QueuedOperation {
  operation: OperationEnvelope;
  enqueuedAt: number;
  retryCount: number;
  status: QueuedOperationStatus;
  rejection?: Error;
}

export interface OfflineQueueOptions {
  maxRetries?: number;
  flush?: (operation: OperationEnvelope) => Promise<number>;
  now?: () => number;
  /** Durable journal for pending operations. */
  load?: () => readonly OperationEnvelope[];
  persist?: (operations: readonly OperationEnvelope[]) => void;
}

function validateOperation(operation: unknown): OperationEnvelope {
  return validateOperationEnvelope(operation);
}

/**
 * Durable-in-session operation queue.
 *
 * The transport callback resolves only after the server ACK is observed. A
 * send() return value is never treated as commit evidence, so reconnects and
 * dropped ACKs cannot silently remove a local operation. Rejected operations
 * remain visible and block later operations until an explicit discard.
 */
export class OfflineQueue {
  private queue: QueuedOperation[] = [];
  private state: OfflineQueueState = 'idle';
  private readonly maxRetries: number;
  private readonly flush?: (operation: OperationEnvelope) => Promise<number>;
  private readonly now: () => number;
  private readonly persist?: OfflineQueueOptions['persist'];
  private flushPromise: Promise<{ flushed: number; failed: number }> | null = null;
  private readonly terminalStatuses = new Map<string, QueuedOperationStatus>();
  private readonly immutableRequests = new Set<string>();
  private readonly resultLookups = new Set<string>();

  constructor(options: OfflineQueueOptions = {}) {
    this.maxRetries = Math.max(1, options.maxRetries ?? 5);
    this.flush = options.flush;
    this.now = options.now ?? Date.now;
    this.persist = options.persist;
    const restored = options.load?.() ?? [];
    const seen = new Set<string>();
    for (const candidate of restored) {
      const operation = validateOperation(candidate);
      if (seen.has(operation.operationId)) continue;
      seen.add(operation.operationId);
      this.immutableRequests.add(operation.operationId);
      this.resultLookups.add(operation.operationId);
      // A process may have terminated after sending but before receiving ACK.
      // Resubmission of the same operationId is server-side idempotent.
      this.queue.push({
        operation: structuredClone(operation),
        enqueuedAt: this.now(),
        retryCount: 0,
        status: 'pending',
      });
    }
    const sequences = new Map<string, number>();
    for (const { operation } of this.queue) {
      if (operation.clientSequence <= (sequences.get(operation.clientSessionId) ?? 0)) throw new Error('RECOVERY_SEQUENCE_INVALID: 恢复日志会话内顺序无效');
      sequences.set(operation.clientSessionId, operation.clientSequence);
    }
  }

  getState(): OfflineQueueState {
    return this.state;
  }

  canRewrite(operationId: string): boolean {
    return !this.immutableRequests.has(operationId);
  }

  requiresResultLookup(operationId: string): boolean { return this.resultLookups.has(operationId); }

  getPendingCount(): number {
    return this.queue.length;
  }

  hasPendingOperation(operationId: string): boolean {
    return this.queue.some((item) => item.operation.operationId === operationId);
  }

  getPendingOperation(operationId: string): OperationEnvelope | undefined {
    const item = this.queue.find((entry) => entry.operation.operationId === operationId);
    return item ? structuredClone(item.operation) : undefined;
  }

  getPendingOperationIds(): string[] {
    return this.queue.map((item) => item.operation.operationId);
  }

  getPending(): readonly QueuedOperation[] {
    return this.queue.map((item) => ({ ...item, operation: structuredClone(item.operation) }));
  }

  getStatus(operationId: string): QueuedOperationStatus | undefined {
    return this.queue.find((item) => item.operation.operationId === operationId)?.status
      ?? this.terminalStatuses.get(operationId);
  }

  enqueue(operation: OperationEnvelope): void {
    operation = validateOperation(operation);
    if (this.terminalStatuses.has(operation.operationId)) throw new Error(`Operation id was already acknowledged: ${operation.operationId}`);
    if (this.queue.some((item) => item.operation.operationId === operation.operationId)) return;
    this.queue.push({
      operation: structuredClone(operation),
      enqueuedAt: this.now(),
      retryCount: 0,
      status: 'pending',
    });
    this.persistQueue();
  }

  /** Remove only after the matching server ACK has been received. */
  acknowledgeMany(operationIds: readonly string[], remainingOperations?: readonly OperationEnvelope[]): string[] {
    const requested = new Set(operationIds);
    const removed = this.queue.filter((entry) => requested.has(entry.operation.operationId)).map((entry) => entry.operation.operationId);
    if (removed.length === 0) return [];
    const remaining = this.queue.filter((entry) => !requested.has(entry.operation.operationId));
    // Removing the ACK and rebasing its dependants is one durable transition.
    this.commitQueueImage(remainingOperations ?? remaining.map((entry) => entry.operation), remaining);
    for (const operationId of removed) {
      this.terminalStatuses.set(operationId, 'acked');
      this.immutableRequests.delete(operationId);
      this.resultLookups.delete(operationId);
    }
    if (this.queue.length === 0 && this.state === 'syncing') this.state = 'idle';
    return removed;
  }

  /** Keep the rejected operation for audit/retry visibility. */
  reject(operationId: string, cause: unknown): boolean {
    const item = this.queue.find((entry) => entry.operation.operationId === operationId);
    if (!item) return false;
    item.status = 'rejected';
    item.rejection = cause instanceof Error ? cause : new Error(String(cause));
    this.persistQueue();
    this.state = 'error';
    return true;
  }

  /** Explicit user/operator action; never called merely because a send failed. */
  discard(operationId: string): boolean {
    return this.discardMany([operationId]).length > 0;
  }

  discardMany(operationIds: readonly string[]): string[] {
    const requested = new Set(operationIds);
    const removed: string[] = [];
    for (let index = this.queue.length - 1; index >= 0; index -= 1) {
      const operationId = this.queue[index]!.operation.operationId;
      if (!requested.has(operationId)) continue;
      removed.push(operationId);
      this.queue.splice(index, 1);
    }
    if (removed.length === 0) return [];
    if (this.queue.length === 0) this.state = 'idle';
    this.persistQueue();
    return removed.reverse();
  }

  setOnline(online: boolean): void {
    this.state = online ? 'idle' : 'offline';
    if (online) void this.flushAll();
  }

  /** Rewrite all queued operations in sequence and persist one journal image. */
  rewrite(operations: readonly OperationEnvelope[]): void {
    this.commitQueueImage(operations, this.queue);
  }

  private commitQueueImage(operations: readonly OperationEnvelope[], previousEntries: readonly QueuedOperation[]): void {
    operations = operations.map((operation) => validateOperation(operation));
    if (operations.length !== previousEntries.length
      || operations.some((operation, index) => operation.operationId !== previousEntries[index]!.operation.operationId)) {
      throw new Error('OPERATION_QUEUE_IDENTITY_MISMATCH: queue transitions must preserve remaining operation order and identity');
    }
    const rewritten: QueuedOperation[] = [];
    for (let index = 0; index < operations.length; index += 1) {
      const operation = operations[index]!;
      const previous = previousEntries[index]!;
      if (operation.clientSessionId !== previous.operation.clientSessionId
        || operation.clientSequence !== previous.operation.clientSequence
        || operation.unitId !== previous.operation.unitId || operation.createdAt !== previous.operation.createdAt) {
        throw new Error('OPERATION_QUEUE_IDENTITY_MISMATCH: rebasing cannot change an operation identity');
      }
      if (!this.canRewrite(operation.operationId)) {
        if (JSON.stringify(previous.operation) !== JSON.stringify(operation)) throw new Error('OPERATION_REQUEST_IMMUTABLE: 不得改写待确认请求');
        rewritten.push(previous);
        continue;
      }
      rewritten.push({
        operation: structuredClone(operation),
        enqueuedAt: previous.enqueuedAt,
        retryCount: previous.retryCount,
        status: previous.status === 'rejected' ? 'rejected' : 'pending',
        rejection: previous.status === 'rejected' ? previous.rejection : undefined,
      });
    }
    // Commit the durable image before swapping the in-memory queue reference.
    this.persist?.(rewritten.map((item) => structuredClone(item.operation)));
    this.queue = rewritten;
  }

  async flushAll(): Promise<{ flushed: number; failed: number }> {
    if (this.flushPromise) return this.flushPromise;
    const run = this.flushQueue();
    this.flushPromise = run;
    try {
      return await run;
    } finally {
      this.flushPromise = null;
    }
  }

  private async flushQueue(): Promise<{ flushed: number; failed: number }> {
    if (!this.flush || this.queue.length === 0) return { flushed: 0, failed: 0 };
    this.state = 'syncing';
    let flushed = 0;
    let failed = 0;

    while (this.queue.length > 0) {
      const item = this.queue[0]!;
      if (item.status === 'rejected') {
        this.state = 'error';
        break;
      }
      item.status = 'sent';
      this.immutableRequests.add(item.operation.operationId);
      try {
        await this.flush(item.operation);
        // The callback is defined to resolve only after ACK. Guard the
        // operation id so an ACK for another operation cannot dequeue this one.
        if (!this.queue.some((entry) => entry.operation.operationId === item.operation.operationId)) {
          flushed += 1;
          continue;
        }
        if (this.queue[0]?.operation.operationId !== item.operation.operationId) continue;
        this.queue.shift();
        this.persistQueue();
        flushed += 1;
      } catch (cause) {
        if ((item as QueuedOperation).status === 'rejected') {
          failed += 1;
          this.state = 'error';
          break;
        }
        if (cause instanceof ApiRequestError && cause.code === 'UNSUPPORTED_FEATURE') {
          item.status = 'rejected';
          item.rejection = cause;
          failed += 1;
          this.state = 'error';
          this.persistQueue();
          break;
        }
        this.resultLookups.add(item.operation.operationId);
        item.retryCount += 1;
        item.status = 'pending';
        if (item.retryCount >= this.maxRetries) {
          item.status = 'rejected';
          item.rejection = cause instanceof Error ? cause : new Error(String(cause));
          failed += 1;
          this.state = 'error';
        } else {
          this.state = 'error';
        }
        this.persistQueue();
        break;
      }
    }

    if (this.queue.length === 0) this.state = 'idle';
    return { flushed, failed };
  }

  private persistQueue(): void {
    this.persist?.(this.queue.map((item) => structuredClone(item.operation)));
  }
}
