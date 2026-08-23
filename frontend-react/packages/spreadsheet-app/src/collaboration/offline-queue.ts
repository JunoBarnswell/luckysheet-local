import type { OperationEnvelopeV2 } from '@react-sheets/protocol';

export type OfflineQueueState = 'idle' | 'syncing' | 'offline' | 'error';
export type QueuedOperationStatus = 'pending' | 'sent' | 'rejected';

export interface QueuedOperation {
  operation: OperationEnvelopeV2;
  enqueuedAt: number;
  retryCount: number;
  status: QueuedOperationStatus;
  rejection?: Error;
}

export interface OfflineQueueOptions {
  maxRetries?: number;
  flush?: (operation: OperationEnvelopeV2) => Promise<number>;
  now?: () => number;
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
  private readonly queue: QueuedOperation[] = [];
  private state: OfflineQueueState = 'idle';
  private readonly maxRetries: number;
  private readonly flush?: (operation: OperationEnvelopeV2) => Promise<number>;
  private readonly now: () => number;
  private flushPromise: Promise<{ flushed: number; failed: number }> | null = null;

  constructor(options: OfflineQueueOptions = {}) {
    this.maxRetries = Math.max(1, options.maxRetries ?? 5);
    this.flush = options.flush;
    this.now = options.now ?? Date.now;
  }

  getState(): OfflineQueueState {
    return this.state;
  }

  getPendingCount(): number {
    return this.queue.length;
  }

  getPending(): readonly QueuedOperation[] {
    return this.queue.map((item) => ({ ...item, operation: structuredClone(item.operation) }));
  }

  enqueue(operation: OperationEnvelopeV2): void {
    if (this.queue.some((item) => item.operation.operationId === operation.operationId)) return;
    this.queue.push({
      operation: structuredClone(operation),
      enqueuedAt: this.now(),
      retryCount: 0,
      status: 'pending',
    });
    this.queue.sort((a, b) => a.operation.clientSequence - b.operation.clientSequence);
  }

  /** Remove only after the matching server ACK has been received. */
  acknowledge(operationId: string): boolean {
    const index = this.queue.findIndex((item) => item.operation.operationId === operationId);
    if (index < 0) return false;
    this.queue.splice(index, 1);
    if (this.queue.length === 0 && this.state === 'syncing') this.state = 'idle';
    return true;
  }

  /** Keep the rejected operation for audit/retry visibility. */
  reject(operationId: string, cause: unknown): boolean {
    const item = this.queue.find((entry) => entry.operation.operationId === operationId);
    if (!item) return false;
    item.status = 'rejected';
    item.rejection = cause instanceof Error ? cause : new Error(String(cause));
    this.state = 'error';
    return true;
  }

  /** Explicit user/operator action; never called merely because a send failed. */
  discard(operationId: string): boolean {
    const index = this.queue.findIndex((item) => item.operation.operationId === operationId);
    if (index < 0) return false;
    this.queue.splice(index, 1);
    if (this.queue.length === 0) this.state = 'idle';
    return true;
  }

  setOnline(online: boolean): void {
    this.state = online ? 'idle' : 'offline';
    if (online) void this.flushAll();
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
        flushed += 1;
      } catch (cause) {
        if ((item as QueuedOperation).status === 'rejected') {
          failed += 1;
          this.state = 'error';
          break;
        }
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
        break;
      }
    }

    if (this.queue.length === 0) this.state = 'idle';
    return { flushed, failed };
  }
}
