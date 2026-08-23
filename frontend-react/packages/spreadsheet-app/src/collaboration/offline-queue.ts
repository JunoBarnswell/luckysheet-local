import type { CollaborationChangeSet } from '@react-sheets/protocol';

export type OfflineQueueState = 'idle' | 'syncing' | 'offline' | 'error';

export interface QueuedChangeSet {
  changeSet: CollaborationChangeSet;
  enqueuedAt: number;
  retryCount: number;
}

export interface OfflineQueueOptions {
  maxRetries?: number;
  flush?: (changeSet: CollaborationChangeSet) => Promise<number>;
}

/** 离线变更队列 — 连接恢复后按 clientSequence 顺序 flush，不丢弃 pending */
export class OfflineQueue {
  private readonly queue: QueuedChangeSet[] = [];
  private state: OfflineQueueState = 'idle';
  private readonly maxRetries: number;
  private readonly flush?: (changeSet: CollaborationChangeSet) => Promise<number>;

  constructor(options: OfflineQueueOptions = {}) {
    this.maxRetries = options.maxRetries ?? 5;
    this.flush = options.flush;
  }

  getState(): OfflineQueueState {
    return this.state;
  }

  getPendingCount(): number {
    return this.queue.length;
  }

  getPending(): readonly QueuedChangeSet[] {
    return [...this.queue];
  }

  enqueue(changeSet: CollaborationChangeSet): void {
    this.queue.push({ changeSet, enqueuedAt: Date.now(), retryCount: 0 });
    this.queue.sort((a, b) => a.changeSet.clientSequence - b.changeSet.clientSequence);
  }

  dequeueByOperationId(operationId: string): boolean {
    const index = this.queue.findIndex((item) => item.changeSet.operationId === operationId);
    if (index < 0) return false;
    this.queue.splice(index, 1);
    if (this.queue.length === 0 && this.state === 'syncing') this.state = 'idle';
    return true;
  }

  setOnline(online: boolean): void {
    this.state = online ? 'idle' : 'offline';
    if (online) void this.flushAll();
  }

  async flushAll(): Promise<{ flushed: number; failed: number }> {
    if (!this.flush || this.queue.length === 0) return { flushed: 0, failed: 0 };
    this.state = 'syncing';
    let flushed = 0;
    let failed = 0;

    while (this.queue.length > 0) {
      const item = this.queue[0]!;
      try {
        await this.flush(item.changeSet);
        this.queue.shift();
        flushed += 1;
      } catch {
        item.retryCount += 1;
        if (item.retryCount >= this.maxRetries) {
          this.queue.shift();
          failed += 1;
          this.state = 'error';
          break;
        }
        this.state = 'error';
        break;
      }
    }

    if (this.queue.length === 0 && this.state !== 'error') {
      this.state = 'idle';
    }
    return { flushed, failed };
  }
}
