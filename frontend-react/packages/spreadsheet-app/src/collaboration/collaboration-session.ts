import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import type { CollaborationChangeSet } from '@react-sheets/protocol';
import { classifyMutation } from './operation-types';
import { rebaseAgainstHistory } from './ot-rebase';
import { OfflineQueue } from './offline-queue';
import { CollaborativeUndoStack } from './collaborative-undo';
import { PresenceStore } from './presence';

export interface CollaborationSessionOptions {
  actorId: string;
  flush?: (changeSet: CollaborationChangeSet) => Promise<number>;
}

/** 协同会话 — OT rebase + offline queue + 协同撤销 + presence */
export class CollaborationSession {
  readonly presence = new PresenceStore();
  readonly offlineQueue: OfflineQueue;
  readonly collaborativeUndo = new CollaborativeUndoStack();

  private readonly actorId: string;
  private clientSequence = 0;
  private baseRevision = 0;
  private committedMutations: ReturnType<typeof classifyMutation>[] = [];

  constructor(
    private runtime: CommandRuntime,
    options: CollaborationSessionOptions,
  ) {
    this.actorId = options.actorId;
    this.offlineQueue = new OfflineQueue({ flush: options.flush });
  }

  rebindCommands(runtime: CommandRuntime): void {
    this.runtime = runtime;
  }

  setRevision(revision: number): void {
    this.baseRevision = revision;
  }

  getRevision(): number {
    return this.baseRevision;
  }

  /** 本地命令执行后 enqueue 协同变更 */
  enqueueLocalMutations(mutations: MutationInfo[], unitId: string, operationId?: string): CollaborationChangeSet {
    this.clientSequence += 1;
    const changeSet: CollaborationChangeSet = {
      schema: 'CollaborationChangeSetV1',
      operationId: operationId ?? (typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() : `op-${Date.now()}`),
      unitId,
      actorId: this.actorId,
      clientSequence: this.clientSequence,
      baseRevision: this.baseRevision,
      mutations: mutations.map((m) => ({
        id: m.id,
        sheetId: m.sheetId,
        params: m.params,
        affectedRanges: m.affectedRanges,
      })),
      createdAt: new Date().toISOString(),
    };
    this.offlineQueue.enqueue(changeSet);
    return changeSet;
  }

  /** 应用远端变更 — 不进本地 undo 栈 */
  applyRemote(changeSet: CollaborationChangeSet): void {
    this.runtime.applyRemoteMutations(changeSet.mutations.map((m) => ({
      id: m.id,
      unitId: changeSet.unitId,
      sheetId: m.sheetId,
      params: m.params,
      affectedRanges: m.affectedRanges,
    })));
    for (const m of changeSet.mutations) {
      this.committedMutations.push(classifyMutation(m.id, m.params, m.sheetId, m.affectedRanges));
    }
    this.baseRevision = Math.max(this.baseRevision, changeSet.baseRevision + 1);
  }

  /** rebase 离线 pending 操作 */
  rebasePending(mutationId: string, params: unknown, sheetId: string, affectedRanges: MutationInfo['affectedRanges']) {
    const pending = classifyMutation(mutationId, params, sheetId, affectedRanges);
    return rebaseAgainstHistory(pending, this.committedMutations);
  }

  recordCommittedMutations(mutations: Array<{ id: string; params: unknown; sheetId: string; affectedRanges: MutationInfo['affectedRanges'] }>): void {
    for (const mutation of mutations) {
      this.committedMutations.push(classifyMutation(mutation.id, mutation.params, mutation.sheetId, mutation.affectedRanges));
    }
  }

  recordLocalUndo(entry: { operationId: string; undoMutations: MutationInfo[] }): void {
    this.collaborativeUndo.push(this.actorId, {
      operationId: entry.operationId,
      actorId: this.actorId,
      undoMutations: entry.undoMutations,
      timestamp: Date.now(),
    });
  }

  undoOwnLast(): MutationInfo[] | undefined {
    const entry = this.collaborativeUndo.pop(this.actorId);
    if (!entry) return undefined;
    return this.collaborativeUndo.createCompensatingCommand(entry);
  }
}
