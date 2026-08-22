import type { CollaborationChangeSet } from '@react-sheets/protocol';

export type ChangesetStatus = 'idle' | 'pending' | 'acknowledged' | 'rejected' | 'offline';

export interface ChangesetState {
  status: ChangesetStatus;
  operationId?: string;
  revision: number;
  error?: string;
}

export class ChangesetStateMachine {
  private state: ChangesetState = { status: 'idle', revision: 0 };

  get snapshot(): ChangesetState { return { ...this.state }; }

  submit(changeSet: CollaborationChangeSet): ChangesetState {
    if (changeSet.baseRevision !== this.state.revision) throw new Error('Changeset base revision is stale');
    this.state = { status: 'pending', operationId: changeSet.operationId, revision: this.state.revision };
    return this.snapshot;
  }

  acknowledge(operationId: string, revision: number): ChangesetState {
    if (this.state.operationId !== operationId) throw new Error('Unknown changeset operation');
    if (revision <= this.state.revision) throw new Error('Revision must increase');
    this.state = { status: 'acknowledged', operationId, revision };
    return this.snapshot;
  }

  reject(operationId: string, error: string): ChangesetState {
    if (this.state.operationId !== operationId) throw new Error('Unknown changeset operation');
    this.state = { status: 'rejected', operationId, revision: this.state.revision, error };
    return this.snapshot;
  }

  markOffline(): ChangesetState {
    this.state = { ...this.state, status: 'offline' };
    return this.snapshot;
  }
}
