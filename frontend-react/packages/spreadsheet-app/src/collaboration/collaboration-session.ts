import { PresenceStore } from './presence';

/**
 * Ephemeral collaboration state.
 *
 * Workbook mutations and history are owned by the cloud kernel. The socket
 * only announces committed revisions and carries presence, so this session
 * deliberately has no operation queue, OT state, or client undo journal.
 */
export class CollaborationSession {
  readonly presence = new PresenceStore();
  private revision = 0;

  setRevision(revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 0) {
      throw new Error('COLLABORATION_REVISION_INVALID');
    }
    if (revision < this.revision) {
      throw new Error('COLLABORATION_REVISION_REGRESSION');
    }
    this.revision = revision;
  }

  getRevision(): number {
    return this.revision;
  }
}
