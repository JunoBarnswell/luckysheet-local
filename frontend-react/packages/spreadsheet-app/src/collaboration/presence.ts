import type { RangeRef } from '@react-sheets/core-model';

/** Presence 不进 revision — 仅 UI 展示 */
export interface PresenceUser {
  actorId: string;
  displayName: string;
  color: string;
  avatarUrl?: string;
}

export interface RemoteSelection {
  actorId: string;
  ranges: RangeRef[];
  activeCell?: { row: number; column: number };
}

export interface RemoteEditSession {
  actorId: string;
  sheetId: string;
  row: number;
  column: number;
  draftPreview?: string;
}

export interface PresenceSnapshot {
  users: PresenceUser[];
  selections: RemoteSelection[];
  editSessions: RemoteEditSession[];
  updatedAt: number;
}

export class PresenceStore {
  private users = new Map<string, PresenceUser>();
  private selections = new Map<string, RemoteSelection>();
  private editSessions = new Map<string, RemoteEditSession>();

  upsertUser(user: PresenceUser): void {
    this.users.set(user.actorId, user);
  }

  removeUser(actorId: string): void {
    this.users.delete(actorId);
    this.selections.delete(actorId);
    this.editSessions.delete(actorId);
  }

  updateSelection(selection: RemoteSelection): void {
    this.selections.set(selection.actorId, selection);
  }

  updateEditSession(session: RemoteEditSession): void {
    this.editSessions.set(session.actorId, session);
  }

  clearEditSession(actorId: string): void {
    this.editSessions.delete(actorId);
  }

  snapshot(): PresenceSnapshot {
    return {
      users: [...this.users.values()],
      selections: [...this.selections.values()],
      editSessions: [...this.editSessions.values()],
      updatedAt: Date.now(),
    };
  }
}
