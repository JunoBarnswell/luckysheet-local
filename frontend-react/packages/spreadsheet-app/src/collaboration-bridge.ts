import type { MutationInfo } from '@react-sheets/command-runtime';
import type { CollaborationChangeSet, CollaborationMutation } from '@react-sheets/protocol';
import type { CollaborationSession } from './collaboration/collaboration-session';
import type { PresenceSnapshot } from './collaboration/presence';
import type { PeerCursor } from './types';
import { hashCode, PEER_COLORS } from './runtime';

export interface CollaborationSnapshot {
  revision: number;
  pendingCount: number;
  offlineQueueState: string;
  presence: PresenceSnapshot;
  peerCount: number;
}

export function buildChangeSet(
  operationId: string,
  unitId: string,
  actorId: string,
  clientSequence: number,
  baseRevision: number,
  mutations: CollaborationMutation[],
): CollaborationChangeSet {
  return {
    schema: 'CollaborationChangeSetV1',
    operationId,
    unitId,
    actorId,
    clientSequence,
    baseRevision,
    mutations,
    createdAt: new Date().toISOString(),
  };
}

export function mutationsFromBatch(mutations: MutationInfo[], unitId: string): CollaborationMutation[] {
  return mutations.map((mutation) => ({
    id: mutation.id,
    sheetId: mutation.sheetId,
    params: mutation.params,
    affectedRanges: mutation.affectedRanges,
    unitId,
  }));
}

export function mapPeerCursor(
  actorId: string,
  state: { row?: number; column?: number; name?: string; sheetId?: string } | null | undefined,
  fallbackSheetId: string,
): PeerCursor {
  const color = PEER_COLORS[Math.abs(hashCode(actorId)) % PEER_COLORS.length]!;
  return {
    actorId,
    name: state?.name ?? actorId.slice(0, 6),
    color,
    sheetId: state?.sheetId ?? fallbackSheetId,
    row: state?.row ?? 0,
    column: state?.column ?? 0,
  };
}

export function buildCollaborationSnapshot(
  session: CollaborationSession,
  peers: readonly PeerCursor[],
): CollaborationSnapshot {
  return {
    revision: session.getRevision(),
    pendingCount: session.offlineQueue.getPendingCount(),
    offlineQueueState: session.offlineQueue.getState(),
    presence: session.presence.snapshot(),
    peerCount: peers.length,
  };
}

export function updatePresenceFromPeer(
  session: CollaborationSession,
  peer: PeerCursor,
): void {
  session.presence.upsertUser({
    actorId: peer.actorId,
    displayName: peer.name,
    color: peer.color,
  });
  session.presence.updateSelection({
    actorId: peer.actorId,
    ranges: [{
      sheetId: peer.sheetId,
      startRow: peer.row,
      endRow: peer.row,
      startColumn: peer.column,
      endColumn: peer.column,
    }],
    activeCell: { row: peer.row, column: peer.column },
  });
}

export function acknowledgeChangeSet(session: CollaborationSession, revision: number, operationId?: string): void {
  session.setRevision(revision);
  if (operationId) session.acknowledge(operationId, revision);
}
