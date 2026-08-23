import type { MutationInfo } from '@react-sheets/command-runtime';
import type { OperationEnvelope, OperationMutation } from '@react-sheets/protocol';
import type { CollaborationSession } from './collaboration-session';
import type { PresenceSnapshot } from './presence';
import type { PeerCursor } from '../types';

export const PEER_COLORS = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];

export function hashCode(input: string): number {
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
  }
  return hash;
}

export interface CollaborationSnapshot {
  revision: number;
  pendingCount: number;
  offlineQueueState: string;
  presence: PresenceSnapshot;
  peerCount: number;
}

export function buildOperation(
  operationId: string,
  unitId: string,
  clientSequence: number,
  baseRevision: number,
  mutations: OperationMutation[],
  createdAt = new Date().toISOString(),
): OperationEnvelope {
  return {
    schema: 'OperationEnvelope',
    operationId,
    unitId,
    clientSequence,
    baseRevision,
    mutations,
    createdAt,
  };
}

export function mutationsFromBatch(mutations: MutationInfo[]): OperationMutation[] {
  return mutations.map((mutation) => ({
    id: mutation.id,
    sheetId: mutation.sheetId,
    params: mutation.params,
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
