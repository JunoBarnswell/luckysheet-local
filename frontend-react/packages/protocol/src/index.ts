import type { RangeRef, WorkbookSnapshotV1 } from '@react-sheets/core-model';

export type ProtocolErrorCode = 'VALIDATION_ERROR' | 'NOT_FOUND' | 'CONFLICT' | 'INTERNAL_ERROR';

export interface ApiError {
  code: ProtocolErrorCode;
  message: string;
  details?: Record<string, unknown>;
}

export interface CollaborationChangeSet {
  schema: 'CollaborationChangeSetV1';
  operationId: string;
  unitId: string;
  actorId: string;
  baseRevision: number;
  mutations: CollaborationMutation[];
  createdAt: string;
}

export interface CollaborationMutation {
  id: string;
  sheetId: string;
  params: unknown;
  affectedRanges: RangeRef[];
}

export interface SnapshotResponse {
  snapshot: WorkbookSnapshotV1;
  revision: number;
}

export class WorkbookApiClient {
  constructor(private readonly baseUrl = '') {}

  async createWorkbook(snapshot: WorkbookSnapshotV1): Promise<SnapshotResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ snapshot }),
    });
    if (!response.ok) throw new Error(`Workbook creation failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async submitChangeSet(changeSet: CollaborationChangeSet): Promise<{ operationId: string; revision: number }> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(changeSet.unitId)}/changesets`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(changeSet),
    });
    if (!response.ok) throw new Error(`Changeset rejected: ${response.status}`);
    return response.json() as Promise<{ operationId: string; revision: number }>;
  }
}

export type CollaborationMessage =
  | { type: 'snapshot.request'; unitId: string }
  | { type: 'snapshot.response'; unitId?: string; snapshot?: WorkbookSnapshotV1; revision?: number; payload?: SnapshotResponse }
  | { type: 'changeset.submit'; payload: CollaborationChangeSet }
  | { type: 'changeset.ack'; operationId: string; revision: number }
  | { type: 'changeset.reject'; operationId: string; error: ApiError }
  | { type: 'revision.created'; payload: CollaborationChangeSet; revision: number }
  | { type: 'presence.updated'; unitId: string; actorId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; actorId: string; state: unknown };

export function encodeMessage(message: CollaborationMessage): string {
  return JSON.stringify(message);
}

export function decodeMessage(input: string): CollaborationMessage {
  const message = JSON.parse(input) as CollaborationMessage;
  if (!message || typeof message !== 'object' || typeof message.type !== 'string') {
    throw new Error('Invalid collaboration message');
  }
  return message;
}
