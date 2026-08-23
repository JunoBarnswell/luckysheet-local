import type { PivotResultTree, RangeRef, TableScalar, WorkbookSnapshotV1, WorkbookTableBlock, WorkbookTableModel } from '@react-sheets/core-model';

export type ProtocolErrorCode =
  | 'VALIDATION_ERROR'
  | 'NOT_FOUND'
  | 'CONFLICT'
  | 'UNAUTHENTICATED'
  | 'FORBIDDEN'
  | 'AUTH_CONFIGURATION_ERROR'
  | 'INTERNAL_ERROR';

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
  clientSequence: number;
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

/**
 * The only client-authored operation wire contract.
 *
 * Actor identity and affected ranges intentionally do not exist on this
 * request type.  The server obtains the actor from the verified bearer token
 * and derives ranges from the authoritative workbook/runtime state.  Keeping
 * those values out of the request makes it impossible for a client to turn a
 * claimed actor or range into an authorization decision by accident.
 */
export const OPERATION_ENVELOPE_V2_SCHEMA = 'OperationEnvelopeV2' as const;

export interface OperationMutationV2 {
  id: string;
  sheetId: string;
  params: unknown;
}

export interface OperationEnvelopeV2 {
  schema: typeof OPERATION_ENVELOPE_V2_SCHEMA;
  operationId: string;
  unitId: string;
  clientSequence: number;
  baseRevision: number;
  mutations: OperationMutationV2[];
  createdAt: string;
}

/** Server-authored metadata added after authentication, validation and commit. */
export interface CommittedOperationMutationV2 extends OperationMutationV2 {
  /** Authoritative range calculated by the server; never read from client input. */
  affectedRanges: RangeRef[];
}

export interface CommittedOperationEnvelopeV2 extends Omit<OperationEnvelopeV2, 'mutations'> {
  actorId: string;
  revision: number;
  committedAt: string;
  mutations: CommittedOperationMutationV2[];
}

export type WorkbookAclRole = 'owner' | 'editor' | 'commenter' | 'viewer';

export interface WorkbookAclRecord {
  unitId: string;
  subject: string;
  role: WorkbookAclRole;
  createdAt: string;
  updatedAt: string;
}

export interface OperationCommitResponse {
  operation: CommittedOperationEnvelopeV2;
}

const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0;

/** Strict runtime validation used at the REST/WebSocket trust boundary. */
export function validateOperationEnvelopeV2(value: unknown): OperationEnvelopeV2 {
  if (!value || typeof value !== 'object') throw new Error('OperationEnvelopeV2 must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_V2_SCHEMA) throw new Error('Unsupported operation schema');
  if (!isNonEmptyString(input.operationId) || !isNonEmptyString(input.unitId)) {
    throw new Error('operationId and unitId are required');
  }
  if (!Number.isSafeInteger(input.clientSequence) || Number(input.clientSequence) < 1) {
    throw new Error('clientSequence must be a positive safe integer');
  }
  if (!Number.isSafeInteger(input.baseRevision) || Number(input.baseRevision) < 0) {
    throw new Error('baseRevision must be a non-negative safe integer');
  }
  if (!isNonEmptyString(input.createdAt) || Number.isNaN(Date.parse(input.createdAt))) {
    throw new Error('createdAt must be an ISO timestamp');
  }
  if (!Array.isArray(input.mutations) || input.mutations.length === 0) {
    throw new Error('mutations must contain at least one mutation');
  }

  // Reject fields that used to be client-controlled security inputs instead
  // of silently ignoring them.  This prevents accidental reintroduction of
  // V1 semantics through an untyped JSON caller.
  if ('actorId' in input || 'affectedRanges' in input) {
    throw new Error('actorId and affectedRanges are server-owned fields');
  }

  const mutations = input.mutations.map((raw, index) => {
    if (!raw || typeof raw !== 'object') throw new Error(`mutation[${index}] must be an object`);
    const mutation = raw as Record<string, unknown>;
    if (!isNonEmptyString(mutation.id) || !isNonEmptyString(mutation.sheetId)) {
      throw new Error(`mutation[${index}] requires id and sheetId`);
    }
    if (!('params' in mutation)) throw new Error(`mutation[${index}] requires params`);
    if ('affectedRanges' in mutation || 'actorId' in mutation) {
      throw new Error(`mutation[${index}] contains server-owned fields`);
    }
    return {
      id: mutation.id,
      sheetId: mutation.sheetId,
      params: mutation.params,
    } satisfies OperationMutationV2;
  });

  return {
    schema: OPERATION_ENVELOPE_V2_SCHEMA,
    operationId: input.operationId,
    unitId: input.unitId,
    clientSequence: Number(input.clientSequence),
    baseRevision: Number(input.baseRevision),
    mutations,
    createdAt: input.createdAt,
  };
}

export interface SnapshotResponse {
  snapshot: WorkbookSnapshotV1;
  revision: number;
}

export interface CompatibilityReportPayload {
  schema: 'CompatibilityReportV1';
  fileName: string;
  importLevel: 'A' | 'B' | 'C';
  exportLevel: 'A' | 'B' | 'C';
  dateSystem: '1900' | '1904';
  issues: Array<{
    level: 'A' | 'B' | 'C';
    severity: 'error' | 'warning' | 'info';
    feature: string;
    location?: string;
    message: string;
    preserved: boolean;
  }>;
  summary: {
    editableFeatures: number;
    preservedOnly: number;
    unsupported: number;
  };
}

export interface XlsxImportResponse extends SnapshotResponse {
  report: CompatibilityReportPayload;
}

export interface XlsxExportResponse {
  unitId: string;
  base64: string;
  fileName: string;
  report: CompatibilityReportPayload;
}

export interface PivotCalculationResponse {
  unitId: string;
  pivotId: string;
  revision: number;
  result: PivotResultTree;
}

export interface TableRowsResponse {
  table: WorkbookTableModel;
  rows: TableScalar[][];
  nextOffset?: number;
}

export interface WorkbookSummary {
  unitId: string;
  name: string;
  revision: number;
  updatedAt: string;
}

export interface RevisionRecord {
  operationId: string;
  revision: number;
  createdAt: string;
  payload: CollaborationChangeSet;
}

export class WorkbookApiClient {
  constructor(private readonly baseUrl = '') {}

  async getSnapshot(unitId: string): Promise<SnapshotResponse> {
    const response = await fetch(
      `${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/snapshot`,
    );
    if (!response.ok) throw new Error(`Workbook snapshot fetch failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async createWorkbook(snapshot: WorkbookSnapshotV1): Promise<SnapshotResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ snapshot }),
    });
    if (!response.ok) throw new Error(`Workbook creation failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async createEmptyWorkbook(name = 'Untitled workbook'): Promise<SnapshotResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name }),
    });
    if (!response.ok) throw new Error(`Workbook creation failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async listWorkbooks(): Promise<WorkbookSummary[]> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks`);
    if (!response.ok) throw new Error(`Workbook list fetch failed: ${response.status}`);
    return response.json() as Promise<WorkbookSummary[]>;
  }

  async deleteWorkbook(unitId: string): Promise<void> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}`, { method: 'DELETE' });
    if (!response.ok) throw new Error(`Workbook deletion failed: ${response.status}`);
  }

  async listRevisions(unitId: string): Promise<RevisionRecord[]> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/revisions`);
    if (!response.ok) throw new Error(`Revision history fetch failed: ${response.status}`);
    const body = await response.json() as { revisions: RevisionRecord[] };
    return body.revisions;
  }

  async getRevisionSnapshot(unitId: string, revision: number): Promise<SnapshotResponse> {
    const response = await fetch(
      `${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/revisions/${revision}/snapshot`,
    );
    if (!response.ok) throw new Error(`Revision snapshot fetch failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async saveSnapshot(unitId: string, snapshot: WorkbookSnapshotV1, baseRevision: number): Promise<SnapshotResponse> {
    const response = await fetch(
      `${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/snapshot`,
      {
        method: 'PUT',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ snapshot, baseRevision }),
      },
    );
    if (response.status === 409) throw new Error('Revision conflict');
    if (!response.ok) throw new Error(`Workbook snapshot save failed: ${response.status}`);
    return response.json() as Promise<SnapshotResponse>;
  }

  async importXlsxBase64(base64: string, fileName = 'import.xlsx'): Promise<XlsxImportResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/files/import-xlsx`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ base64, fileName }),
    });
    if (!response.ok) throw new Error(`XLSX import failed: ${response.status}`);
    return response.json() as Promise<XlsxImportResponse>;
  }

  async exportXlsx(unitId: string, fileName?: string): Promise<XlsxExportResponse> {
    const query = fileName ? `?fileName=${encodeURIComponent(fileName)}` : '';
    const response = await fetch(
      `${this.baseUrl}/api/v1/files/${encodeURIComponent(unitId)}/export${query}`,
    );
    if (!response.ok) throw new Error(`XLSX export failed: ${response.status}`);
    return response.json() as Promise<XlsxExportResponse>;
  }

  async calculatePivot(unitId: string, pivotId: string): Promise<PivotCalculationResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/calculations/pivot`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ pivotId }),
    });
    if (!response.ok) throw new Error(`Pivot calculation failed: ${response.status}`);
    return response.json() as Promise<PivotCalculationResponse>;
  }

  async createDataTable(unitId: string, table: WorkbookTableModel): Promise<WorkbookTableModel> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/tables`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(table),
    });
    if (!response.ok) throw new Error(`Data table creation failed: ${response.status}`);
    return response.json() as Promise<WorkbookTableModel>;
  }

  async appendDataBlock(unitId: string, tableId: string, startRow: number, rows: TableScalar[][]): Promise<WorkbookTableBlock> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}/blocks`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ startRow, rows }),
    });
    if (!response.ok) throw new Error(`Data block upload failed: ${response.status}`);
    return response.json() as Promise<WorkbookTableBlock>;
  }

  async deleteDataTable(unitId: string, tableId: string): Promise<void> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}`, { method: 'DELETE' });
    if (!response.ok) throw new Error(`Data table deletion failed: ${response.status}`);
  }

  async readDataRows(unitId: string, tableId: string, offset = 0, limit = 500): Promise<TableRowsResponse> {
    const response = await fetch(`${this.baseUrl}/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}/rows?offset=${offset}&limit=${limit}`);
    if (!response.ok) throw new Error(`Data table query failed: ${response.status}`);
    return response.json() as Promise<TableRowsResponse>;
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

/**
 * V2 collaboration messages.  Client-originated presence/cursor messages do
 * not carry actorId; the server adds it to broadcast messages after token
 * verification.  The V1 message union above remains only for source-level
 * migration and is not accepted by the V2 server endpoint.
 */
export type OperationMessageV2 =
  | { type: 'snapshot.request'; unitId: string }
  | { type: 'snapshot.response'; payload: SnapshotResponse }
  | { type: 'changeset.submit'; payload: OperationEnvelopeV2 }
  | { type: 'changeset.ack'; operationId: string; revision: number }
  | { type: 'changeset.reject'; operationId: string; error: ApiError }
  | { type: 'revision.created'; payload: CommittedOperationEnvelopeV2; revision: number }
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown }
  | { type: 'presence.broadcast'; unitId: string; actorId: string; state: unknown }
  | { type: 'cursor.broadcast'; unitId: string; actorId: string; state: unknown };

export type ClientOperationMessageV2 =
  | { type: 'snapshot.request'; unitId: string }
  | { type: 'changeset.submit'; payload: OperationEnvelopeV2 }
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown };

export function encodeOperationMessageV2(message: OperationMessageV2): string {
  return JSON.stringify(message);
}

function validateCommittedOperationEnvelopeV2(value: unknown): CommittedOperationEnvelopeV2 {
  if (!value || typeof value !== 'object') throw new Error('Committed operation must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_V2_SCHEMA || !isNonEmptyString(input.operationId) || !isNonEmptyString(input.unitId)) {
    throw new Error('Invalid committed operation schema');
  }
  if (!isNonEmptyString(input.actorId) || !Number.isSafeInteger(input.revision) || Number(input.revision) < 1) {
    throw new Error('Invalid committed operation metadata');
  }
  if (!isNonEmptyString(input.committedAt) || Number.isNaN(Date.parse(input.committedAt))) {
    throw new Error('Invalid committed operation timestamp');
  }
  const operation = validateOperationEnvelopeV2({
    schema: input.schema,
    operationId: input.operationId,
    unitId: input.unitId,
    clientSequence: input.clientSequence,
    baseRevision: input.baseRevision,
    createdAt: input.createdAt,
    mutations: Array.isArray(input.mutations)
      ? input.mutations.map((mutation) => {
        if (!mutation || typeof mutation !== 'object') return mutation;
        const candidate = mutation as Record<string, unknown>;
        return {
          id: candidate.id,
          sheetId: candidate.sheetId,
          params: candidate.params,
        };
      })
      : input.mutations,
  });
  const mutations = (input.mutations as unknown[]).map((raw, index) => {
    if (!raw || typeof raw !== 'object') throw new Error(`committed mutation[${index}] must be an object`);
    const mutation = raw as Record<string, unknown>;
    if (!Array.isArray(mutation.affectedRanges)) throw new Error(`committed mutation[${index}] requires affectedRanges`);
    return { ...operation.mutations[index]!, affectedRanges: mutation.affectedRanges as RangeRef[] };
  });
  return {
    ...operation,
    actorId: input.actorId,
    revision: Number(input.revision),
    committedAt: input.committedAt,
    mutations,
  };
}

export function decodeOperationMessageV2(input: string): OperationMessageV2 {
  let raw: unknown;
  try {
    raw = JSON.parse(input);
  } catch {
    throw new Error('Invalid collaboration JSON');
  }
  if (!raw || typeof raw !== 'object' || typeof (raw as Record<string, unknown>).type !== 'string') {
    throw new Error('Invalid collaboration message');
  }
  const message = raw as Record<string, unknown>;
  switch (message.type) {
    case 'changeset.submit':
      return { type: 'changeset.submit', payload: validateOperationEnvelopeV2(message.payload) };
    case 'snapshot.request':
      if (!isNonEmptyString(message.unitId)) throw new Error('snapshot.request requires unitId');
      return { type: 'snapshot.request', unitId: message.unitId };
    case 'snapshot.response':
      if (!message.payload || typeof message.payload !== 'object') throw new Error('snapshot.response requires payload');
      return { type: 'snapshot.response', payload: message.payload as SnapshotResponse };
    case 'changeset.ack':
      if (!isNonEmptyString(message.operationId) || !Number.isSafeInteger(message.revision) || Number(message.revision) < 1) {
        throw new Error('changeset.ack requires operationId and revision');
      }
      return { type: 'changeset.ack', operationId: message.operationId, revision: Number(message.revision) };
    case 'changeset.reject':
      if (!isNonEmptyString(message.operationId) || !message.error || typeof message.error !== 'object') {
        throw new Error('changeset.reject requires operationId and error');
      }
      return { type: 'changeset.reject', operationId: message.operationId, error: message.error as ApiError };
    case 'revision.created':
      if (!Number.isSafeInteger(message.revision) || Number(message.revision) < 1) throw new Error('revision.created requires revision');
      return {
        type: 'revision.created',
        payload: validateCommittedOperationEnvelopeV2(message.payload),
        revision: Number(message.revision),
      };
    case 'presence.updated':
    case 'cursor.updated':
      if (!isNonEmptyString(message.unitId)) throw new Error(`${message.type} requires unitId`);
      if ('actorId' in message) throw new Error('actorId is server-owned');
      return { type: message.type, unitId: message.unitId, state: message.state };
    case 'presence.broadcast':
    case 'cursor.broadcast':
      if (!isNonEmptyString(message.unitId) || !isNonEmptyString(message.actorId)) {
        throw new Error(`${message.type} requires unitId and actorId`);
      }
      return { type: message.type, unitId: message.unitId, actorId: message.actorId, state: message.state };
    default:
      throw new Error(`Unsupported collaboration message: ${String(message.type)}`);
  }
}

export function decodeClientOperationMessageV2(input: string): ClientOperationMessageV2 {
  const message = decodeOperationMessageV2(input);
  if (message.type === 'changeset.submit' || message.type === 'snapshot.request'
    || message.type === 'presence.updated' || message.type === 'cursor.updated') return message;
  throw new Error(`Server-only collaboration message: ${message.type}`);
}

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

/** 连接状态 */
export type CollabSocketStatus = 'connecting' | 'open' | 'closed';

export interface CollabSocketOptions {
  /** 断线重连基础延迟(毫秒),指数退避 */
  reconnectBaseDelayMs?: number;
  /** 重连延迟上限(毫秒) */
  reconnectMaxDelayMs?: number;
}

type CollabSocketListener = (message: CollaborationMessage) => void;
type StatusListener = (status: CollabSocketStatus) => void;

/**
 * 浏览器端协同 WebSocket 客户端。
 * - 断线自动指数退避重连;
 * - 断线期间 send() 的消息进入待发队列,连接恢复后按序冲刷;
 * - 消息格式复用 encode/decodeMessage 协议。
 */
export class CollabSocketClient {
  private socket: WebSocket | null = null;
  private readonly messageListeners = new Set<CollabSocketListener>();
  private readonly statusListeners = new Set<StatusListener>();
  private readonly pendingOutbound: string[] = [];
  private reconnectAttempt = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private closedByUser = false;

  constructor(
    private readonly url: string,
    private readonly options: CollabSocketOptions = {},
  ) {}

  get status(): CollabSocketStatus {
    if (this.socket?.readyState === WebSocket.OPEN) return 'open';
    if (this.closedByUser) return 'closed';
    return this.socket || this.reconnectTimer ? 'connecting' : 'closed';
  }

  open(): void {
    this.closedByUser = false;
    if (this.socket || this.reconnectTimer) return;
    this.connect();
  }

  close(): void {
    this.closedByUser = true;
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    this.socket?.close();
    this.socket = null;
    this.emitStatus('closed');
  }

  /** 发送消息;未连接时进入队列,连接恢复后冲刷。返回是否即时发出。 */
  send(message: CollaborationMessage): boolean {
    const encoded = encodeMessage(message);
    if (this.socket?.readyState === WebSocket.OPEN) {
      this.socket.send(encoded);
      return true;
    }
    if (message.type !== 'changeset.submit') return false;
    this.pendingOutbound.push(encoded);
    return false;
  }

  onMessage(listener: CollabSocketListener): () => void {
    this.messageListeners.add(listener);
    return () => this.messageListeners.delete(listener);
  }

  onStatus(listener: StatusListener): () => void {
    this.statusListeners.add(listener);
    listener(this.status);
    return () => this.statusListeners.delete(listener);
  }

  /** 仅用于测试:取出当前待发队列长度 */
  get queuedCount(): number {
    return this.pendingOutbound.length;
  }

  private connect(): void {
    this.emitStatus('connecting');
    const socket = new WebSocket(this.url);
    this.socket = socket;

    socket.onopen = () => {
      this.reconnectAttempt = 0;
      this.emitStatus('open');
      while (this.pendingOutbound.length > 0 && socket.readyState === WebSocket.OPEN) {
        const encoded = this.pendingOutbound.shift();
        if (encoded !== undefined) socket.send(encoded);
      }
    };

    socket.onmessage = (event) => {
      if (typeof event.data !== 'string') return;
      try {
        const message = decodeMessage(event.data);
        for (const listener of this.messageListeners) listener(message);
      } catch {
        // 非法消息直接丢弃,保持连接可用
      }
    };

    socket.onclose = () => {
      this.socket = null;
      if (this.closedByUser) {
        this.emitStatus('closed');
        return;
      }
      this.scheduleReconnect();
    };

    socket.onerror = () => {
      socket.close();
    };
  }

  private scheduleReconnect(): void {
    if (this.reconnectTimer || this.closedByUser) return;
    const base = this.options.reconnectBaseDelayMs ?? 800;
    const max = this.options.reconnectMaxDelayMs ?? 15000;
    const delay = Math.min(max, base * Math.pow(2, this.reconnectAttempt));
    this.reconnectAttempt += 1;
    this.emitStatus('connecting');
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.connect();
    }, delay);
  }

  private emitStatus(status: CollabSocketStatus): void {
    for (const listener of this.statusListeners) listener(status);
  }
}
