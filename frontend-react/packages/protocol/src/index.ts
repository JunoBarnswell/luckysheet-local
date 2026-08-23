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

/**
 * Authentication is deliberately supplied by the host application.  The
 * protocol package never reads cookies, localStorage or a client-provided
 * actor id.  A missing/empty token is an authentication failure, not an
 * unauthenticated fallback.
 */
export type AuthTokenProvider = () => string | null | Promise<string | null>;

export interface WorkbookApiClientOptions {
  baseUrl?: string;
  authTokenProvider?: AuthTokenProvider;
  fetchImpl?: typeof fetch;
}

export class ApiRequestError extends Error {
  readonly status: number;
  readonly code: ProtocolErrorCode | 'REQUEST_FAILED';
  readonly details?: Record<string, unknown>;

  constructor(
    message: string,
    status: number,
    code: ProtocolErrorCode | 'REQUEST_FAILED' = 'REQUEST_FAILED',
    details?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'ApiRequestError';
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

export class AuthenticationRequiredError extends Error {
  readonly status = 401;
  readonly code = 'UNAUTHENTICATED' as const;

  constructor(message = 'Bearer authentication is required') {
    super(message);
    this.name = 'AuthenticationRequiredError';
  }
}

const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0;

function isRangeRef(value: unknown): value is RangeRef {
  if (!value || typeof value !== 'object') return false;
  const range = value as Record<string, unknown>;
  return isNonEmptyString(range.sheetId)
    && Number.isSafeInteger(range.startRow) && Number(range.startRow) >= 0
    && Number.isSafeInteger(range.endRow) && Number(range.endRow) >= Number(range.startRow)
    && Number.isSafeInteger(range.startColumn) && Number(range.startColumn) >= 0
    && Number.isSafeInteger(range.endColumn) && Number(range.endColumn) >= Number(range.startColumn);
}

function validateSnapshotResponse(value: unknown): SnapshotResponse {
  if (!value || typeof value !== 'object') throw new Error('snapshot.response payload must be an object');
  const input = value as Record<string, unknown>;
  if (!input.snapshot || typeof input.snapshot !== 'object') throw new Error('snapshot.response requires snapshot');
  const snapshot = input.snapshot as Record<string, unknown>;
  if (!isNonEmptyString(snapshot.unitId)) throw new Error('snapshot.response requires snapshot.unitId');
  if (!Number.isSafeInteger(input.revision) || Number(input.revision) < 0) {
    throw new Error('snapshot.response requires a valid revision');
  }
  return value as SnapshotResponse;
}

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
  payload: CommittedOperationEnvelopeV2;
}

export class WorkbookApiClient {
  private readonly baseUrl: string;
  private readonly authTokenProvider?: AuthTokenProvider;
  private readonly fetchImpl: typeof fetch;

  constructor(options: WorkbookApiClientOptions = {}) {
    this.baseUrl = options.baseUrl ?? '';
    this.authTokenProvider = options.authTokenProvider;
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  private async request(path: string, init: RequestInit = {}): Promise<Response> {
    if (!this.authTokenProvider) throw new AuthenticationRequiredError('An AuthTokenProvider must be injected');
    const rawToken = await this.authTokenProvider();
    const token = rawToken?.trim();
    if (!token) throw new AuthenticationRequiredError();
    const headers = new Headers(init.headers);
    headers.set('authorization', `Bearer ${token}`);
    const response = await this.fetchImpl(`${this.baseUrl}${path}`, { ...init, headers });
    if (response.ok) return response;

    let payload: Partial<ApiError> | undefined;
    try {
      const candidate = await response.clone().json() as unknown;
      if (candidate && typeof candidate === 'object') payload = candidate as Partial<ApiError>;
    } catch {
      // The status remains authoritative when a server returns a non-JSON
      // failure body.  Do not swallow the request failure itself.
    }
    throw new ApiRequestError(
      typeof payload?.message === 'string' ? payload.message : `Request failed: ${response.status}`,
      response.status,
      payload?.code ?? 'REQUEST_FAILED',
      payload?.details,
    );
  }

  private async json<T>(path: string, init: RequestInit = {}): Promise<T> {
    const response = await this.request(path, init);
    try {
      return await response.json() as T;
    } catch {
      throw new ApiRequestError(`Invalid JSON response from ${path}`, response.status, 'INTERNAL_ERROR');
    }
  }

  async getSnapshot(unitId: string): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>(
      `/api/v1/workbooks/${encodeURIComponent(unitId)}/snapshot`,
    );
  }

  async createWorkbook(snapshot: WorkbookSnapshotV1): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>('/api/v1/workbooks', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ snapshot }),
    });
  }

  async createEmptyWorkbook(name = 'Untitled workbook'): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>('/api/v1/workbooks', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name }),
    });
  }

  async listWorkbooks(): Promise<WorkbookSummary[]> {
    return this.json<WorkbookSummary[]>('/api/v1/workbooks');
  }

  async deleteWorkbook(unitId: string): Promise<void> {
    await this.request(`/api/v1/workbooks/${encodeURIComponent(unitId)}`, { method: 'DELETE' });
  }

  async listRevisions(unitId: string): Promise<RevisionRecord[]> {
    const body = await this.json<{ revisions: RevisionRecord[] }>(`/api/v1/workbooks/${encodeURIComponent(unitId)}/revisions`);
    return body.revisions;
  }

  async getRevisionSnapshot(unitId: string, revision: number): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>(
      `/api/v1/workbooks/${encodeURIComponent(unitId)}/revisions/${revision}/snapshot`,
    );
  }

  async saveSnapshot(unitId: string, snapshot: WorkbookSnapshotV1, baseRevision: number): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>(
      `/api/v1/workbooks/${encodeURIComponent(unitId)}/snapshot`,
      {
        method: 'PUT',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ snapshot, baseRevision }),
      },
    );
  }

  async importXlsxBase64(base64: string, fileName = 'import.xlsx'): Promise<XlsxImportResponse> {
    return this.json<XlsxImportResponse>('/api/v1/files/import-xlsx', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ base64, fileName }),
    });
  }

  async exportXlsx(unitId: string, fileName?: string): Promise<XlsxExportResponse> {
    const query = fileName ? `?fileName=${encodeURIComponent(fileName)}` : '';
    return this.json<XlsxExportResponse>(
      `/api/v1/files/${encodeURIComponent(unitId)}/export${query}`,
    );
  }

  async calculatePivot(unitId: string, pivotId: string): Promise<PivotCalculationResponse> {
    return this.json<PivotCalculationResponse>(`/api/v1/workbooks/${encodeURIComponent(unitId)}/calculations/pivot`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ pivotId }),
    });
  }

  async createDataTable(unitId: string, table: WorkbookTableModel): Promise<WorkbookTableModel> {
    return this.json<WorkbookTableModel>(`/api/v1/workbooks/${encodeURIComponent(unitId)}/tables`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(table),
    });
  }

  async appendDataBlock(unitId: string, tableId: string, startRow: number, rows: TableScalar[][]): Promise<WorkbookTableBlock> {
    return this.json<WorkbookTableBlock>(`/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}/blocks`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ startRow, rows }),
    });
  }

  async deleteDataTable(unitId: string, tableId: string): Promise<void> {
    await this.request(`/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}`, { method: 'DELETE' });
  }

  async readDataRows(unitId: string, tableId: string, offset = 0, limit = 500): Promise<TableRowsResponse> {
    return this.json<TableRowsResponse>(`/api/v1/workbooks/${encodeURIComponent(unitId)}/tables/${encodeURIComponent(tableId)}/rows?offset=${offset}&limit=${limit}`);
  }

}

/**
 * V2 collaboration messages. Client-originated presence/cursor messages do
 * not carry actorId; the server adds it to broadcast messages after token
 * verification.  There is intentionally no V1 message decoder.
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

/** The public collaboration message surface is V2 only. */
export type CollaborationMessage = OperationMessageV2;

export function encodeOperationMessageV2(message: OperationMessageV2): string {
  return JSON.stringify(message);
}

export function encodeClientOperationMessageV2(message: ClientOperationMessageV2): string {
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
    if (!mutation.affectedRanges.every(isRangeRef)) throw new Error(`committed mutation[${index}] contains invalid affectedRanges`);
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
      return { type: 'snapshot.response', payload: validateSnapshotResponse(message.payload) };
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

export function encodeMessage(message: OperationMessageV2): string {
  return encodeOperationMessageV2(message);
}

export function decodeMessage(input: string): OperationMessageV2 {
  return decodeOperationMessageV2(input);
}

/** 连接状态 */
export type CollabSocketStatus = 'connecting' | 'open' | 'closed';

export interface CollabSocketOptions {
  /** Required token source for the browser handshake. */
  authTokenProvider: AuthTokenProvider;
  /** 断线重连基础延迟(毫秒),指数退避 */
  reconnectBaseDelayMs?: number;
  /** 重连延迟上限(毫秒) */
  reconnectMaxDelayMs?: number;
  /** Injectable constructor for deterministic browser/Node tests. */
  webSocketFactory?: (url: string, protocols: string | string[]) => WebSocket;
}

type CollabSocketListener = (message: OperationMessageV2) => void;
type StatusListener = (status: CollabSocketStatus) => void;
type ProtocolErrorListener = (error: Error) => void;

const BEARER_SUBPROTOCOL_PREFIX = 'bearer.';

function encodeBase64Url(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/g, '');
}

/** Browser-safe bearer token transport for a WebSocket handshake. */
export function createBearerSubprotocol(token: string): string {
  const normalized = token.trim();
  if (!normalized) throw new AuthenticationRequiredError();
  return `${BEARER_SUBPROTOCOL_PREFIX}${encodeBase64Url(normalized)}`;
}

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
  private readonly protocolErrorListeners = new Set<ProtocolErrorListener>();
  private reconnectAttempt = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private closedByUser = false;
  private connecting = false;

  constructor(
    private readonly url: string,
    private readonly options: CollabSocketOptions,
  ) {}

  get status(): CollabSocketStatus {
    if (this.socket?.readyState === 1) return 'open';
    if (this.closedByUser) return 'closed';
    return this.socket || this.reconnectTimer || this.connecting ? 'connecting' : 'closed';
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

  /** 发送 V2 客户端消息；未连接时返回 false，由 OfflineQueue 保留 operation。 */
  send(message: ClientOperationMessageV2): boolean {
    const encoded = encodeClientOperationMessageV2(message);
    if (this.socket?.readyState === 1) {
      this.socket.send(encoded);
      return true;
    }
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

  onProtocolError(listener: ProtocolErrorListener): () => void {
    this.protocolErrorListeners.add(listener);
    return () => this.protocolErrorListeners.delete(listener);
  }

  private async connect(): Promise<void> {
    this.connecting = true;
    this.emitStatus('connecting');
    let token: string | null;
    try {
      token = (await this.options.authTokenProvider())?.trim() || null;
    } catch (cause) {
      this.connecting = false;
      this.failClosed(cause instanceof Error ? cause : new Error('Unable to resolve bearer token'));
      return;
    }
    if (!token) {
      this.connecting = false;
      this.failClosed(new AuthenticationRequiredError());
      return;
    }
    if (this.closedByUser) {
      this.connecting = false;
      return;
    }
    const factory = this.options.webSocketFactory ?? ((target: string, protocols: string | string[]) => new WebSocket(target, protocols));
    const socket = factory(this.url, createBearerSubprotocol(token));
    this.socket = socket;
    this.connecting = false;

    socket.onopen = () => {
      this.reconnectAttempt = 0;
      this.emitStatus('open');
    };

    socket.onmessage = (event) => {
      if (typeof event.data !== 'string') {
        this.failClosed(new Error('Collaboration server sent a non-text message'));
        return;
      }
      try {
        const message = decodeOperationMessageV2(event.data);
        for (const listener of this.messageListeners) listener(message);
      } catch (cause) {
        this.failClosed(cause instanceof Error ? cause : new Error('Invalid collaboration server message'));
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

  private failClosed(error: Error): void {
    this.closedByUser = true;
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    for (const listener of this.protocolErrorListeners) listener(error);
    this.socket?.close(1002, error.message.slice(0, 120));
    this.socket = null;
    this.connecting = false;
    this.emitStatus('closed');
  }
}
