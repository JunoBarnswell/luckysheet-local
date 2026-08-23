import type { PivotResultTree, RangeRef, TableScalar, WorkbookSnapshot, WorkbookTableBlock, WorkbookTableModel } from '@react-sheets/core-model';

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
export const OPERATION_ENVELOPE_SCHEMA = 'OperationEnvelope' as const;

export interface OperationMutation {
  id: string;
  sheetId: string;
  params: unknown;
}

export interface OperationEnvelope {
  schema: typeof OPERATION_ENVELOPE_SCHEMA;
  operationId: string;
  unitId: string;
  clientSequence: number;
  baseRevision: number;
  mutations: OperationMutation[];
  createdAt: string;
}

/** Server-authored metadata added after authentication, validation and commit. */
export interface CommittedOperationMutation extends OperationMutation {
  /** Authoritative range calculated by the server; never read from client input. */
  affectedRanges: RangeRef[];
}

export interface CommittedOperationEnvelope extends Omit<OperationEnvelope, 'mutations'> {
  actorId: string;
  revision: number;
  committedAt: string;
  mutations: CommittedOperationMutation[];
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
  operation: CommittedOperationEnvelope;
}

/**
 * Authentication is deliberately supplied by the host application.  The
 * protocol package never reads cookies, localStorage or a client-provided
 * actor id.  A missing/empty token is an authentication failure, not an
 * unauthenticated fallback.
 */
export type AuthTokenProvider = () => string | null | Promise<string | null>;
/** Opaque, server-issued workbook sharing token for a guest session. */
export type ShareTokenProvider = () => string | null | Promise<string | null>;

export interface WorkbookApiClientOptions {
  baseUrl?: string;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
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

  constructor(message = 'Authentication or workbook share token is required') {
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

/** Validate the only snapshot representation accepted at the wire boundary. */
export function validateWorkbookSnapshot(value: unknown): WorkbookSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('WorkbookSnapshot must be an object');
  }
  const input = value as Record<string, unknown>;
  if (input.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
  if (!isNonEmptyString(input.unitId) || !isNonEmptyString(input.name)) {
    throw new Error('WorkbookSnapshot requires unitId and name');
  }
  if (!Array.isArray(input.sheets) || input.sheets.length === 0) {
    throw new Error('WorkbookSnapshot requires at least one sheet');
  }
  for (const [index, rawSheet] of input.sheets.entries()) {
    if (!rawSheet || typeof rawSheet !== 'object' || Array.isArray(rawSheet)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] must be an object`);
    }
    const sheet = rawSheet as Record<string, unknown>;
    if (!isNonEmptyString(sheet.id) || !isNonEmptyString(sheet.name)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] requires id and name`);
    }
    if (!Number.isSafeInteger(sheet.rowCount) || Number(sheet.rowCount) <= 0
      || !Number.isSafeInteger(sheet.columnCount) || Number(sheet.columnCount) <= 0
      || !sheet.cells || typeof sheet.cells !== 'object'
      || !Array.isArray(sheet.merges)
      || !Array.isArray(sheet.pivots)
      || !Array.isArray(sheet.sparklines)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] has invalid grid data`);
    }
    if ('charts' in sheet || 'shapes' in sheet || 'images' in sheet) {
      throw new Error(`WorkbookSnapshot sheet[${index}] contains legacy drawing collections`);
    }
    if (!Array.isArray(sheet.drawings) || !sheet.drawingPayloads || typeof sheet.drawingPayloads !== 'object') {
      throw new Error(`WorkbookSnapshot sheet[${index}] requires canonical drawings and payloads`);
    }
  }
  return value as WorkbookSnapshot;
}

function validateSnapshotResponse(value: unknown): SnapshotResponse {
  if (!value || typeof value !== 'object') throw new Error('snapshot.response payload must be an object');
  const input = value as Record<string, unknown>;
  if (!input.snapshot || typeof input.snapshot !== 'object') throw new Error('snapshot.response requires snapshot');
  const snapshot = validateWorkbookSnapshot(input.snapshot);
  if (!Number.isSafeInteger(input.revision) || Number(input.revision) < 0) {
    throw new Error('snapshot.response requires a valid revision');
  }
  return { snapshot, revision: Number(input.revision) };
}

/** Strict runtime validation used at the REST/WebSocket trust boundary. */
export function validateOperationEnvelope(value: unknown): OperationEnvelope {
  if (!value || typeof value !== 'object') throw new Error('OperationEnvelope must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_SCHEMA) throw new Error('Unsupported operation schema');
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
  // of silently ignoring them. This prevents accidental reintroduction of
  // obsolete semantics through an untyped JSON caller.
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
    } satisfies OperationMutation;
  });

  return {
    schema: OPERATION_ENVELOPE_SCHEMA,
    operationId: input.operationId,
    unitId: input.unitId,
    clientSequence: Number(input.clientSequence),
    baseRevision: Number(input.baseRevision),
    mutations,
    createdAt: input.createdAt,
  };
}

export interface SnapshotResponse {
  snapshot: WorkbookSnapshot;
  revision: number;
}

/**
 * The only client-authored history restore request.  The historical snapshot
 * is deliberately absent: the server resolves targetRevision from its own
 * immutable revision log and creates the restore operation after ACL checks.
 */
export interface HistoryRestoreRequest {
  targetRevision: number;
  reason?: string;
}

/** Server response after committing a server-generated workbook.restore op. */
export interface HistoryRestoreResponse {
  operation: CommittedOperationEnvelope;
  snapshot: SnapshotResponse;
}

export interface HistoryAuditRecord {
  auditId: string;
  unitId: string;
  operationId: string;
  actorId: string;
  action: 'workbook.restore';
  targetRevision: number;
  revision: number;
  reason?: string;
  createdAt: string;
}

/** Strict validation for the REST history restore trust boundary. */
export function validateHistoryRestoreRequest(value: unknown): HistoryRestoreRequest {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('History restore request must be an object');
  }
  const input = value as Record<string, unknown>;
  const unknownKeys = Object.keys(input).filter((key) => key !== 'targetRevision' && key !== 'reason');
  if (unknownKeys.length > 0) {
    throw new Error(`History restore request contains unsupported fields: ${unknownKeys.join(', ')}`);
  }
  if (!Number.isSafeInteger(input.targetRevision) || Number(input.targetRevision) < 0) {
    throw new Error('targetRevision must be a non-negative safe integer');
  }
  if (input.reason !== undefined && (typeof input.reason !== 'string' || input.reason.length > 1000)) {
    throw new Error('reason must be a string with at most 1000 characters');
  }
  return {
    targetRevision: Number(input.targetRevision),
    ...(input.reason === undefined ? {} : { reason: input.reason }),
  };
}

export interface CompatibilityReportPayload {
  schema: 'CompatibilityReport';
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

/** Sanitized intent for a server-executed database or credentialed REST query. */
export interface ServerQueryRequest {
  queryId: string;
  name: string;
  connectorId: 'sqlite' | 'jdbc' | 'rest';
  sourceRef: string;
  statement: string;
  method?: 'GET' | 'POST';
  body?: unknown;
  parameters?: unknown[];
  steps: Array<{ id: string; kind: string; name: string; config: Record<string, unknown>; enabled: boolean }>;
}

export interface ServerQueryResponse {
  queryId: string;
  connectorId: string;
  sourceRef: string;
  sourceRevision: number;
  columns: string[];
  rows: TableScalar[][];
  rowCount: number;
  executedAt: string;
  durationMs: number;
}

export type GuestShareRole = 'viewer' | 'commenter' | 'editor';

export interface GuestShareRequest {
  role: GuestShareRole;
  expiresAt?: string;
}

export interface GuestShareResponse {
  shareId: string;
  unitId: string;
  role: GuestShareRole;
  expiresAt: string;
  revokedAt?: string;
  createdBy: string;
  createdAt: string;
  /** Returned exactly once when a share is created. */
  token?: string;
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
  payload: CommittedOperationEnvelope;
}

export class WorkbookApiClient {
  private readonly baseUrl: string;
  private readonly authTokenProvider?: AuthTokenProvider;
  private readonly shareTokenProvider?: ShareTokenProvider;
  private readonly fetchImpl: typeof fetch;

  constructor(options: WorkbookApiClientOptions = {}) {
    this.baseUrl = options.baseUrl ?? '';
    this.authTokenProvider = options.authTokenProvider;
    this.shareTokenProvider = options.shareTokenProvider;
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  private async request(path: string, init: RequestInit = {}): Promise<Response> {
    const token = (await this.authTokenProvider?.())?.trim();
    const shareToken = token ? undefined : (await this.shareTokenProvider?.())?.trim();
    if (!token && !shareToken) throw new AuthenticationRequiredError();
    const headers = new Headers(init.headers);
    if (token) headers.set('authorization', `Bearer ${token}`);
    else headers.set('x-workbook-share-token', shareToken!);
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
      `/api/workbooks/${encodeURIComponent(unitId)}/snapshot`,
    );
  }

  async createWorkbook(snapshot: WorkbookSnapshot): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>('/api/workbooks', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ unitId: snapshot.unitId, name: snapshot.name, snapshot }),
    });
  }

  async listWorkbooks(): Promise<WorkbookSummary[]> {
    return this.json<WorkbookSummary[]>('/api/workbooks');
  }

  async listRevisions(unitId: string): Promise<RevisionRecord[]> {
    return this.json<RevisionRecord[]>(`/api/workbooks/${encodeURIComponent(unitId)}/revisions`);
  }

  async getRevisionSnapshot(unitId: string, revision: number): Promise<SnapshotResponse> {
    return this.json<SnapshotResponse>(
      `/api/workbooks/${encodeURIComponent(unitId)}/revisions/${revision}/snapshot`,
    );
  }

  async restoreToRevision(unitId: string, targetRevision: number, reason?: string): Promise<HistoryRestoreResponse> {
    const body = validateHistoryRestoreRequest({ targetRevision, ...(reason === undefined ? {} : { reason }) });
    return this.json<HistoryRestoreResponse>(
      `/api/workbooks/${encodeURIComponent(unitId)}/restore`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      },
    );
  }

  async executeServerQuery(unitId: string, request: ServerQueryRequest): Promise<ServerQueryResponse> {
    return this.json<ServerQueryResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/queries/execute`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(request),
    });
  }

  async cancelServerQuery(unitId: string, queryId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/queries/${encodeURIComponent(queryId)}/cancel`, {
      method: 'POST',
    });
  }

  async createGuestShare(unitId: string, request: GuestShareRequest): Promise<GuestShareResponse> {
    return this.json<GuestShareResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/shares`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(request),
    });
  }

  async listGuestShares(unitId: string): Promise<GuestShareResponse[]> {
    return this.json<GuestShareResponse[]>(`/api/workbooks/${encodeURIComponent(unitId)}/shares`);
  }

  async revokeGuestShare(unitId: string, shareId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/shares/${encodeURIComponent(shareId)}`, {
      method: 'DELETE',
    });
  }

}

/**
 * Collaboration messages. Client-originated presence/cursor messages do
 * not carry actorId; the server adds it to broadcast messages after token
 * verification. There is intentionally no legacy message decoder.
 */
export type OperationMessage =
  | { type: 'snapshot.request'; unitId: string }
  | { type: 'snapshot.response'; payload: SnapshotResponse }
  | { type: 'changeset.submit'; payload: OperationEnvelope }
  | { type: 'changeset.ack'; operationId: string; revision: number }
  | { type: 'changeset.reject'; operationId: string; error: ApiError }
  | { type: 'revision.created'; payload: CommittedOperationEnvelope; revision: number }
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown }
  | { type: 'presence.broadcast'; unitId: string; actorId: string; state: unknown }
  | { type: 'cursor.broadcast'; unitId: string; actorId: string; state: unknown };

export type ClientOperationMessage =
  | { type: 'snapshot.request'; unitId: string }
  | { type: 'changeset.submit'; payload: OperationEnvelope }
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown };

/** The public collaboration message surface is the single operation contract. */
export type CollaborationMessage = OperationMessage;

export function encodeOperationMessage(message: OperationMessage): string {
  return JSON.stringify(message);
}

export function encodeClientOperationMessage(message: ClientOperationMessage): string {
  return JSON.stringify(message);
}

function validateCommittedOperationEnvelope(value: unknown): CommittedOperationEnvelope {
  if (!value || typeof value !== 'object') throw new Error('Committed operation must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_SCHEMA || !isNonEmptyString(input.operationId) || !isNonEmptyString(input.unitId)) {
    throw new Error('Invalid committed operation schema');
  }
  if (!isNonEmptyString(input.actorId) || !Number.isSafeInteger(input.revision) || Number(input.revision) < 1) {
    throw new Error('Invalid committed operation metadata');
  }
  if (!isNonEmptyString(input.committedAt) || Number.isNaN(Date.parse(input.committedAt))) {
    throw new Error('Invalid committed operation timestamp');
  }
  const operation = validateOperationEnvelope({
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

export function decodeOperationMessage(input: string): OperationMessage {
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
      return { type: 'changeset.submit', payload: validateOperationEnvelope(message.payload) };
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
        payload: validateCommittedOperationEnvelope(message.payload),
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

export function decodeClientOperationMessage(input: string): ClientOperationMessage {
  const message = decodeOperationMessage(input);
  if (message.type === 'changeset.submit' || message.type === 'snapshot.request'
    || message.type === 'presence.updated' || message.type === 'cursor.updated') return message;
  throw new Error(`Server-only collaboration message: ${message.type}`);
}

export function encodeMessage(message: OperationMessage): string {
  return encodeOperationMessage(message);
}

export function decodeMessage(input: string): OperationMessage {
  return decodeOperationMessage(input);
}

/** 连接状态 */
export type CollabSocketStatus = 'connecting' | 'open' | 'closed';

export interface CollabSocketOptions {
  /** OIDC source for an authenticated browser handshake. */
  authTokenProvider?: AuthTokenProvider;
  /** Server-issued share token source for an anonymous guest handshake. */
  shareTokenProvider?: ShareTokenProvider;
  /** 断线重连基础延迟(毫秒),指数退避 */
  reconnectBaseDelayMs?: number;
  /** 重连延迟上限(毫秒) */
  reconnectMaxDelayMs?: number;
  /** Injectable constructor for deterministic browser/Node tests. */
  webSocketFactory?: (url: string, protocols: string | string[]) => WebSocket;
}

type CollabSocketListener = (message: OperationMessage) => void;
type StatusListener = (status: CollabSocketStatus) => void;
type ProtocolErrorListener = (error: Error) => void;

const BEARER_SUBPROTOCOL_PREFIX = 'bearer.';

function encodeBase64Url(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/g, '');
}

function withShareToken(url: string, shareToken: string): string {
  const target = new URL(url, typeof window === 'undefined' ? 'ws://localhost' : window.location.origin);
  target.searchParams.set('shareToken', shareToken);
  return target.toString();
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

  /** 发送客户端消息；未连接时返回 false，由 OfflineQueue 保留 operation。 */
  send(message: ClientOperationMessage): boolean {
    const encoded = encodeClientOperationMessage(message);
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
    let shareToken: string | null;
    try {
      token = (await this.options.authTokenProvider?.())?.trim() || null;
      shareToken = token ? null : (await this.options.shareTokenProvider?.())?.trim() || null;
    } catch (cause) {
      this.connecting = false;
      this.failClosed(cause instanceof Error ? cause : new Error('Unable to resolve bearer token'));
      return;
    }
    if (!token && !shareToken) {
      this.connecting = false;
      this.failClosed(new AuthenticationRequiredError());
      return;
    }
    if (this.closedByUser) {
      this.connecting = false;
      return;
    }
    const factory = this.options.webSocketFactory ?? ((target: string, protocols: string | string[]) => new WebSocket(target, protocols));
    const socket = token
      ? factory(this.url, createBearerSubprotocol(token))
      : factory(withShareToken(this.url, shareToken!), []);
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
        const message = decodeOperationMessage(event.data);
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
