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

export interface WorkbookSummary {
  unitId: string;
  name: string;
  revision: number;
  updatedAt: string;
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
