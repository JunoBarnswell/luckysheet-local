export type WorkspaceMemoryBucket =
  | 'workspaceHeads'
  | 'workspaceSnapshots'
  | 'workspaceOperations'
  | 'workspaceCatalog'
  | 'operationJournals'
  | 'dataBlocks'
  | 'sparseOverlays'
  | 'nativeDocuments'
  | 'assets';

export interface WorkspaceMemoryState {
  workspaceHeads: Map<string, unknown>;
  workspaceSnapshots: Map<string, unknown>;
  workspaceOperations: Map<string, unknown>;
  workspaceCatalog: Map<string, unknown>;
  operationJournals: Map<string, unknown>;
  dataBlocks: Map<string, unknown>;
  sparseOverlays: Map<string, unknown>;
  nativeDocuments: Map<string, unknown>;
  assets: Map<string, unknown>;
}

export interface WorkspaceMemoryTransaction {
  get<T>(bucket: WorkspaceMemoryBucket, key: string): T | undefined;
  getAll<T>(bucket: WorkspaceMemoryBucket): T[];
  set<T>(bucket: WorkspaceMemoryBucket, key: string, value: T): void;
  delete(bucket: WorkspaceMemoryBucket, key: string): void;
}

export type WorkspacePersistenceState = 'ready' | 'disposed';
export type WorkspacePersistenceMode = 'memory';

export type WorkspaceStorageErrorCode =
  | 'STORAGE_MEMORY_DISPOSED'
  | 'STORAGE_MEMORY_TRANSACTION_FAILED'
  | 'STORAGE_WRITER_UNAVAILABLE'
  | 'STORAGE_REVISION_CONFLICT'
  | 'STORAGE_TRANSACTION_FAILED'
  | 'STORAGE_SCHEMA_INVALID';

export class WorkspaceStorageError extends Error {
  readonly code: WorkspaceStorageErrorCode;
  readonly operation: string;
  readonly recovery: string;

  constructor(input: {
    code: WorkspaceStorageErrorCode;
    operation: string;
    message: string;
    recovery: string;
    cause?: unknown;
  }) {
    super(input.message, input.cause === undefined ? undefined : { cause: input.cause });
    this.name = 'WorkspaceStorageError';
    this.code = input.code;
    this.operation = input.operation;
    this.recovery = input.recovery;
  }
}

export function isWorkspaceStorageError(error: unknown): error is WorkspaceStorageError {
  if (!error || typeof error !== 'object' || !('code' in error)) return false;
  return [
    'STORAGE_MEMORY_DISPOSED',
    'STORAGE_MEMORY_TRANSACTION_FAILED',
    'STORAGE_WRITER_UNAVAILABLE',
    'STORAGE_REVISION_CONFLICT',
    'STORAGE_TRANSACTION_FAILED',
    'STORAGE_SCHEMA_INVALID',
  ].includes(String((error as { code: unknown }).code));
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

function createState(): WorkspaceMemoryState {
  return {
    workspaceHeads: new Map(),
    workspaceSnapshots: new Map(),
    workspaceOperations: new Map(),
    workspaceCatalog: new Map(),
    operationJournals: new Map(),
    dataBlocks: new Map(),
    sparseOverlays: new Map(),
    nativeDocuments: new Map(),
    assets: new Map(),
  };
}

function cloneState(source: WorkspaceMemoryState): WorkspaceMemoryState {
  const target = createState();
  for (const bucket of Object.keys(target) as WorkspaceMemoryBucket[]) {
    for (const [key, value] of source[bucket]) target[bucket].set(key, clone(value));
  }
  return target;
}

function memoryTransaction(state: WorkspaceMemoryState): WorkspaceMemoryTransaction {
  return {
    get<T>(bucket: WorkspaceMemoryBucket, key: string) {
      const value = state[bucket].get(key);
      return value === undefined ? undefined : value as T;
    },
    getAll<T>(bucket: WorkspaceMemoryBucket) {
      return [...state[bucket].values()] as T[];
    },
    set<T>(bucket: WorkspaceMemoryBucket, key: string, value: T) {
      state[bucket].set(key, value);
    },
    delete(bucket: WorkspaceMemoryBucket, key: string) {
      state[bucket].delete(key);
    },
  };
}

/**
 * Page-session storage owner. Every local Store in one ApplicationServicesProvider
 * receives this same context. A transaction commits by replacing the complete
 * staged state, so a failed multi-store mutation cannot partially persist.
 */
export class WorkspaceMemoryCoordinator {
  private stateValue = createState();
  private stateValueStatus: WorkspacePersistenceState = 'ready';
  private readonly activeWriters = new Set<string>();
  private transactionTail: Promise<void> = Promise.resolve();

  get state(): WorkspacePersistenceState { return this.stateValueStatus; }

  get mode(): WorkspacePersistenceMode { return 'memory'; }

  ensureReady(): void {
    if (this.stateValueStatus === 'disposed') throw this.disposedError('ensure-ready');
  }

  async read<T>(operation: (transaction: WorkspaceMemoryTransaction) => T | Promise<T>): Promise<T> {
    const read = this.transactionTail.then(async () => {
      this.ensureReady();
      try {
        const result = await operation(memoryTransaction(this.stateValue));
        return clone(result);
      } catch (cause) {
        throw this.transactionError('read', cause);
      }
    });
    await read.catch(() => undefined);
    return read;
  }

  async transaction<T>(operation: (transaction: WorkspaceMemoryTransaction) => T | Promise<T>): Promise<T> {
    const run = this.transactionTail.then(async () => {
      this.ensureReady();
      const staged = cloneState(this.stateValue);
      try {
        const result = await operation(memoryTransaction(staged));
        this.ensureReady();
        this.stateValue = staged;
        return clone(result);
      } catch (cause) {
        if (cause instanceof WorkspaceStorageError) throw cause;
        throw this.transactionError('transaction', cause);
      }
    });
    this.transactionTail = run.then(() => undefined, () => undefined);
    return run;
  }

  async withWorkbookWriter<T>(unitId: string, operation: () => Promise<T>): Promise<T> {
    this.ensureReady();
    if (this.activeWriters.has(unitId)) {
      throw new WorkspaceStorageError({
        code: 'STORAGE_WRITER_UNAVAILABLE',
        operation: 'acquire-writer',
        message: `工作簿正在由当前页面的其他编辑会话写入：${unitId}`,
        recovery: '请等待当前写入完成后重试；当前工作簿不会被覆盖。',
      });
    }
    this.activeWriters.add(unitId);
    try {
      return await operation();
    } finally {
      this.activeWriters.delete(unitId);
    }
  }

  async disposeAsync(): Promise<void> {
    if (this.stateValueStatus === 'disposed') return;
    this.stateValueStatus = 'disposed';
    this.activeWriters.clear();
    this.stateValue = createState();
  }

  private disposedError(operation: string): WorkspaceStorageError {
    return new WorkspaceStorageError({
      code: 'STORAGE_MEMORY_DISPOSED',
      operation,
      message: '当前页面内存工作簿会话已经结束。',
      recovery: '请返回文件中心或刷新页面，重新创建或导入工作簿。',
    });
  }

  private transactionError(operation: string, cause: unknown): WorkspaceStorageError {
    return new WorkspaceStorageError({
      code: 'STORAGE_MEMORY_TRANSACTION_FAILED',
      operation,
      message: '内存工作簿事务失败，未提交任何部分数据。',
      recovery: '请重试当前操作；如果问题持续，请重新开始页面内存会话。',
      cause,
    });
  }
}

export function memoryKey(...parts: readonly (string | number)[]): string {
  return parts.map((part) => String(part).replaceAll('\u0000', '\u0000\u0000')).join('\u0000');
}
