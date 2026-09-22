import { migrateStoredWorkbookSnapshot, type WorkbookSnapshot } from '@react-sheets/core-model';
import type { NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import type { OperationEnvelope } from '@react-sheets/protocol';
import { computeChecksum, verifyChecksum } from './checksum';
import { LocalDataBlockStore } from './data-block-store';
import { buildNativeDocumentRecord, LocalNativeDocumentStore } from './native-document-store';
import { LocalSparseOverlayStore } from '../data-source/overlay-store';
import type { AssetStore } from './asset-store';
import { normalizeWorkspaceRecordWithAssets } from './asset-migration';
import {
  WorkspaceMemoryCoordinator,
  WorkspaceStorageError,
  memoryKey,
  type WorkspacePersistenceMode,
  type WorkspacePersistenceState,
} from './memory';

export interface WorkspacePersistenceOptions {
  mode?: WorkspacePersistenceMode;
  unitId?: string | (() => string);
}

export type WorkspaceLifecycle = 'active' | 'trashed';
export type WorkspaceStorageLocation = 'local' | 'remote' | 'mirrored';
export type WorkspaceSource = 'native' | 'document-import';
export type WorkspaceRole = 'owner' | 'editor' | 'commenter' | 'viewer';

export interface WorkspaceRecordMetadata {
  location: WorkspaceStorageLocation;
  lifecycle: WorkspaceLifecycle;
  source: WorkspaceSource;
  role: WorkspaceRole;
  ownerId?: string;
  sourceFileName?: string;
  spaceId?: string;
  folderId?: string;
  locationPath?: string;
  deletedAt?: string;
}

export interface WorkspaceUserState {
  lastOpenedAt?: string;
  favorite: boolean;
  defaultLocation?: { spaceId?: string; folderId?: string };
  autoSave: boolean;
  autoSync: boolean;
  offlineCache: boolean;
  importCompatibility: 'A' | 'B' | 'C';
  language?: string;
  theme?: 'light' | 'dark' | 'system';
}

/** The only browser-persistent workbook record. */
export interface WorkspaceRecord {
  schema: 'WorkspaceRecord';
  unitId: string;
  snapshot: WorkbookSnapshot;
  checksum: string;
  localRevision: number;
  serverRevision: number;
  storageRevision: number;
  syncMode: 'remote' | 'local-only';
  pending: PendingOperationJournal;
  updatedAt: string;
  metadata: WorkspaceRecordMetadata;
  userState: WorkspaceUserState;
}

export interface PendingOperationJournal {
  schema: 'PendingOperationJournal';
  unitId: string;
  nextClientSequence: number;
  operations: OperationEnvelope[];
  checksum: string;
}

export interface WorkspaceRecordInput {
  unitId: string;
  snapshot: WorkbookSnapshot;
  localRevision: number;
  serverRevision: number;
  storageRevision?: number;
  syncMode: 'remote' | 'local-only';
  operations: readonly OperationEnvelope[];
  nextClientSequence: number;
  updatedAt?: string;
  metadata?: Partial<WorkspaceRecordMetadata>;
  userState?: Partial<WorkspaceUserState>;
}

const DEFAULT_WORKSPACE_METADATA: WorkspaceRecordMetadata = {
  location: 'local',
  lifecycle: 'active',
  source: 'native',
  role: 'owner',
};

const DEFAULT_WORKSPACE_USER_STATE: WorkspaceUserState = {
  favorite: false,
  autoSave: true,
  autoSync: true,
  offlineCache: true,
  importCompatibility: 'B',
};

export function normalizeWorkspaceRecord(record: WorkspaceRecord): WorkspaceRecord {
  const snapshot = migrateStoredWorkbookSnapshot(record.snapshot);
  return {
    ...clone(record),
    snapshot,
    checksum: computeChecksum(snapshotPayload(snapshot)),
    metadata: {
      ...DEFAULT_WORKSPACE_METADATA,
      ...(record.metadata ?? {}),
    },
    userState: {
      ...DEFAULT_WORKSPACE_USER_STATE,
      ...(record.userState ?? {}),
    },
  };
}

export function buildWorkspaceUserState(input: Partial<WorkspaceUserState> = {}): WorkspaceUserState {
  return { ...DEFAULT_WORKSPACE_USER_STATE, ...input };
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

function snapshotPayload(snapshot: WorkbookSnapshot): string {
  return JSON.stringify(snapshot);
}

function journalPayload(
  unitId: string,
  nextClientSequence: number,
  operations: readonly OperationEnvelope[],
): Omit<PendingOperationJournal, 'checksum'> {
  return {
    schema: 'PendingOperationJournal',
    unitId,
    nextClientSequence,
    operations: operations.map((operation) => clone(operation)),
  };
}

function buildJournal(
  unitId: string,
  nextClientSequence: number,
  operations: readonly OperationEnvelope[],
): PendingOperationJournal {
  const payload = journalPayload(unitId, nextClientSequence, operations);
  return { ...payload, checksum: computeChecksum(JSON.stringify(payload)) };
}

interface WorkspaceHeadRecord {
  schema: 'WorkspaceHead';
  unitId: string;
  snapshotRevision: number;
  checkpointRevision: number;
  localRevision: number;
  serverRevision: number;
  syncMode: 'remote' | 'local-only';
  storageRevision: number;
  nextClientSequence: number;
  updatedAt: string;
}

interface WorkspaceSnapshotRecord {
  schema: 'WorkspaceSnapshot';
  unitId: string;
  revision: number;
  snapshot: unknown;
  checksum: string;
  localRevision: number;
  serverRevision: number;
  syncMode: 'remote' | 'local-only';
  updatedAt: string;
}

interface WorkspaceOperationRecord extends OperationEnvelope {
  unitId: string;
}

interface WorkspaceCatalogRecord {
  schema: 'WorkspaceCatalog';
  unitId: string;
  metadata: unknown;
  userState: unknown;
  updatedAt: string;
}

function headRecordFrom(record: WorkspaceRecord, storageRevision: number): WorkspaceHeadRecord {
  return {
    schema: 'WorkspaceHead',
    unitId: record.unitId,
    snapshotRevision: record.localRevision,
    checkpointRevision: record.localRevision,
    localRevision: record.localRevision,
    serverRevision: record.serverRevision,
    syncMode: record.syncMode,
    storageRevision,
    nextClientSequence: record.pending.nextClientSequence,
    updatedAt: record.updatedAt,
  };
}

function snapshotRecordFrom(record: WorkspaceRecord): WorkspaceSnapshotRecord {
  return {
    schema: 'WorkspaceSnapshot',
    unitId: record.unitId,
    revision: record.localRevision,
    snapshot: clone(record.snapshot),
    checksum: record.checksum,
    localRevision: record.localRevision,
    serverRevision: record.serverRevision,
    syncMode: record.syncMode,
    updatedAt: record.updatedAt,
  };
}

function catalogRecordFrom(record: WorkspaceRecord): WorkspaceCatalogRecord {
  return {
    schema: 'WorkspaceCatalog',
    unitId: record.unitId,
    metadata: clone(record.metadata),
    userState: clone(record.userState),
    updatedAt: record.updatedAt,
  };
}

function schemaError(unitId: string): WorkspaceStorageError {
  return new WorkspaceStorageError({
    code: 'STORAGE_SCHEMA_INVALID',
    operation: 'load-workspace',
    message: `工作簿记录结构校验失败：${unitId}`,
    recovery: '请保留浏览器数据并联系管理员执行显式数据恢复。',
  });
}

function revisionConflict(unitId: string, expected: number, actual: number): WorkspaceStorageError {
  return new WorkspaceStorageError({
    code: 'STORAGE_REVISION_CONFLICT',
    operation: 'write-workspace',
    message: `工作簿本地版本冲突：${unitId}（期望 ${expected}，实际 ${actual}）。`,
    recovery: '请重新加载工作簿并重试，当前版本不会被覆盖。',
  });
}

function isSafeNonNegative(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

export function verifyPendingOperationJournal(journal: PendingOperationJournal): boolean {
  if (!journal || typeof journal !== 'object' || journal.schema !== 'PendingOperationJournal' || typeof journal.unitId !== 'string' || !journal.unitId.trim()) return false;
  if (!isSafeNonNegative(journal.nextClientSequence) || !journal.checksum) return false;
  if (!Array.isArray(journal.operations)) return false;
  const seen = new Set<string>();
  let previousSequence = 0;
  for (const operation of journal.operations) {
    if (
      !operation
      || typeof operation !== 'object'
      || operation.schema !== 'OperationEnvelope'
      || operation.unitId !== journal.unitId
      || !operation.operationId
      || seen.has(operation.operationId)
      || !Number.isSafeInteger(operation.clientSequence)
      || operation.clientSequence <= previousSequence
      || operation.clientSequence > journal.nextClientSequence
    ) return false;
    seen.add(operation.operationId);
    previousSequence = operation.clientSequence;
  }
  const payload = journalPayload(journal.unitId, journal.nextClientSequence, journal.operations);
  return computeChecksum(JSON.stringify(payload)) === journal.checksum;
}

export function verifyWorkspaceRecord(record: WorkspaceRecord): boolean {
  if (
    !record
    || typeof record !== 'object'
    || record.schema !== 'WorkspaceRecord'
    || typeof record.unitId !== 'string'
    || !record.unitId.trim()
    || !record.snapshot
    || typeof record.snapshot !== 'object'
    || record.snapshot.schema !== 'WorkbookSnapshot'
    || record.snapshot.unitId !== record.unitId
    || !isSafeNonNegative(record.localRevision)
    || !isSafeNonNegative(record.serverRevision)
    || !isSafeNonNegative(record.storageRevision)
    || (record.syncMode !== 'remote' && record.syncMode !== 'local-only')
    || !record.updatedAt
    || !record.checksum
  ) return false;
  return verifyChecksum(snapshotPayload(record.snapshot), record.checksum)
    && verifyPendingOperationJournal(record.pending);
}

export function buildWorkspaceRecord(input: WorkspaceRecordInput): WorkspaceRecord {
  if (!input.unitId.trim()) throw new Error('unitId is required');
  if (!isSafeNonNegative(input.localRevision) || !isSafeNonNegative(input.serverRevision)) {
    throw new Error('Workspace revisions must be non-negative safe integers');
  }
  if (!isSafeNonNegative(input.nextClientSequence)) {
    throw new Error('nextClientSequence must be a non-negative safe integer');
  }
  const snapshot = input.snapshot;
  if (snapshot.unitId !== input.unitId) throw new Error('Workspace snapshot unitId does not match record');
  const pending = buildJournal(input.unitId, input.nextClientSequence, input.operations);
  const record: WorkspaceRecord = {
    schema: 'WorkspaceRecord',
    unitId: input.unitId,
    snapshot,
    checksum: computeChecksum(snapshotPayload(snapshot)),
    localRevision: input.localRevision,
    serverRevision: input.serverRevision,
    storageRevision: input.storageRevision ?? 0,
    syncMode: input.syncMode,
    pending,
    updatedAt: input.updatedAt ?? new Date().toISOString(),
    metadata: {
      ...DEFAULT_WORKSPACE_METADATA,
      ...(input.metadata ?? {}),
    },
    userState: buildWorkspaceUserState(input.userState),
  };
  if (!verifyWorkspaceRecord(record)) throw new Error('WorkspaceRecord failed validation');
  return record;
}

/** Synchronous cache used by CollaborationSession callbacks. */
export class OperationJournalStore {
  private readonly journals = new Map<string, PendingOperationJournal>();

  hydrate(record: WorkspaceRecord | null): void {
    if (!record) return;
    if (!verifyWorkspaceRecord(record)) throw new Error(`Invalid WorkspaceRecord: ${record.unitId}`);
    this.journals.set(record.unitId, clone(record.pending));
  }

  write(unitId: string, operations: readonly OperationEnvelope[], nextClientSequence: number): void {
    const journal = buildJournal(unitId, nextClientSequence, operations);
    if (!verifyPendingOperationJournal(journal)) throw new Error('PendingOperationJournal failed validation');
    this.journals.set(unitId, journal);
  }

  read(unitId: string): PendingOperationJournal | null {
    const journal = this.journals.get(unitId);
    return journal ? clone(journal) : null;
  }

  clear(unitId: string): void {
    this.journals.delete(unitId);
  }
}

/** Canonical workspace record storage backed by the page-session memory root. */
export class MemoryWorkspaceStore {
  constructor(private readonly coordinator: WorkspaceMemoryCoordinator) {}

  async load(unitId: string, assetStore?: AssetStore): Promise<WorkspaceRecord | null> {
    const records = await this.readRecords(unitId);
    const value = records[0] ?? null;
    return assetStore && value ? await normalizeWorkspaceRecordWithAssets(value, assetStore) : value;
  }

  async save(record: WorkspaceRecord): Promise<WorkspaceRecord> {
    if (!verifyWorkspaceRecord(record)) throw new Error(`Invalid WorkspaceRecord: ${record.unitId}`);
    const normalized = normalizeWorkspaceRecord(record);
    return this.coordinator.transaction((transaction) => {
    const existing = transaction.get<WorkspaceHeadRecord>('workspaceHeads', normalized.unitId);
    const currentRevision = existing?.storageRevision ?? 0;
    if (existing && currentRevision !== normalized.storageRevision) {
      throw revisionConflict(normalized.unitId, normalized.storageRevision, currentRevision);
    }
    const nextStorageRevision = currentRevision + 1;
    transaction.set('workspaceSnapshots', memoryKey(normalized.unitId, normalized.localRevision), snapshotRecordFrom(normalized));
    transaction.set('workspaceCatalog', normalized.unitId, catalogRecordFrom(normalized));
    const existingOperations = transaction.getAll<WorkspaceOperationRecord>('workspaceOperations');
    for (const operation of existingOperations) {
      if (operation.unitId === normalized.unitId) transaction.delete('workspaceOperations', memoryKey(operation.unitId, operation.clientSequence));
    }
    for (const operation of normalized.pending.operations) {
      transaction.set('workspaceOperations', memoryKey(normalized.unitId, operation.clientSequence), { ...clone(operation), unitId: normalized.unitId });
    }
    transaction.set('workspaceHeads', normalized.unitId, headRecordFrom(normalized, nextStorageRevision));
    return { ...normalized, storageRevision: nextStorageRevision };
    });
  }

  async clear(unitId: string): Promise<void> {
    await this.coordinator.transaction((transaction) => {
      transaction.delete('workspaceHeads', unitId);
      transaction.delete('workspaceCatalog', unitId);
      transaction.delete('nativeDocuments', unitId);
      for (const snapshot of transaction.getAll<WorkspaceSnapshotRecord>('workspaceSnapshots')) {
        if (snapshot.unitId === unitId) transaction.delete('workspaceSnapshots', memoryKey(unitId, snapshot.revision));
      }
      for (const operation of transaction.getAll<WorkspaceOperationRecord>('workspaceOperations')) {
        if (operation.unitId === unitId) transaction.delete('workspaceOperations', memoryKey(unitId, operation.clientSequence));
      }
    });
  }

  async list(): Promise<WorkspaceRecord[]> {
    return this.readRecords();
  }

  async patchCatalog(
    unitId: string,
    patch: { metadata?: Partial<WorkspaceRecordMetadata>; userState?: Partial<WorkspaceUserState> },
    expectedStorageRevision?: number,
  ): Promise<WorkspaceRecord> {
    return this.coordinator.transaction((transaction) => {
    const head = transaction.get<WorkspaceHeadRecord>('workspaceHeads', unitId);
    const catalog = transaction.get<WorkspaceCatalogRecord>('workspaceCatalog', unitId);
    if (!head || !catalog) throw new Error(`Unknown local workbook: ${unitId}`);
    if (expectedStorageRevision !== undefined && head.storageRevision !== expectedStorageRevision) {
      throw revisionConflict(unitId, expectedStorageRevision, head.storageRevision);
    }
    transaction.set('workspaceCatalog', unitId, {
      ...catalog,
      metadata: { ...(catalog.metadata as Partial<WorkspaceRecordMetadata>), ...(patch.metadata ?? {}) },
      userState: { ...(catalog.userState as Partial<WorkspaceUserState>), ...(patch.userState ?? {}) },
      updatedAt: new Date().toISOString(),
    });
    transaction.set('workspaceHeads', unitId, { ...head, storageRevision: head.storageRevision + 1, updatedAt: new Date().toISOString() });
    return this.recordFromTransaction(transaction, unitId);
    });
  }

  private async readRecords(unitId?: string): Promise<WorkspaceRecord[]> {
    return this.coordinator.read((transaction) => {
    const heads = transaction.getAll<WorkspaceHeadRecord>('workspaceHeads');
    const snapshots = transaction.getAll<WorkspaceSnapshotRecord>('workspaceSnapshots');
    const operations = transaction.getAll<WorkspaceOperationRecord>('workspaceOperations');
    const catalogs = transaction.getAll<WorkspaceCatalogRecord>('workspaceCatalog');
    return heads.filter((candidate) => unitId === undefined || candidate.unitId === unitId).map((head) => {
      const snapshot = snapshots.find((candidate) => candidate.unitId === head.unitId && candidate.revision === head.snapshotRevision);
      const catalog = catalogs.find((candidate) => candidate.unitId === head.unitId);
      if (!snapshot || !catalog) throw schemaError(head.unitId);
      const pending = buildJournal(head.unitId, head.nextClientSequence, operations
        .filter((candidate) => candidate.unitId === head.unitId)
        .sort((left, right) => left.clientSequence - right.clientSequence));
      const record: WorkspaceRecord = normalizeWorkspaceRecord({
        schema: 'WorkspaceRecord',
        unitId: head.unitId,
        snapshot: snapshot.snapshot as WorkbookSnapshot,
        checksum: snapshot.checksum,
        localRevision: head.localRevision,
        serverRevision: head.serverRevision,
        storageRevision: head.storageRevision,
        syncMode: head.syncMode,
        pending,
        updatedAt: head.updatedAt,
        metadata: catalog.metadata as WorkspaceRecordMetadata,
        userState: catalog.userState as WorkspaceUserState,
      });
      if (!verifyWorkspaceRecord(record)) throw schemaError(head.unitId);
      return record;
    });
    });
  }

  private recordFromTransaction(transaction: import('./memory').WorkspaceMemoryTransaction, unitId: string): WorkspaceRecord {
    const head = transaction.get<WorkspaceHeadRecord>('workspaceHeads', unitId);
    const snapshot = head && transaction.get<WorkspaceSnapshotRecord>('workspaceSnapshots', memoryKey(unitId, head.snapshotRevision));
    const catalog = transaction.get<WorkspaceCatalogRecord>('workspaceCatalog', unitId);
    const operations = transaction.getAll<WorkspaceOperationRecord>('workspaceOperations').filter((candidate) => candidate.unitId === unitId).sort((left, right) => left.clientSequence - right.clientSequence);
    if (!head || !snapshot || !catalog) throw schemaError(unitId);
    const record = normalizeWorkspaceRecord({
      schema: 'WorkspaceRecord', unitId, snapshot: snapshot.snapshot as WorkbookSnapshot, checksum: snapshot.checksum,
      localRevision: head.localRevision, serverRevision: head.serverRevision, storageRevision: head.storageRevision,
      syncMode: head.syncMode, pending: buildJournal(unitId, head.nextClientSequence, operations), updatedAt: head.updatedAt,
      metadata: catalog.metadata as WorkspaceRecordMetadata, userState: catalog.userState as WorkspaceUserState,
    });
    if (!verifyWorkspaceRecord(record)) throw schemaError(unitId);
    return record;
  }

}

export interface LocalWorkspaceSummary {
  unitId: string;
  name: string;
  localRevision: number;
  serverRevision: number;
  syncMode: WorkspaceRecord['syncMode'];
  checksum: string;
  pendingOperationCount: number;
  updatedAt: string;
  metadata: WorkspaceRecordMetadata;
  userState: WorkspaceUserState;
}

function summarizeWorkspace(record: WorkspaceRecord): LocalWorkspaceSummary {
  return {
    unitId: record.unitId,
    name: record.snapshot.name,
    localRevision: record.localRevision,
    serverRevision: record.serverRevision,
    syncMode: record.syncMode,
    checksum: record.checksum,
    pendingOperationCount: record.pending.operations.length,
    updatedAt: record.updatedAt,
    metadata: structuredClone(record.metadata),
    userState: structuredClone(record.userState),
  };
}

/** Public direct API for the Web catalog/PWA local workspace surface. */
export class LocalWorkspaceStore {
  private readonly memory: MemoryWorkspaceStore;

  constructor(coordinator: WorkspaceMemoryCoordinator) { this.memory = new MemoryWorkspaceStore(coordinator); }

  open(unitId: string, assetStore?: AssetStore): Promise<WorkspaceRecord | null> {
    return this.memory.load(unitId, assetStore);
  }

  async create(input: WorkspaceRecordInput | WorkspaceRecord): Promise<WorkspaceRecord> {
    const record = 'pending' in input
      ? clone(input as WorkspaceRecord)
      : buildWorkspaceRecord(input as WorkspaceRecordInput);
    const normalized = normalizeWorkspaceRecord(record);
    return clone(await this.memory.save(normalized));
  }

  async save(record: WorkspaceRecord): Promise<WorkspaceRecord> {
    const normalized = normalizeWorkspaceRecord(record);
    return clone(await this.memory.save(normalized));
  }

  async checkpoint(input: WorkspaceRecordInput): Promise<WorkspaceRecord> {
    return this.create(buildWorkspaceRecord(input));
  }

  async list(): Promise<LocalWorkspaceSummary[]> {
    const records = await this.memory.list();
    return records.map(summarizeWorkspace);
  }

  listRecords(): Promise<WorkspaceRecord[]> {
    return this.memory.list();
  }

  updateMetadata(unitId: string, metadata: Partial<WorkspaceRecordMetadata>, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.memory.patchCatalog(unitId, { metadata }, expectedStorageRevision);
  }

  updateUserState(unitId: string, userState: Partial<WorkspaceUserState>, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.memory.patchCatalog(unitId, { userState }, expectedStorageRevision);
  }

  async moveToTrash(unitId: string, deletedAt = new Date().toISOString(), expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.updateMetadata(unitId, { lifecycle: 'trashed', deletedAt }, expectedStorageRevision);
  }

  async restore(unitId: string, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.updateMetadata(unitId, { lifecycle: 'active', deletedAt: undefined }, expectedStorageRevision);
  }

  delete(unitId: string): Promise<void> {
    return this.memory.clear(unitId);
  }
}

export class WorkspacePersistence {
  readonly operationJournal: OperationJournalStore;
  readonly store: LocalWorkspaceStore;
  readonly dataBlocks: LocalDataBlockStore;
  readonly sparseOverlays: LocalSparseOverlayStore;
  readonly nativeDocuments: LocalNativeDocumentStore;
  readonly coordinator: WorkspaceMemoryCoordinator;

  constructor(options: WorkspacePersistenceOptions = {}, operationJournal = new OperationJournalStore()) {
    this.coordinator = new WorkspaceMemoryCoordinator();
    this.store = new LocalWorkspaceStore(this.coordinator);
    this.dataBlocks = new LocalDataBlockStore(this.coordinator, options.unitId);
    this.sparseOverlays = new LocalSparseOverlayStore({ coordinator: this.coordinator, unitId: options.unitId });
    this.nativeDocuments = new LocalNativeDocumentStore(this.coordinator);
    this.operationJournal = operationJournal;
  }

  get state(): WorkspacePersistenceState { return this.coordinator.state; }

  get mode(): WorkspacePersistenceMode { return this.coordinator.mode; }

  async ensureReady(): Promise<void> { this.coordinator.ensureReady(); }

  async withWorkbookWriter<T>(unitId: string, operation: () => Promise<T>): Promise<T> {
    return this.coordinator.withWorkbookWriter(unitId, operation);
  }

  async load(unitId: string, assetStore?: AssetStore): Promise<WorkspaceRecord | null> {
    const record = await this.store.open(unitId, assetStore);
    if (record) this.operationJournal.hydrate(record);
    else this.operationJournal.clear(unitId);
    return record;
  }

  list(): Promise<LocalWorkspaceSummary[]> {
    return this.store.list();
  }

  listRecords(): Promise<WorkspaceRecord[]> {
    return this.store.listRecords();
  }

  async checkpoint(
    snapshot: WorkbookSnapshot,
    localRevision: number,
    serverRevision: number,
    syncMode: 'remote' | 'local-only',
    pendingJournal = this.operationJournal.read(snapshot.unitId),
    metadata?: Partial<WorkspaceRecordMetadata>,
    userState?: Partial<WorkspaceUserState>,
  ): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(snapshot.unitId, async () => {
      const previous = await this.store.open(snapshot.unitId);
      const record = buildWorkspaceRecord({
        unitId: snapshot.unitId,
        snapshot,
        localRevision,
        serverRevision,
        syncMode,
        storageRevision: previous?.storageRevision ?? 0,
        operations: pendingJournal?.operations ?? [],
        nextClientSequence: pendingJournal?.nextClientSequence ?? 0,
        metadata: { ...(previous?.metadata ?? {}), ...(metadata ?? {}) },
        userState: { ...(previous?.userState ?? {}), ...(userState ?? {}) },
      });
      return this.store.save(record);
    });
  }

  /**
   * Commits the canonical workspace checkpoint and its source native document artifact
   * in one memory transaction.
   */
  async checkpointWithArtifact(
    snapshot: WorkbookSnapshot,
    localRevision: number,
    serverRevision: number,
    syncMode: 'remote' | 'local-only',
    artifact: NativeDocumentArtifact,
    pendingJournal = this.operationJournal.read(snapshot.unitId),
    metadata?: Partial<WorkspaceRecordMetadata>,
    userState?: Partial<WorkspaceUserState>,
  ): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(snapshot.unitId, async () => {
      const previous = await this.store.open(snapshot.unitId);
      const record = buildWorkspaceRecord({
        unitId: snapshot.unitId,
        snapshot,
        localRevision,
        serverRevision,
        syncMode,
        storageRevision: previous?.storageRevision ?? 0,
        operations: pendingJournal?.operations ?? [],
        nextClientSequence: pendingJournal?.nextClientSequence ?? 0,
        metadata: { ...(previous?.metadata ?? {}), ...(metadata ?? {}) },
        userState: { ...(previous?.userState ?? {}), ...(userState ?? {}) },
      });
      const artifactRecord = await buildNativeDocumentRecord(snapshot.unitId, artifact);
      return this.coordinator.transaction((transaction) => {
        const current = transaction.get<WorkspaceHeadRecord>('workspaceHeads', record.unitId);
        const currentStorageRevision = current?.storageRevision ?? 0;
        if (current && currentStorageRevision !== record.storageRevision) {
          throw revisionConflict(record.unitId, record.storageRevision, currentStorageRevision);
        }
        const saved = { ...record, storageRevision: currentStorageRevision + 1 };
        transaction.set('workspaceHeads', saved.unitId, headRecordFrom(saved, saved.storageRevision));
        transaction.set('workspaceSnapshots', memoryKey(saved.unitId, saved.localRevision), snapshotRecordFrom(saved));
        transaction.set('workspaceCatalog', saved.unitId, catalogRecordFrom(saved));
        for (const operation of transaction.getAll<WorkspaceOperationRecord>('workspaceOperations')) {
          if (operation.unitId === saved.unitId) transaction.delete('workspaceOperations', memoryKey(operation.unitId, operation.clientSequence));
        }
        for (const operation of saved.pending.operations) {
          transaction.set('workspaceOperations', memoryKey(saved.unitId, operation.clientSequence), { ...clone(operation), unitId: saved.unitId });
        }
        transaction.set('nativeDocuments', saved.unitId, artifactRecord);
        return saved;
      });
    });
  }

  async commitOperationJournal(
    unitId: string,
    operations: readonly OperationEnvelope[],
    nextClientSequence: number,
    expectedStorageRevision?: number,
  ): Promise<number> {
    return this.withWorkbookWriter(unitId, async () => {
      return this.coordinator.transaction((transaction) => {
      const head = transaction.get<WorkspaceHeadRecord>('workspaceHeads', unitId);
      if (!head) throw new Error(`Unknown local workbook: ${unitId}`);
      if (expectedStorageRevision !== undefined && head.storageRevision !== expectedStorageRevision) {
        throw revisionConflict(unitId, expectedStorageRevision, head.storageRevision);
      }
      const existing = transaction.getAll<WorkspaceOperationRecord>('workspaceOperations');
      for (const operation of existing) if (operation.unitId === unitId) transaction.delete('workspaceOperations', memoryKey(unitId, operation.clientSequence));
      for (const operation of operations) transaction.set('workspaceOperations', memoryKey(unitId, operation.clientSequence), { ...clone(operation), unitId });
      const storageRevision = head.storageRevision + 1;
      transaction.set('workspaceHeads', unitId, { ...head, storageRevision, nextClientSequence, updatedAt: new Date().toISOString() });
      return storageRevision;
      });
    });
  }

  async moveToTrash(unitId: string, deletedAt = new Date().toISOString(), expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(unitId, () => this.store.moveToTrash(unitId, deletedAt, expectedStorageRevision));
  }

  async restore(unitId: string, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(unitId, () => this.store.restore(unitId, expectedStorageRevision));
  }

  async updateMetadata(unitId: string, metadata: Partial<WorkspaceRecordMetadata>, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(unitId, () => this.store.updateMetadata(unitId, metadata, expectedStorageRevision));
  }

  async updateUserState(unitId: string, userState: Partial<WorkspaceUserState>, expectedStorageRevision?: number): Promise<WorkspaceRecord> {
    return this.withWorkbookWriter(unitId, () => this.store.updateUserState(unitId, userState, expectedStorageRevision));
  }

  disposeAsync(): Promise<void> {
    return this.coordinator.disposeAsync();
  }

  async purge(unitId: string, cleanup?: { removeSparseSource?: (sourceId: string) => Promise<void> }): Promise<void> {
    const record = await this.store.open(unitId);
    for (const source of record?.snapshot.dataModel.sources ?? []) {
      await this.dataBlocks.removeSource(source.id);
      await this.sparseOverlays.removeSource(source.id);
      await cleanup?.removeSparseSource?.(source.id);
    }
    await this.clear(unitId);
  }

  async clear(unitId: string): Promise<void> {
    this.operationJournal.clear(unitId);
    await this.store.delete(unitId);
  }
}

export interface PersistenceSnapshotMeta {
  unitId: string;
  revision: number;
  checksum: string;
  updatedAt: string;
  hasPendingOperations: boolean;
  pendingOperationCount: number;
  localRevision?: number;
  syncMode?: 'remote' | 'local-only';
}

export function buildPersistenceMeta(
  snapshot: WorkbookSnapshot,
  revision: number,
  pendingOperationCount = 0,
  record?: WorkspaceRecord | null,
): PersistenceSnapshotMeta {
  return {
    unitId: snapshot.unitId,
    revision,
    checksum: computeChecksum(snapshotPayload(snapshot)),
    updatedAt: new Date().toISOString(),
    hasPendingOperations: pendingOperationCount > 0,
    pendingOperationCount,
    localRevision: record?.localRevision,
    syncMode: record?.syncMode,
  };
}
