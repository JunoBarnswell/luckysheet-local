import { migrateStoredWorkbookSnapshot, type WorkbookSnapshot } from '@react-sheets/core-model';
import type { XlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import type { OperationEnvelope } from '@react-sheets/protocol';
import { computeChecksum, verifyChecksum } from './checksum';
import { LocalDataBlockStore } from './data-block-store';
import { buildXlsxArtifactRecord, LocalXlsxArtifactStore } from './xlsx-artifact-store';
import { LocalSparseOverlayStore } from '../data-source/overlay-store';
import {
  openWorkspaceDatabase,
  requestResult,
  transactionComplete,
  WORKSPACE_DATABASE_NAME,
  WORKSPACE_STORE_NAME,
  XLSX_ARTIFACT_STORE_NAME,
  type IndexedDbStoreOptions,
} from './indexed-db';

/** Public options retained for all browser and runtime persistence callers. */
export type IndexedDbWorkspaceStoreOptions = IndexedDbStoreOptions;

export type WorkspaceLifecycle = 'active' | 'trashed';
export type WorkspaceStorageLocation = 'local' | 'remote' | 'mirrored';
export type WorkspaceSource = 'native' | 'xlsx-import';
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
  return {
    ...clone(record),
    snapshot: migrateStoredWorkbookSnapshot(record.snapshot),
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

const memoryWorkspaceDatabases = new Map<string, Map<string, WorkspaceRecord>>();

function memoryWorkspaceRecords(databaseName: string): Map<string, WorkspaceRecord> {
  let records = memoryWorkspaceDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, WorkspaceRecord>();
    memoryWorkspaceDatabases.set(databaseName, records);
  }
  return records;
}

/** IndexedDB-backed WorkspaceRecord storage with a memory-only Node fallback. */
export class IndexedDbWorkspaceStore {
  private readonly options: IndexedDbWorkspaceStoreOptions;
  private readonly databaseName: string;
  private databasePromise: Promise<IDBDatabase | null> | null = null;

  constructor(options: IndexedDbWorkspaceStoreOptions = {}) {
    this.databaseName = options.databaseName ?? WORKSPACE_DATABASE_NAME;
    this.options = options;
  }

  async load(unitId: string): Promise<WorkspaceRecord | null> {
    const database = await this.database();
    if (!database) {
      const value = memoryWorkspaceRecords(this.databaseName).get(unitId);
      return value ? normalizeWorkspaceRecord(value) : null;
    }
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readonly');
    const complete = transactionComplete(transaction);
    const value = await requestResult(transaction.objectStore(WORKSPACE_STORE_NAME).get(unitId)) as WorkspaceRecord | undefined;
    await complete;
    if (!value) return null;
    if (!verifyWorkspaceRecord(value)) {
      await this.clear(unitId);
      return null;
    }
    return normalizeWorkspaceRecord(value);
  }

  async save(record: WorkspaceRecord): Promise<void> {
    if (!verifyWorkspaceRecord(record)) throw new Error(`Invalid WorkspaceRecord: ${record.unitId}`);
    const normalized = normalizeWorkspaceRecord(record);
    const database = await this.database();
    if (!database) {
      memoryWorkspaceRecords(this.databaseName).set(record.unitId, clone(normalized));
      return;
    }
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readwrite');
    transaction.objectStore(WORKSPACE_STORE_NAME).put(clone(normalized));
    await transactionComplete(transaction);
  }

  async clear(unitId: string): Promise<void> {
    const database = await this.database();
    if (!database) {
      memoryWorkspaceRecords(this.databaseName).delete(unitId);
      return;
    }
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readwrite');
    transaction.objectStore(WORKSPACE_STORE_NAME).delete(unitId);
    await transactionComplete(transaction);
  }

  async list(): Promise<WorkspaceRecord[]> {
    const database = await this.database();
    if (!database) {
      return [...memoryWorkspaceRecords(this.databaseName).values()].map((record) => normalizeWorkspaceRecord(record));
    }
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readonly');
    const complete = transactionComplete(transaction);
    const values = await requestResult(transaction.objectStore(WORKSPACE_STORE_NAME).getAll()) as WorkspaceRecord[];
    await complete;
    const valid: WorkspaceRecord[] = [];
    for (const value of values) {
      if (verifyWorkspaceRecord(value)) valid.push(normalizeWorkspaceRecord(value));
      else await this.clear(value.unitId);
    }
    return valid;
  }

  private database(): Promise<IDBDatabase | null> {
    if (!this.databasePromise) this.databasePromise = openWorkspaceDatabase(this.options);
    return this.databasePromise;
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
  private readonly indexedDb: IndexedDbWorkspaceStore;

  constructor(options: IndexedDbWorkspaceStoreOptions = {}) {
    this.indexedDb = new IndexedDbWorkspaceStore(options);
  }

  open(unitId: string): Promise<WorkspaceRecord | null> {
    return this.indexedDb.load(unitId);
  }

  async create(input: WorkspaceRecordInput | WorkspaceRecord): Promise<WorkspaceRecord> {
    const record = 'pending' in input
      ? clone(input as WorkspaceRecord)
      : buildWorkspaceRecord(input as WorkspaceRecordInput);
    const normalized = normalizeWorkspaceRecord(record);
    await this.indexedDb.save(normalized);
    return clone(normalized);
  }

  async save(record: WorkspaceRecord): Promise<WorkspaceRecord> {
    const normalized = normalizeWorkspaceRecord(record);
    await this.indexedDb.save(normalized);
    return clone(normalized);
  }

  async checkpoint(input: WorkspaceRecordInput): Promise<WorkspaceRecord> {
    return this.create(buildWorkspaceRecord(input));
  }

  async list(): Promise<LocalWorkspaceSummary[]> {
    const records = await this.indexedDb.list();
    return records.map(summarizeWorkspace);
  }

  listRecords(): Promise<WorkspaceRecord[]> {
    return this.indexedDb.list();
  }

  async updateMetadata(unitId: string, metadata: Partial<WorkspaceRecordMetadata>): Promise<WorkspaceRecord> {
    const record = await this.open(unitId);
    if (!record) throw new Error(`Unknown local workbook: ${unitId}`);
    const updated = normalizeWorkspaceRecord({
      ...record,
      metadata: { ...record.metadata, ...metadata },
      updatedAt: new Date().toISOString(),
    });
    await this.save(updated);
    return updated;
  }

  async updateUserState(unitId: string, userState: Partial<WorkspaceUserState>): Promise<WorkspaceRecord> {
    const record = await this.open(unitId);
    if (!record) throw new Error(`Unknown local workbook: ${unitId}`);
    const updated = normalizeWorkspaceRecord({
      ...record,
      userState: { ...record.userState, ...userState },
      updatedAt: new Date().toISOString(),
    });
    await this.save(updated);
    return updated;
  }

  async moveToTrash(unitId: string, deletedAt = new Date().toISOString()): Promise<WorkspaceRecord> {
    return this.updateMetadata(unitId, { lifecycle: 'trashed', deletedAt });
  }

  async restore(unitId: string): Promise<WorkspaceRecord> {
    return this.updateMetadata(unitId, { lifecycle: 'active', deletedAt: undefined });
  }

  delete(unitId: string): Promise<void> {
    return this.indexedDb.clear(unitId);
  }
}

let defaultLocalWorkspaceStore: LocalWorkspaceStore | null = null;

export function getLocalWorkspaceStore(options?: IndexedDbWorkspaceStoreOptions): LocalWorkspaceStore {
  if (!defaultLocalWorkspaceStore || options) {
    defaultLocalWorkspaceStore = new LocalWorkspaceStore(options);
  }
  return defaultLocalWorkspaceStore;
}

export class WorkspacePersistence {
  readonly operationJournal: OperationJournalStore;
  readonly store: LocalWorkspaceStore;
  readonly dataBlocks: LocalDataBlockStore;
  readonly sparseOverlays: LocalSparseOverlayStore;
  readonly xlsxArtifacts: LocalXlsxArtifactStore;
  private readonly options: IndexedDbWorkspaceStoreOptions;

  constructor(options: IndexedDbWorkspaceStoreOptions = {}, operationJournal = new OperationJournalStore()) {
    this.options = options;
    this.store = new LocalWorkspaceStore(options);
    this.dataBlocks = new LocalDataBlockStore(options);
    this.sparseOverlays = new LocalSparseOverlayStore(options);
    this.xlsxArtifacts = new LocalXlsxArtifactStore(options);
    this.operationJournal = operationJournal;
  }

  async load(unitId: string): Promise<WorkspaceRecord | null> {
    const record = await this.store.open(unitId);
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
    const previous = await this.store.open(snapshot.unitId);
    const record = buildWorkspaceRecord({
      unitId: snapshot.unitId,
      snapshot,
      localRevision,
      serverRevision,
      syncMode,
      operations: pendingJournal?.operations ?? [],
      nextClientSequence: pendingJournal?.nextClientSequence ?? 0,
      metadata: { ...(previous?.metadata ?? {}), ...(metadata ?? {}) },
      userState: { ...(previous?.userState ?? {}), ...(userState ?? {}) },
    });
    await this.store.save(record);
    return record;
  }

  /**
   * Commits the canonical workspace checkpoint and its source XLSX artifact
   * in one IndexedDB transaction.  The memory-only runtime keeps the same
   * public operation and applies both records through their normal stores.
   */
  async checkpointWithArtifact(
    snapshot: WorkbookSnapshot,
    localRevision: number,
    serverRevision: number,
    syncMode: 'remote' | 'local-only',
    artifact: XlsxSourceArtifact,
    pendingJournal = this.operationJournal.read(snapshot.unitId),
    metadata?: Partial<WorkspaceRecordMetadata>,
    userState?: Partial<WorkspaceUserState>,
  ): Promise<WorkspaceRecord> {
    const previous = await this.store.open(snapshot.unitId);
    const record = buildWorkspaceRecord({
      unitId: snapshot.unitId,
      snapshot,
      localRevision,
      serverRevision,
      syncMode,
      operations: pendingJournal?.operations ?? [],
      nextClientSequence: pendingJournal?.nextClientSequence ?? 0,
      metadata: { ...(previous?.metadata ?? {}), ...(metadata ?? {}) },
      userState: { ...(previous?.userState ?? {}), ...(userState ?? {}) },
    });
    const artifactRecord = await buildXlsxArtifactRecord(snapshot.unitId, artifact);
    const database = await openWorkspaceDatabase(this.options);
    if (!database) {
      await this.store.save(record);
      await this.xlsxArtifacts.save(snapshot.unitId, artifact);
      return record;
    }
    const transaction = database.transaction(
      [WORKSPACE_STORE_NAME, XLSX_ARTIFACT_STORE_NAME],
      'readwrite',
    );
    transaction.objectStore(WORKSPACE_STORE_NAME).put(clone(record));
    transaction.objectStore(XLSX_ARTIFACT_STORE_NAME).put(artifactRecord);
    await transactionComplete(transaction);
    return record;
  }

  async moveToTrash(unitId: string, deletedAt = new Date().toISOString()): Promise<WorkspaceRecord> {
    return this.store.moveToTrash(unitId, deletedAt);
  }

  async restore(unitId: string): Promise<WorkspaceRecord> {
    return this.store.restore(unitId);
  }

  async updateMetadata(unitId: string, metadata: Partial<WorkspaceRecordMetadata>): Promise<WorkspaceRecord> {
    return this.store.updateMetadata(unitId, metadata);
  }

  async updateUserState(unitId: string, userState: Partial<WorkspaceUserState>): Promise<WorkspaceRecord> {
    return this.store.updateUserState(unitId, userState);
  }

  async purge(unitId: string, cleanup?: { removeSparseSource?: (sourceId: string) => Promise<void> }): Promise<void> {
    const record = await this.store.open(unitId);
    for (const source of record?.snapshot.dataSources ?? []) {
      await this.dataBlocks.removeSource(source.id);
      await this.sparseOverlays.removeSource(source.id);
      await cleanup?.removeSparseSource?.(source.id);
    }
    await this.clear(unitId);
  }

  async clear(unitId: string): Promise<void> {
    this.operationJournal.clear(unitId);
    await Promise.all([this.store.delete(unitId), this.xlsxArtifacts.remove(unitId)]);
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
