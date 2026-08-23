import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type { OperationEnvelope } from '@react-sheets/protocol';
import { computeChecksum, verifyChecksum } from './checksum';

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

interface IndexedDbRequest<T> {
  result: T;
  onsuccess: (() => void) | null;
  onerror: (() => void) | null;
  error?: DOMException | null;
}

interface IndexedDbTransaction {
  objectStore(name: string): {
    get(key: string): IndexedDbRequest<WorkspaceRecord | undefined>;
    getAll(): IndexedDbRequest<WorkspaceRecord[]>;
    put(value: WorkspaceRecord): IndexedDbRequest<unknown>;
    delete(key: string): IndexedDbRequest<unknown>;
  };
  oncomplete: (() => void) | null;
  onerror: (() => void) | null;
  onabort: (() => void) | null;
}

interface IndexedDbDatabase {
  objectStoreNames: { contains(name: string): boolean };
  createObjectStore(name: string, options?: { keyPath?: string }): unknown;
  transaction(name: string, mode: 'readonly' | 'readwrite'): IndexedDbTransaction;
  close(): void;
}

interface IndexedDbOpenRequest extends IndexedDbRequest<IndexedDbDatabase> {
  onupgradeneeded: (() => void) | null;
}

interface IndexedDbFactory {
  open(name: string): IndexedDbOpenRequest;
}

export interface IndexedDbWorkspaceStoreOptions {
  databaseName?: string;
  indexedDB?: IndexedDbFactory | null;
}

const WORKSPACE_STORE_NAME = 'workspaces';
const memoryWorkspaceDatabases = new Map<string, Map<string, WorkspaceRecord>>();

function memoryWorkspaceRecords(databaseName: string): Map<string, WorkspaceRecord> {
  let records = memoryWorkspaceDatabases.get(databaseName);
  if (!records) {
    records = new Map<string, WorkspaceRecord>();
    memoryWorkspaceDatabases.set(databaseName, records);
  }
  return records;
}

function resolveIndexedDb(explicit: IndexedDbFactory | null | undefined): IndexedDbFactory | null {
  if (explicit !== undefined) return explicit;
  if (typeof globalThis !== 'undefined' && 'indexedDB' in globalThis) {
    return (globalThis as typeof globalThis & { indexedDB?: IndexedDbFactory }).indexedDB ?? null;
  }
  return null;
}

function requestResult<T>(request: IndexedDbRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'));
  });
}

function openDatabase(factory: IndexedDbFactory, name: string): Promise<IndexedDbDatabase> {
  return new Promise<IndexedDbDatabase>((resolve, reject) => {
    const request = factory.open(name);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(WORKSPACE_STORE_NAME)) {
        request.result.createObjectStore(WORKSPACE_STORE_NAME, { keyPath: 'unitId' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB open failed'));
  });
}

function completeTransaction(transaction: IndexedDbTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(new Error('IndexedDB transaction aborted'));
  });
}

/** IndexedDB-backed WorkspaceRecord storage with a memory-only Node fallback. */
export class IndexedDbWorkspaceStore {
  private readonly databaseName: string;
  private readonly factory: IndexedDbFactory | null;
  private databasePromise: Promise<IndexedDbDatabase> | null = null;

  constructor(options: IndexedDbWorkspaceStoreOptions = {}) {
    this.databaseName = options.databaseName ?? 'react-sheets-workspaces';
    this.factory = resolveIndexedDb(options.indexedDB);
  }

  async load(unitId: string): Promise<WorkspaceRecord | null> {
    if (!this.factory) return clone(memoryWorkspaceRecords(this.databaseName).get(unitId) ?? null);
    const database = await this.database();
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readonly');
    const transactionComplete = completeTransaction(transaction);
    const value = await requestResult(transaction.objectStore(WORKSPACE_STORE_NAME).get(unitId));
    await transactionComplete;
    if (!value) return null;
    if (!verifyWorkspaceRecord(value)) {
      await this.clear(unitId);
      return null;
    }
    return clone(value);
  }

  async save(record: WorkspaceRecord): Promise<void> {
    if (!verifyWorkspaceRecord(record)) throw new Error(`Invalid WorkspaceRecord: ${record.unitId}`);
    if (!this.factory) {
      memoryWorkspaceRecords(this.databaseName).set(record.unitId, clone(record));
      return;
    }
    const database = await this.database();
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readwrite');
    transaction.objectStore(WORKSPACE_STORE_NAME).put(clone(record));
    await completeTransaction(transaction);
  }

  async clear(unitId: string): Promise<void> {
    if (!this.factory) {
      memoryWorkspaceRecords(this.databaseName).delete(unitId);
      return;
    }
    const database = await this.database();
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readwrite');
    transaction.objectStore(WORKSPACE_STORE_NAME).delete(unitId);
    await completeTransaction(transaction);
  }

  async list(): Promise<WorkspaceRecord[]> {
    if (!this.factory) {
      return [...memoryWorkspaceRecords(this.databaseName).values()].map((record) => clone(record));
    }
    const database = await this.database();
    const transaction = database.transaction(WORKSPACE_STORE_NAME, 'readonly');
    const transactionComplete = completeTransaction(transaction);
    const values = await requestResult(transaction.objectStore(WORKSPACE_STORE_NAME).getAll());
    await transactionComplete;
    const valid: WorkspaceRecord[] = [];
    for (const value of values) {
      if (verifyWorkspaceRecord(value)) valid.push(clone(value));
      else await this.clear(value.unitId);
    }
    return valid;
  }

  private database(): Promise<IndexedDbDatabase> {
    if (!this.databasePromise) this.databasePromise = openDatabase(this.factory!, this.databaseName);
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
    await this.indexedDb.save(record);
    return clone(record);
  }

  async save(record: WorkspaceRecord): Promise<WorkspaceRecord> {
    await this.indexedDb.save(record);
    return clone(record);
  }

  async checkpoint(input: WorkspaceRecordInput): Promise<WorkspaceRecord> {
    return this.create(buildWorkspaceRecord(input));
  }

  async list(): Promise<LocalWorkspaceSummary[]> {
    const records = await this.indexedDb.list();
    return records.map(summarizeWorkspace);
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

  constructor(options: IndexedDbWorkspaceStoreOptions = {}, operationJournal = new OperationJournalStore()) {
    this.store = new LocalWorkspaceStore(options);
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

  async checkpoint(
    snapshot: WorkbookSnapshot,
    localRevision: number,
    serverRevision: number,
    syncMode: 'remote' | 'local-only',
    pendingJournal = this.operationJournal.read(snapshot.unitId),
  ): Promise<WorkspaceRecord> {
    const record = buildWorkspaceRecord({
      unitId: snapshot.unitId,
      snapshot,
      localRevision,
      serverRevision,
      syncMode,
      operations: pendingJournal?.operations ?? [],
      nextClientSequence: pendingJournal?.nextClientSequence ?? 0,
    });
    await this.store.save(record);
    return record;
  }

  clear(unitId: string): Promise<void> {
    this.operationJournal.clear(unitId);
    return this.store.delete(unitId);
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
