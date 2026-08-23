import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type { XlsxSourceArtifact } from '@react-sheets/exchange-xlsx';
import {
  ApiRequestError,
  AuthenticationRequiredError,
  MAX_WORKBOOK_NAME_LENGTH,
  type OperationEnvelope,
  type SnapshotResponse,
  type WorkbookCatalogQuery as ProtocolWorkbookCatalogQuery,
  type WorkbookCreateMetadata,
  type WorkbookSummary,
} from '@react-sheets/protocol';
import { buildOperation } from '../../collaboration/helpers';
import {
  exchangeExportXlsx,
  exchangeImportXlsx,
} from '../xlsx';
import {
  type WorkspaceRecord,
  type WorkspaceRecordMetadata,
  type WorkspaceRole,
  type WorkspaceUserState,
  WorkspacePersistence,
} from '../persistence/storage';
import { filterWorkbookCatalog, resolveWorkbookSyncState } from './state';
import { createWorkbookUnitId } from './templates';
import type {
  WorkbookCatalogCreateInput,
  WorkbookCatalogEntry,
  WorkbookCatalogExportInput,
  WorkbookCatalogExportResult,
  WorkbookCatalogImportInput,
  WorkbookCatalogImportResult,
  WorkbookCatalogPage,
  WorkbookCatalogOpenResult,
  WorkbookCatalogQuery,
  WorkbookCatalogRequestOptions,
  WorkbookCatalogRemoteClient,
  WorkbookRole,
} from './types';

export const DEFAULT_XLSX_IMPORT_MAX_BYTES = 50 * 1024 * 1024;

export class WorkbookCatalogError extends Error {
  readonly code: 'not-found' | 'permission-denied' | 'remote-unavailable' | 'invalid-input' | 'conflict';

  constructor(code: WorkbookCatalogError['code'], message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'WorkbookCatalogError';
    this.code = code;
  }
}

export interface WorkbookCatalogServiceOptions {
  persistence?: WorkspacePersistence;
  remote?: WorkbookCatalogRemoteClient;
  now?: () => Date;
  unitIdFactory?: () => string;
  remoteAvailable?: () => boolean;
}

export interface WorkbookCatalogMoveInput {
  spaceId?: string | null;
  folderId?: string | null;
}

export interface WorkbookCatalogSyncResult {
  entry: WorkbookCatalogEntry;
  committedOperationCount: number;
  revision: number;
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object';
}

function isRemoteUnavailable(error: unknown): boolean {
  if (error instanceof AuthenticationRequiredError) return false;
  if (error instanceof ApiRequestError) return error.status === 408 || error.status === 429 || error.status >= 500;
  if (isRecord(error) && typeof error.status === 'number') return error.status === 0 || error.status >= 500;
  return error instanceof TypeError;
}

function normalizeRole(role: WorkbookRole | undefined, fallback: WorkspaceRole = 'owner'): WorkspaceRole {
  return role === 'owner' || role === 'editor' || role === 'commenter' || role === 'viewer' ? role : fallback;
}

function sourceFromProtocol(summary: WorkbookSummary): WorkspaceRecordMetadata['source'] {
  return summary.source === 'xlsx-import' ? 'xlsx-import' : 'native';
}

function metadataFromRemote(summary: WorkbookSummary): WorkspaceRecordMetadata {
  return {
    location: summary.storageLocation ?? 'remote',
    lifecycle: summary.lifecycle ?? (summary.deletedAt ? 'trashed' : 'active'),
    source: sourceFromProtocol(summary),
    // An omitted server role is fail-closed. The backend contract should
    // always return the actor's effective role in the catalog row.
    role: normalizeRole(summary.role, 'viewer'),
    ownerId: summary.ownerSubject,
    sourceFileName: summary.sourceFileName,
    spaceId: summary.spaceId,
    folderId: summary.folderId,
    locationPath: summary.locationPath?.join(' / '),
    deletedAt: summary.deletedAt,
  };
}

function metadataToProtocol(metadata: Partial<WorkspaceRecordMetadata> | undefined): WorkbookCreateMetadata {
  return {
    spaceId: metadata?.spaceId,
    folderId: metadata?.folderId,
    source: metadata?.source === 'xlsx-import' ? 'xlsx-import' : 'native',
  };
}

function userStateToRemote(state: WorkspaceUserState): Omit<import('@react-sheets/protocol').WorkbookUserState, 'unitId'> {
  return {
    favorite: state.favorite,
    lastOpenedAt: state.lastOpenedAt,
  };
}

function userStateFromRemote(input: import('@react-sheets/protocol').WorkbookUserState): Partial<WorkspaceUserState> {
  return {
    favorite: input.favorite,
    lastOpenedAt: input.lastOpenedAt,
  };
}

function localEntry(record: WorkspaceRecord, remoteAvailable: boolean): WorkbookCatalogEntry {
  const metadata = record.metadata;
  return {
    unitId: record.unitId,
    name: record.snapshot.name,
    revision: record.serverRevision,
    updatedAt: record.updatedAt,
    storage: metadata.location,
    syncState: resolveWorkbookSyncState({
      storage: metadata.location,
      pendingOperationCount: record.pending.operations.length,
      syncMode: record.syncMode,
      remoteAvailable,
    }),
    role: metadata.role,
    lifecycle: metadata.lifecycle,
    source: metadata.source,
    ownerId: metadata.ownerId,
    spaceId: metadata.spaceId,
    folderId: metadata.folderId,
    sourceFileName: metadata.sourceFileName,
    locationPath: metadata.locationPath ? metadata.locationPath.split(' / ') : [],
    deletedAt: metadata.deletedAt,
    favorite: record.userState.favorite,
    lastOpenedAt: record.userState.lastOpenedAt,
    pendingOperationCount: record.pending.operations.length,
    localRecord: record,
  };
}

function remoteEntry(summary: WorkbookSummary): WorkbookCatalogEntry {
  const metadata = metadataFromRemote(summary);
  return {
    unitId: summary.unitId,
    name: summary.name,
    revision: summary.revision,
    updatedAt: summary.updatedAt,
    storage: summary.storageLocation ?? 'remote',
    syncState: summary.syncStatus ?? 'synced',
    role: metadata.role,
    lifecycle: metadata.lifecycle,
    source: metadata.source,
    ownerId: metadata.ownerId,
    ownerName: undefined,
    spaceId: metadata.spaceId,
    spaceName: summary.spaceName,
    folderId: metadata.folderId,
    locationPath: summary.locationPath ?? [],
    sourceFileName: summary.sourceFileName,
    deletedAt: summary.deletedAt,
    favorite: Boolean(summary.favorite),
    lastOpenedAt: summary.lastOpenedAt,
    pendingOperationCount: 0,
  };
}

function mergeEntries(local: WorkbookCatalogEntry, remote: WorkbookCatalogEntry): WorkbookCatalogEntry {
  return {
    ...remote,
    storage: 'mirrored',
    syncState: local.pendingOperationCount > 0 ? 'pending' : remote.syncState,
    favorite: local.favorite || remote.favorite,
    lastOpenedAt: local.lastOpenedAt ?? remote.lastOpenedAt,
    pendingOperationCount: local.pendingOperationCount,
    localRecord: local.localRecord,
  };
}

function assertSnapshotUnitId(snapshot: WorkbookSnapshot, unitId: string): WorkbookSnapshot {
  if (snapshot.unitId !== unitId) throw new WorkbookCatalogError('conflict', `Workbook snapshot unitId mismatch: ${unitId}`);
  return snapshot;
}

function assertRole(record: WorkspaceRecord, action: string, allowed: readonly WorkspaceRole[]): void {
  if (!allowed.includes(record.metadata.role)) {
    throw new WorkbookCatalogError('permission-denied', `Role ${record.metadata.role} cannot ${action} workbook ${record.unitId}`);
  }
}

function renameSnapshot(snapshot: WorkbookSnapshot, name: string): WorkbookSnapshot {
  const next = clone(snapshot);
  next.name = name.trim();
  if (!next.name) throw new WorkbookCatalogError('invalid-input', 'Workbook name is required');
  if (next.name.length > MAX_WORKBOOK_NAME_LENGTH) throw new WorkbookCatalogError('invalid-input', 'Workbook name is too long');
  return next;
}

function newOperationId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return crypto.randomUUID();
  return `operation-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function reidentifySnapshot(snapshot: WorkbookSnapshot, unitId: string): WorkbookSnapshot {
  const next = clone(snapshot);
  next.unitId = unitId;
  next.printDocuments = (next.printDocuments ?? []).map((document) => ({ ...document, unitId }));
  return next;
}

export class WorkbookCatalogService {
  readonly persistence: WorkspacePersistence;
  readonly remote?: WorkbookCatalogRemoteClient;
  private readonly now: () => Date;
  private readonly unitIdFactory: () => string;
  private readonly remoteAvailable?: () => boolean;

  constructor(options: WorkbookCatalogServiceOptions = {}) {
    this.persistence = options.persistence ?? new WorkspacePersistence();
    this.remote = options.remote;
    this.now = options.now ?? (() => new Date());
    this.unitIdFactory = options.unitIdFactory ?? (() => createWorkbookUnitId());
    this.remoteAvailable = options.remoteAvailable;
  }

  private canUseRemote(): boolean {
    return Boolean(this.remote && (this.remoteAvailable ? this.remoteAvailable() : true));
  }

  private requireRemote(): WorkbookCatalogRemoteClient {
    if (!this.remote || !this.canUseRemote()) throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is unavailable');
    return this.remote;
  }

  async listPage(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogPage> {
    const records = await this.persistence.listRecords();
    const local = records.map((record) => localEntry(record, this.canUseRemote()));
    const byId = new Map(local.map((entry) => [entry.unitId, entry]));
    let nextCursor: string | null = null;
    const shouldQueryRemote = this.canUseRemote() && query.view !== 'local';
    if (shouldQueryRemote) {
      try {
        const protocolQuery: ProtocolWorkbookCatalogQuery = {
          view: query.view === 'local' ? 'all' : query.view,
          query: query.query,
          spaceId: query.spaceId,
          folderId: query.folderId,
          cursor: query.cursor,
          limit: query.limit,
        };
        const page = await this.requireRemote().listWorkbookPage(protocolQuery, options);
        nextCursor = page.nextCursor;
        for (const summary of page.items) {
          const remote = remoteEntry(summary);
          const existing = byId.get(remote.unitId);
          byId.set(remote.unitId, existing ? mergeEntries(existing, remote) : remote);
        }
      } catch (error) {
        if (options.signal?.aborted || (isRecord(error) && error.name === 'AbortError')) throw error;
        if (!isRemoteUnavailable(error)) throw error;
        for (const entry of byId.values()) {
          if (entry.storage !== 'remote') entry.syncState = 'offline';
        }
      }
    }
    return { entries: filterWorkbookCatalog([...byId.values()], query), nextCursor };
  }

  async list(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogEntry[]> {
    const page = await this.listPage(query, options);
    return page.entries;
  }

  async create(input: WorkbookCatalogCreateInput): Promise<WorkbookCatalogEntry> {
    const snapshot = assertSnapshotUnitId(input.snapshot, input.snapshot.unitId);
    const destination = input.destination ?? (this.remote ? 'remote' : 'local');
    const metadata = input.metadata ?? {};
    if (destination === 'remote') {
      const api = this.requireRemote();
      const response = await api.createWorkbook(snapshot, metadata);
      const remoteSnapshot = assertSnapshotUnitId(response.snapshot, snapshot.unitId);
      const record = await this.persistence.checkpoint(
        remoteSnapshot,
        1,
        response.revision,
        'remote',
        undefined,
        {
          location: 'remote',
          lifecycle: 'active',
          source: input.source ?? 'native',
          role: normalizeRole(input.role),
          spaceId: metadata.spaceId,
          folderId: metadata.folderId,
          ownerId: undefined,
          sourceFileName: undefined,
        },
      );
      return localEntry(record, true);
    }
    const record = await this.persistence.checkpoint(snapshot, 1, 0, 'local-only', undefined, {
      location: 'local',
      lifecycle: 'active',
      source: input.source ?? 'native',
      role: normalizeRole(input.role),
      spaceId: metadata.spaceId,
      folderId: metadata.folderId,
    });
    return localEntry(record, false);
  }

  async open(unitId: string): Promise<WorkbookCatalogOpenResult> {
    let record = await this.persistence.load(unitId);
    if (record?.metadata.lifecycle === 'trashed') throw new WorkbookCatalogError('not-found', `Workbook is in trash: ${unitId}`);
    if (!record && this.canUseRemote()) {
      const response = await this.requireRemote().getSnapshot(unitId);
      const access = await this.requireRemote().getAccess(unitId);
      record = await this.persistence.checkpoint(response.snapshot, 0, response.revision, 'remote', undefined, {
        location: 'remote', lifecycle: 'active', source: 'native', role: normalizeRole(access.role),
      });
    }
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    const openedAt = this.now().toISOString();
    record = await this.persistence.updateUserState(unitId, { lastOpenedAt: openedAt });
    if (this.canUseRemote()) {
      try {
        const remoteState = await this.requireRemote().getWorkbookUserState(unitId);
        record = await this.persistence.updateUserState(unitId, userStateFromRemote(remoteState));
        record = await this.persistence.updateUserState(unitId, { lastOpenedAt: openedAt });
        await this.requireRemote().putWorkbookUserState(unitId, userStateToRemote(record.userState));
      } catch (error) {
        if (!isRemoteUnavailable(error)
          && !(error instanceof ApiRequestError && [401, 403, 404].includes(error.status))) throw error;
      }
    }
    return { entry: localEntry(record, this.canUseRemote()), snapshot: clone(record.snapshot) };
  }

  async importXlsx(input: WorkbookCatalogImportInput): Promise<WorkbookCatalogImportResult> {
    if (!input.fileName.trim()) throw new WorkbookCatalogError('invalid-input', 'XLSX file name is required');
    if (input.buffer.byteLength > DEFAULT_XLSX_IMPORT_MAX_BYTES) {
      throw new WorkbookCatalogError('invalid-input', `XLSX file exceeds ${DEFAULT_XLSX_IMPORT_MAX_BYTES} byte limit`);
    }
    const imported = await exchangeImportXlsx({
      fileName: input.fileName,
      buffer: input.buffer,
      options: input.options,
      workerPort: input.workerPort,
      execution: input.execution,
    });
    if (!imported.snapshot) throw new WorkbookCatalogError('invalid-input', 'XLSX import did not produce a workbook snapshot');
    const unitId = this.unitIdFactory();
    const importedSnapshot = reidentifySnapshot(imported.snapshot, unitId);
    const destination = input.destination ?? (this.remote ? 'remote' : 'local');
    const metadata: WorkbookCreateMetadata = {
      source: 'xlsx-import',
      spaceId: input.spaceId,
      folderId: input.folderId,
    };
    let record: WorkspaceRecord;
    if (destination === 'remote') {
      const api = this.requireRemote();
      const response = await api.createWorkbookImport({
        artifact: new Blob([input.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        artifactFileName: input.fileName,
        snapshot: importedSnapshot,
        ...metadata,
      });
      const serverSnapshot = assertSnapshotUnitId(response.snapshot, importedSnapshot.unitId);
      record = imported.sourceArtifact
        ? await this.persistence.checkpointWithArtifact(serverSnapshot, 1, response.summary.revision, 'remote', imported.sourceArtifact, undefined, {
          location: 'remote', lifecycle: 'active', source: 'xlsx-import', role: 'owner', spaceId: input.spaceId, folderId: input.folderId,
          sourceFileName: input.fileName,
        })
        : await this.persistence.checkpoint(serverSnapshot, 1, response.summary.revision, 'remote', undefined, {
          location: 'remote', lifecycle: 'active', source: 'xlsx-import', role: 'owner', spaceId: input.spaceId, folderId: input.folderId,
          sourceFileName: input.fileName,
        });
    } else {
      record = imported.sourceArtifact
        ? await this.persistence.checkpointWithArtifact(importedSnapshot, 1, 0, 'local-only', imported.sourceArtifact, undefined, {
          location: 'local', lifecycle: 'active', source: 'xlsx-import', role: 'owner',
          sourceFileName: input.fileName,
        })
        : await this.persistence.checkpoint(importedSnapshot, 1, 0, 'local-only', undefined, {
          location: 'local', lifecycle: 'active', source: 'xlsx-import', role: 'owner',
          sourceFileName: input.fileName,
        });
    }
    return {
      entry: localEntry(record, destination === 'remote'),
      snapshot: clone(record.snapshot),
      report: imported.report,
      sourceArtifact: imported.sourceArtifact,
    };
  }

  async exportXlsx(unitId: string, input: WorkbookCatalogExportInput = {}): Promise<WorkbookCatalogExportResult> {
    const opened = await this.open(unitId);
    let artifact = await this.persistence.xlsxArtifacts.load(unitId);
    if (!artifact && this.canUseRemote()) {
      try {
        const remoteArtifact = await this.requireRemote().getWorkbookSourceArtifact(unitId);
        const imported = await exchangeImportXlsx({
          fileName: remoteArtifact.metadata.fileName,
          buffer: await remoteArtifact.artifact.arrayBuffer(),
          execution: 'inline-test',
        });
        artifact = imported.sourceArtifact ?? null;
        if (artifact) await this.persistence.xlsxArtifacts.save(unitId, artifact);
      } catch (error) {
        if (!isRemoteUnavailable(error)) throw error;
      }
    }
    const exported = await exchangeExportXlsx(opened.snapshot, {
      ...input,
      fileName: input.fileName ?? `${opened.snapshot.name || 'workbook'}.xlsx`,
      sourceArtifact: artifact ?? undefined,
    });
    if (!exported.buffer || !exported.fileName) throw new WorkbookCatalogError('invalid-input', 'XLSX export did not produce a file');
    return { unitId, fileName: exported.fileName, buffer: exported.buffer, report: exported.report };
  }

  async syncToServer(unitId: string): Promise<WorkbookCatalogSyncResult> {
    const api = this.requireRemote();
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    let revision = record.serverRevision;
    let checkpoint: SnapshotResponse;
    if (record.syncMode === 'local-only' || record.metadata.location === 'local') {
      checkpoint = await api.createWorkbook(record.snapshot, metadataToProtocol(record.metadata));
      const artifact = await this.persistence.xlsxArtifacts.load(unitId);
      if (artifact) {
        await api.putWorkbookSourceArtifact(
          unitId,
          new Blob([artifact.buffer], { type: 'application/octet-stream' }),
          artifact.fileName,
        );
      }
      revision = checkpoint.revision;
    } else {
      for (const operation of record.pending.operations) {
        const result = await api.commitOperation(unitId, operation as OperationEnvelope);
        revision = Math.max(revision, result.operation.revision);
      }
      checkpoint = (await api.checkpointWorkbook(unitId)).snapshot;
    }
    assertSnapshotUnitId(checkpoint.snapshot, unitId);
    const nextSequence = record.pending.nextClientSequence;
    this.persistence.operationJournal.write(unitId, [], nextSequence);
    const saved = await this.persistence.checkpoint(checkpoint.snapshot, record.localRevision, checkpoint.revision, 'remote', {
      schema: 'PendingOperationJournal', unitId, nextClientSequence: nextSequence, operations: [], checksum: '',
    }, {
      ...record.metadata,
      location: 'remote',
      lifecycle: 'active',
      deletedAt: undefined,
    }, record.userState);
    // checkpoint() recomputes the journal checksum; the temporary journal
    // above only provides the sequence watermark and is never persisted.
    return {
      entry: localEntry(saved, true),
      committedOperationCount: record.pending.operations.length,
      revision: Math.max(revision, checkpoint.revision),
    };
  }

  async rename(unitId: string, name: string): Promise<WorkbookCatalogEntry> {
    const trimmed = name.trim();
    if (!trimmed) throw new WorkbookCatalogError('invalid-input', 'Workbook name is required');
    if (trimmed.length > MAX_WORKBOOK_NAME_LENGTH) throw new WorkbookCatalogError('invalid-input', 'Workbook name is too long');
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(record, 'rename', ['owner', 'editor']);
    const snapshot = renameSnapshot(record.snapshot, trimmed);
    const nextClientSequence = record.pending.nextClientSequence + 1;
    const operation = buildOperation(
      newOperationId(),
      unitId,
      nextClientSequence,
      record.serverRevision,
      [{ id: 'workbook.renamed', sheetId: snapshot.sheets[0]!.id, params: { name: trimmed } }],
    );
    let serverRevision = record.serverRevision;
    const pending = [...record.pending.operations, operation];
    if (this.canUseRemote() && record.pending.operations.length === 0) {
      const committed = await this.requireRemote().commitOperation(unitId, operation);
      serverRevision = committed.operation.revision;
      pending.length = 0;
    }
    this.persistence.operationJournal.write(unitId, pending, nextClientSequence);
    const saved = await this.persistence.checkpoint(
      snapshot,
      record.localRevision + 1,
      serverRevision,
      pending.length > 0 ? 'remote' : record.syncMode,
      undefined,
    );
    return localEntry(saved, this.canUseRemote());
  }

  async copy(unitId: string, request: { name?: string; spaceId?: string; folderId?: string; destination?: 'local' | 'remote' } = {}): Promise<WorkbookCatalogEntry> {
    const source = await this.open(unitId);
    const newUnitId = this.unitIdFactory();
    let snapshot = reidentifySnapshot(renameSnapshot(source.snapshot, request.name ?? `${source.snapshot.name} - 副本`), newUnitId);
    const destination = request.destination ?? (this.canUseRemote() ? 'remote' : 'local');
    let remoteRevision = 0;
    if (destination === 'remote') {
      const api = this.requireRemote();
      const copied = await api.copyWorkbook(unitId, { name: snapshot.name, spaceId: request.spaceId, folderId: request.folderId });
      snapshot = reidentifySnapshot(snapshot, copied.unitId);
      remoteRevision = copied.revision;
    }
    const artifact = await this.persistence.xlsxArtifacts.load(unitId);
    const saved = artifact
      ? await this.persistence.checkpointWithArtifact(snapshot, 1, remoteRevision, destination === 'remote' ? 'remote' : 'local-only', artifact, undefined, {
        location: destination === 'remote' ? 'remote' : 'local', lifecycle: 'active', source: source.entry.source, role: destination === 'remote' ? 'owner' : 'owner', spaceId: request.spaceId, folderId: request.folderId,
      })
      : await this.persistence.checkpoint(snapshot, 1, remoteRevision, destination === 'remote' ? 'remote' : 'local-only', undefined, {
        location: destination === 'remote' ? 'remote' : 'local', lifecycle: 'active', source: source.entry.source, role: 'owner', spaceId: request.spaceId, folderId: request.folderId,
      });
    return localEntry(saved, destination === 'remote');
  }

  async move(unitId: string, input: WorkbookCatalogMoveInput): Promise<WorkbookCatalogEntry> {
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(record, 'move', ['owner', 'editor']);
    if (record.metadata.role === 'editor'
      && input.spaceId !== undefined
      && input.spaceId !== record.metadata.spaceId) {
      throw new WorkbookCatalogError('permission-denied', 'Editors may move a workbook only inside its current space');
    }
    if (this.canUseRemote()) await this.requireRemote().updateWorkbook(unitId, { spaceId: input.spaceId, folderId: input.folderId });
    const saved = await this.persistence.updateMetadata(unitId, { spaceId: input.spaceId ?? undefined, folderId: input.folderId ?? undefined });
    return localEntry(saved, this.canUseRemote());
  }

  async grantAccess(unitId: string, subject: string, role: WorkbookRole): Promise<void> {
    const normalizedSubject = subject.trim();
    if (!normalizedSubject) throw new WorkbookCatalogError('invalid-input', 'Share subject is required');
    if (role === 'owner') throw new WorkbookCatalogError('invalid-input', 'Owner role cannot be granted through workbook sharing');
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(record, 'share', ['owner']);
    await this.requireRemote().putWorkbookAcl(unitId, normalizedSubject, normalizeRole(role));
  }

  async revokeAccess(unitId: string, subject: string): Promise<void> {
    const normalizedSubject = subject.trim();
    if (!normalizedSubject) throw new WorkbookCatalogError('invalid-input', 'Share subject is required');
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(record, 'revoke sharing', ['owner']);
    await this.requireRemote().deleteWorkbookAcl(unitId, normalizedSubject);
  }

  async moveToTrash(unitId: string): Promise<WorkbookCatalogEntry> {
    const current = await this.persistence.load(unitId);
    if (!current) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(current, 'move to trash', ['owner']);
    if (this.canUseRemote()) await this.requireRemote().moveToTrash(unitId);
    const record = await this.persistence.moveToTrash(unitId, this.now().toISOString());
    return localEntry(record, this.canUseRemote());
  }

  async restore(unitId: string): Promise<WorkbookCatalogEntry> {
    const current = await this.persistence.load(unitId);
    if (!current) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(current, 'restore', ['owner']);
    if (this.canUseRemote()) await this.requireRemote().restoreFromTrash(unitId);
    const record = await this.persistence.restore(unitId);
    return localEntry(record, this.canUseRemote());
  }

  async purge(unitId: string): Promise<void> {
    const current = await this.persistence.load(unitId);
    if (!current) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    assertRole(current, 'purge', ['owner']);
    if (this.canUseRemote()) await this.requireRemote().purgeWorkbook(unitId);
    await this.persistence.purge(unitId);
  }

  async setFavorite(unitId: string, favorite: boolean): Promise<WorkbookCatalogEntry> {
    const record = await this.persistence.load(unitId);
    if (!record) throw new WorkbookCatalogError('not-found', `Workbook not found: ${unitId}`);
    if (this.canUseRemote()) {
      await this.requireRemote().putWorkbookUserState(unitId, { ...userStateToRemote({ ...record.userState, favorite }), favorite });
    }
    const saved = await this.persistence.updateUserState(unitId, { favorite });
    return localEntry(saved, this.canUseRemote());
  }

  async listAccess(unitId: string): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['listWorkbookAcl']>>> {
    return this.requireRemote().listWorkbookAcl(unitId);
  }

  async listSpaces(options: WorkbookCatalogRequestOptions = {}): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['listSpaces']>>> {
    return this.requireRemote().listSpaces(options);
  }

  async getUserPreferences(options: WorkbookCatalogRequestOptions = {}): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['getUserPreferences']>>> {
    return this.requireRemote().getUserPreferences(options);
  }

  async putUserPreferences(input: Parameters<WorkbookCatalogRemoteClient['putUserPreferences']>[0]): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['putUserPreferences']>>> {
    return this.requireRemote().putUserPreferences(input);
  }

  async listFolders(spaceId: string, options: WorkbookCatalogRequestOptions = {}): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['listFolders']>>> {
    return this.requireRemote().listFolders(spaceId, options);
  }

  async createSpace(input: Parameters<WorkbookCatalogRemoteClient['createSpace']>[0]): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['createSpace']>>> {
    return this.requireRemote().createSpace(input);
  }

  async createFolder(spaceId: string, input: Parameters<WorkbookCatalogRemoteClient['createFolder']>[1]): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['createFolder']>>> {
    return this.requireRemote().createFolder(spaceId, input);
  }

  async updateFolder(folderId: string, input: Parameters<WorkbookCatalogRemoteClient['updateFolder']>[1]): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['updateFolder']>>> {
    return this.requireRemote().updateFolder(folderId, input);
  }

  async deleteFolder(folderId: string): Promise<void> {
    await this.requireRemote().deleteFolder(folderId);
  }

  async listSpaceMembers(spaceId: string, options: WorkbookCatalogRequestOptions = {}): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['listSpaceMembers']>>> {
    return this.requireRemote().listSpaceMembers(spaceId, options);
  }

  async putSpaceMember(spaceId: string, subject: string, role: WorkbookRole): Promise<Awaited<ReturnType<WorkbookCatalogRemoteClient['putSpaceMember']>>> {
    return this.requireRemote().putSpaceMember(spaceId, subject.trim(), normalizeRole(role));
  }

  async deleteSpaceMember(spaceId: string, subject: string): Promise<void> {
    await this.requireRemote().deleteSpaceMember(spaceId, subject.trim());
  }
}
