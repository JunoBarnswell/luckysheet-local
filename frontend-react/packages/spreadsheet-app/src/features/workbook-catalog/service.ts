import { MAX_WORKBOOK_NAME_LENGTH, type WorkbookCatalogQuery as ProtocolWorkbookCatalogQuery, type WorkbookSummary } from '@react-sheets/protocol';
import {
  NativeDocumentTransactionRegistry,
  WorkbookApiNativeDocumentTransport,
  createNativeDocumentTransaction,
} from '../native-document';
import { filterWorkbookCatalog } from './state';
import type {
  WorkbookCatalogCreateInput,
  WorkbookCatalogEntry,
  WorkbookCatalogExportInput,
  WorkbookCatalogExportResult,
  WorkbookCatalogImportInput,
  WorkbookCatalogImportResult,
  WorkbookCatalogPage,
  WorkbookCatalogQuery,
  WorkbookCatalogRequestOptions,
  WorkbookCatalogRemoteClient,
  WorkbookRole,
  WorkbookResolution,
} from './types';
import { WorkbookResolver } from './resolver';

export const DEFAULT_NATIVE_IMPORT_MAX_BYTES = 1024 * 1024 * 1024;

export class WorkbookCatalogError extends Error {
  readonly code: 'not-found' | 'permission-denied' | 'remote-unavailable' | 'invalid-input' | 'conflict';

  constructor(code: WorkbookCatalogError['code'], message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'WorkbookCatalogError';
    this.code = code;
  }
}

export interface WorkbookCatalogServiceOptions {
  remote?: WorkbookCatalogRemoteClient;
  now?: () => Date;
  remoteAvailable?: () => boolean;
  shareTokenProvider?: import('@react-sheets/protocol').ShareTokenProvider;
  nativeDocumentTransactions?: NativeDocumentTransactionRegistry;
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

function normalizeRole(role: WorkbookRole | undefined, fallback: WorkbookRole = 'owner'): WorkbookRole {
  return role === 'owner' || role === 'editor' || role === 'commenter' || role === 'viewer' ? role : fallback;
}

function remoteEntry(summary: WorkbookSummary): WorkbookCatalogEntry {
  return {
    unitId: summary.unitId,
    name: summary.name,
    revision: summary.revision,
    updatedAt: summary.updatedAt,
    storage: 'remote',
    syncState: summary.syncStatus ?? 'synced',
    role: normalizeRole(summary.role, 'viewer'),
    lifecycle: summary.lifecycle ?? (summary.deletedAt ? 'trashed' : 'active'),
    source: summary.source === 'document-import' ? 'document-import' : 'native',
    ownerId: summary.ownerSubject,
    spaceId: summary.spaceId,
    spaceName: summary.spaceName,
    folderId: summary.folderId,
    locationPath: summary.locationPath ?? [],
    sourceFileName: summary.sourceFileName,
    deletedAt: summary.deletedAt,
    favorite: Boolean(summary.favorite),
    lastOpenedAt: summary.lastOpenedAt,
    pendingOperationCount: 0,
  };
}

export class WorkbookCatalogService {
  private readonly nativeTransactions: NativeDocumentTransactionRegistry;
  readonly remote?: WorkbookCatalogRemoteClient;
  readonly resolver: WorkbookResolver;
  private readonly now: () => Date;
  private readonly remoteAvailable?: () => boolean;

  constructor(options: WorkbookCatalogServiceOptions = {}) {
    this.remote = options.remote;
    this.now = options.now ?? (() => new Date());
    this.remoteAvailable = options.remoteAvailable;
    const transport = this.remote ? new WorkbookApiNativeDocumentTransport(this.remote) : undefined;
    this.nativeTransactions = options.nativeDocumentTransactions ?? new NativeDocumentTransactionRegistry(transport);
    this.resolver = new WorkbookResolver({
      remote: this.remote,
      remoteAvailable: () => this.canUseRemote(),
      shareTokenProvider: options.shareTokenProvider,
    });
  }

  private canUseRemote(): boolean {
    return Boolean(this.remote && (this.remoteAvailable ? this.remoteAvailable() : true));
  }

  private requireRemote(): WorkbookCatalogRemoteClient {
    if (!this.remote || !this.canUseRemote()) throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is unavailable');
    return this.remote;
  }

  private async entryFromManifest(
    unitId: string,
    overrides: Partial<Pick<WorkbookCatalogEntry, 'lifecycle' | 'source' | 'spaceId' | 'folderId' | 'sourceFileName'>> = {},
  ): Promise<WorkbookCatalogEntry> {
    const api = this.requireRemote();
    const [manifest, access, state] = await Promise.all([
      api.getManifest(unitId),
      api.getAccess(unitId),
      api.getWorkbookUserState(unitId),
    ]);
    return {
      unitId,
      name: manifest.name,
      revision: manifest.revision,
      updatedAt: this.now().toISOString(),
      storage: 'remote',
      syncState: 'synced',
      role: access.role,
      lifecycle: overrides.lifecycle ?? 'active',
      source: overrides.source ?? 'native',
      spaceId: overrides.spaceId,
      folderId: overrides.folderId,
      locationPath: [],
      sourceFileName: overrides.sourceFileName,
      favorite: Boolean(state.favorite),
      lastOpenedAt: state.lastOpenedAt,
      pendingOperationCount: 0,
    };
  }

  async listPage(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogPage> {
    const protocolQuery: ProtocolWorkbookCatalogQuery = {
      view: query.view,
      query: query.query,
      spaceId: query.spaceId,
      folderId: query.folderId,
      cursor: query.cursor,
      limit: query.limit,
    };
    const page = await this.requireRemote().listWorkbookPage(protocolQuery, options);
    return { entries: filterWorkbookCatalog(page.items.map(remoteEntry), query), nextCursor: page.nextCursor };
  }

  async list(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogEntry[]> {
    return (await this.listPage(query, options)).entries;
  }

  async create(input: WorkbookCatalogCreateInput): Promise<WorkbookCatalogEntry> {
    const plan = input.plan;
    if (!plan.unitId.trim() || !plan.name.trim() || !plan.sheets?.length) {
      throw new WorkbookCatalogError('invalid-input', 'Workbook creation plan requires identity, name, and worksheets');
    }
    if ((input.destination ?? 'remote') !== 'remote') throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is required');
    const opened = await this.requireRemote().createKernelWorkbook({
      ...plan,
      spaceId: input.metadata?.spaceId,
      folderId: input.metadata?.folderId,
      source: input.source ?? 'native',
    });
    return {
      unitId: opened.unitId,
      name: opened.manifest.name,
      revision: opened.revision,
      updatedAt: this.now().toISOString(),
      storage: 'remote',
      syncState: 'synced',
      role: normalizeRole(input.role),
      lifecycle: 'active',
      source: input.source ?? 'native',
      spaceId: input.metadata?.spaceId,
      folderId: input.metadata?.folderId,
      locationPath: [],
      favorite: false,
      pendingOperationCount: 0,
    };
  }

  resolve(unitId: string, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookResolution> {
    return this.resolver.resolve(unitId, options);
  }

  async markOpened(resolution: WorkbookResolution): Promise<WorkbookCatalogEntry> {
    if (resolution.schema !== 'WorkbookResolution' || resolution.lifecycle !== 'active') {
      throw new WorkbookCatalogError('conflict', `Workbook resolution is not openable: ${resolution.unitId}`);
    }
    const remote = this.requireRemote();
    const current = await remote.getWorkbookUserState(resolution.unitId);
    const { unitId: _unitId, ...state } = current;
    const openedAt = this.now().toISOString();
    const saved = await remote.putWorkbookUserState(resolution.unitId, { ...state, lastOpenedAt: openedAt });
    return {
      unitId: resolution.unitId,
      name: resolution.manifest.name,
      revision: resolution.revision,
      updatedAt: openedAt,
      storage: 'remote',
      syncState: 'synced',
      role: resolution.access?.role ?? 'viewer',
      lifecycle: 'active',
      source: 'native',
      locationPath: [],
      favorite: Boolean(saved.favorite),
      lastOpenedAt: saved.lastOpenedAt,
      pendingOperationCount: 0,
    };
  }

  async importWorkbook(input: WorkbookCatalogImportInput): Promise<WorkbookCatalogImportResult> {
    const fileName = input.fileName.trim();
    if (!fileName) throw new WorkbookCatalogError('invalid-input', 'Native document file name is required');
    if (input.buffer.byteLength < 1 || input.buffer.byteLength > DEFAULT_NATIVE_IMPORT_MAX_BYTES) {
      throw new WorkbookCatalogError('invalid-input', `Native document must contain between 1 and ${DEFAULT_NATIVE_IMPORT_MAX_BYTES} bytes`);
    }
    if ((input.destination ?? 'remote') !== 'remote') throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is required');
    const api = this.requireRemote();
    const compatibilityTarget = input.options?.compatibilityTarget ?? (await api.getUserPreferences()).importCompatibility;
    const transport = new WorkbookApiNativeDocumentTransport(api, {
      name: fileName,
      spaceId: input.spaceId,
      folderId: input.folderId,
    });
    const transaction = createNativeDocumentTransaction(transport);
    const imported = await transaction.import({
      fileName,
      content: input.buffer,
      options: { ...input.options, compatibilityTarget },
    });
    if (!imported.manifest) throw new WorkbookCatalogError('conflict', 'Native import omitted its committed manifest');
    const bound = this.nativeTransactions.getOrCreate(imported.manifest.unitId, new WorkbookApiNativeDocumentTransport(api));
    await bound.attach(imported.artifact);
    const entry = await this.entryFromManifest(imported.manifest.unitId, {
      source: 'document-import',
      spaceId: input.spaceId,
      folderId: input.folderId,
      sourceFileName: fileName,
    });
    return { entry, manifest: imported.manifest, report: imported.report, artifact: imported.artifact };
  }

  async exportWorkbook(unitId: string, input: WorkbookCatalogExportInput = {}): Promise<WorkbookCatalogExportResult> {
    const resolved = await this.resolve(unitId);
    const transaction = this.nativeTransactions.getOrCreate(unitId, new WorkbookApiNativeDocumentTransport(this.requireRemote()));
    const fileName = input.fileName ?? transaction.artifact?.fileName ?? `${resolved.manifest.name || 'workbook'}.xlsx`;
    const exported = await transaction.export({
      unitId,
      revision: resolved.revision,
      fileName,
      options: input.options,
      mode: input.fileName ? 'save-as' : 'export',
    });
    if (!exported.content || !exported.fileName) throw new WorkbookCatalogError('conflict', 'Native document export omitted its bytes or file name');
    return { unitId, fileName: exported.fileName, buffer: exported.content, report: exported.report };
  }

  async syncToServer(unitId: string): Promise<WorkbookCatalogSyncResult> {
    const api = this.requireRemote();
    const checkpoint = await api.checkpointWorkbook(unitId);
    const entry = await this.entryFromManifest(unitId);
    if (entry.revision !== checkpoint.workbook.revision) throw new WorkbookCatalogError('conflict', 'Checkpoint revision changed while reading the catalog');
    return { entry, committedOperationCount: 0, revision: entry.revision };
  }

  async rename(unitId: string, name: string): Promise<WorkbookCatalogEntry> {
    const trimmed = name.trim();
    if (!trimmed) throw new WorkbookCatalogError('invalid-input', 'Workbook name is required');
    if (trimmed.length > MAX_WORKBOOK_NAME_LENGTH) throw new WorkbookCatalogError('invalid-input', 'Workbook name is too long');
    return remoteEntry(await this.requireRemote().updateWorkbook(unitId, { name: trimmed }));
  }

  async copy(unitId: string, request: { name?: string; spaceId?: string; folderId?: string; destination?: 'remote' } = {}): Promise<WorkbookCatalogEntry> {
    if ((request.destination ?? 'remote') !== 'remote') throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is required');
    return remoteEntry(await this.requireRemote().copyWorkbook(unitId, {
      name: request.name,
      spaceId: request.spaceId,
      folderId: request.folderId,
    }));
  }

  async move(unitId: string, input: WorkbookCatalogMoveInput): Promise<WorkbookCatalogEntry> {
    return remoteEntry(await this.requireRemote().updateWorkbook(unitId, { spaceId: input.spaceId, folderId: input.folderId }));
  }

  async grantAccess(unitId: string, subject: string, role: WorkbookRole): Promise<void> {
    const normalizedSubject = subject.trim();
    if (!normalizedSubject) throw new WorkbookCatalogError('invalid-input', 'Share subject is required');
    if (role === 'owner') throw new WorkbookCatalogError('invalid-input', 'Owner role cannot be granted through workbook sharing');
    await this.requireRemote().putWorkbookAcl(unitId, normalizedSubject, normalizeRole(role));
  }

  async revokeAccess(unitId: string, subject: string): Promise<void> {
    const normalizedSubject = subject.trim();
    if (!normalizedSubject) throw new WorkbookCatalogError('invalid-input', 'Share subject is required');
    await this.requireRemote().deleteWorkbookAcl(unitId, normalizedSubject);
  }

  async moveToTrash(unitId: string): Promise<WorkbookCatalogEntry> {
    await this.requireRemote().moveToTrash(unitId);
    return this.entryFromManifest(unitId, { lifecycle: 'trashed' });
  }

  async restore(unitId: string): Promise<WorkbookCatalogEntry> {
    return remoteEntry(await this.requireRemote().restoreFromTrash(unitId));
  }

  async purge(unitId: string): Promise<void> {
    await this.requireRemote().purgeWorkbook(unitId);
    this.nativeTransactions.delete(unitId);
  }

  async setFavorite(unitId: string, favorite: boolean): Promise<WorkbookCatalogEntry> {
    const remote = this.requireRemote();
    const current = await remote.getWorkbookUserState(unitId);
    const { unitId: _unitId, ...state } = current;
    await remote.putWorkbookUserState(unitId, { ...state, favorite });
    const entry = await this.entryFromManifest(unitId);
    return { ...entry, favorite };
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
