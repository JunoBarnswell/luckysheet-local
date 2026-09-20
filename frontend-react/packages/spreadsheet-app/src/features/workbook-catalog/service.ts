import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type { NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
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
  exchangeImportDocument,
  exchangeSaveAsDocument,
  exchangeSaveDocument,
} from '../native-document';
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
  WorkbookCatalogQuery,
  WorkbookCatalogRequestOptions,
  WorkbookCatalogRemoteClient,
  WorkbookRole,
  WorkbookResolution,
} from './types';
import { WorkbookResolver } from './resolver';

export const DEFAULT_NATIVE_IMPORT_MAX_BYTES = 50 * 1024 * 1024;

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
  shareTokenProvider?: import('@react-sheets/protocol').ShareTokenProvider;
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

function normalizeRole(role: WorkbookRole | undefined, fallback: WorkspaceRole = 'owner'): WorkspaceRole {
  return role === 'owner' || role === 'editor' || role === 'commenter' || role === 'viewer' ? role : fallback;
}

function sourceFromProtocol(summary: WorkbookSummary): WorkspaceRecordMetadata['source'] {
  return summary.source === 'document-import' ? 'document-import' : 'native';
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

function reidentifySnapshot(snapshot: WorkbookSnapshot, unitId: string): WorkbookSnapshot {
  return {
    ...snapshot,
    unitId,
    printDocuments: (snapshot.printDocuments ?? []).map((document) => ({ ...document, unitId })),
  };
}

export class WorkbookCatalogService {
  readonly persistence: WorkspacePersistence;
  readonly remote?: WorkbookCatalogRemoteClient;
  readonly resolver: WorkbookResolver;
  private readonly now: () => Date;
  private readonly unitIdFactory: () => string;
  private readonly remoteAvailable?: () => boolean;

  constructor(options: WorkbookCatalogServiceOptions = {}) {
    this.persistence = options.persistence ?? new WorkspacePersistence();
    this.remote = options.remote;
    this.now = options.now ?? (() => new Date());
    this.unitIdFactory = options.unitIdFactory ?? (() => createWorkbookUnitId());
    this.remoteAvailable = options.remoteAvailable;
    this.resolver = new WorkbookResolver({
      persistence: this.persistence,
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

  async listPage(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogPage> {
    if (query.view === 'local') throw new WorkbookCatalogError('invalid-input', '页面内存文件已移除，请使用服务器文件中心');
    const page = await this.requireRemote().listWorkbookPage(query as ProtocolWorkbookCatalogQuery, options);
    return { entries: page.items.map(remoteEntry), nextCursor: page.nextCursor };
  }

  async list(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookCatalogEntry[]> {
    const page = await this.listPage(query, options);
    return page.entries;
  }

  async create(input: WorkbookCatalogCreateInput): Promise<WorkbookCatalogEntry> {
    if (input.destination === 'local') throw new WorkbookCatalogError('invalid-input', '工作簿必须保存到服务器');
    const response = await this.requireRemote().createWorkbook(input.snapshot, input.metadata);
    return remoteEntry(await this.requireRemote().getWorkbookSummary(response.snapshot.unitId));
  }

  resolve(unitId: string, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookResolution> {
    return this.resolver.resolve(unitId, options);
  }

  async markOpened(resolution: WorkbookResolution): Promise<WorkbookCatalogEntry> {
    const api = this.requireRemote();
    const state = await api.getWorkbookUserState(resolution.unitId);
    await api.putWorkbookUserState(resolution.unitId, { favorite: state.favorite, lastOpenedAt: this.now().toISOString() });
    return remoteEntry(await api.getWorkbookSummary(resolution.unitId));
  }

  async importWorkbook(input: WorkbookCatalogImportInput): Promise<WorkbookCatalogImportResult> {
    if (input.destination === 'local') throw new WorkbookCatalogError('invalid-input', '导入文件必须保存到服务器');
    if (!input.fileName.trim() || input.buffer.byteLength > DEFAULT_NATIVE_IMPORT_MAX_BYTES) throw new WorkbookCatalogError('invalid-input', '文件名无效或超过导入大小限制');
    const api = this.requireRemote();
    const compatibilityTarget = input.options?.compatibilityTarget ?? (await api.getUserPreferences()).importCompatibility;
    const imported = await exchangeImportDocument({ ...input, options: { ...input.options, compatibilityTarget } });
    if (!imported.snapshot) throw new WorkbookCatalogError('invalid-input', '导入未生成工作簿');
    const response = await api.createWorkbookImport({
      artifact: new Blob([input.buffer], { type: nativeDocumentMimeType(input.fileName) }), artifactFileName: input.fileName,
      snapshot: reidentifySnapshot(imported.snapshot, this.unitIdFactory()),
      format: `${imported.artifact.format.family}/${imported.artifact.format.variant}`,
      nativeMetadata: { schema: 'NativeDocumentMetadata', codecRevision: imported.artifact.codecRevision, detectedFeatures: imported.artifact.detectedFeatures, compatibility: imported.report },
      source: 'document-import', spaceId: input.spaceId, folderId: input.folderId,
    });
    return { entry: remoteEntry(response.summary), snapshot: response.snapshot, report: imported.report, artifact: imported.artifact };
  }

  async exportWorkbook(unitId: string, input: WorkbookCatalogExportInput = {}): Promise<WorkbookCatalogExportResult> {
    const api = this.requireRemote();
    const resolved = await this.resolve(unitId);
    const summary = await api.getWorkbookSummary(unitId);
    let artifact: NativeDocumentArtifact | undefined;
    if (summary.sourceFileName) {
      const source = await api.getWorkbookSourceArtifact(unitId);
      artifact = (await exchangeImportDocument({ fileName: source.metadata.fileName, buffer: await source.artifact.arrayBuffer(), execution: 'worker' })).artifact;
    }
    const fileName = input.fileName ?? artifact?.fileName ?? `${resolved.snapshot.name || 'workbook'}.xlsx`;
    const exported = await exchangeSaveAsDocument(resolved.snapshot, { ...input, fileName, artifact });
    if (!exported.buffer || !exported.fileName) throw new WorkbookCatalogError('invalid-input', '导出未生成文件');
    return { unitId, fileName: exported.fileName, buffer: exported.buffer, report: exported.report };
  }

  async syncToServer(unitId: string): Promise<WorkbookCatalogSyncResult> {
    const api = this.requireRemote();
    const checkpoint = await api.checkpointWorkbook(unitId);
    return { entry: remoteEntry(await api.getWorkbookSummary(unitId)), committedOperationCount: 0, revision: checkpoint.revision };
  }

  async rename(unitId: string, name: string): Promise<WorkbookCatalogEntry> {
    const trimmed = name.trim();
    if (!trimmed || trimmed.length > MAX_WORKBOOK_NAME_LENGTH) throw new WorkbookCatalogError('invalid-input', '工作簿名称无效');
    const api = this.requireRemote();
    const current = await api.getSnapshot(unitId);
    await api.commitOperation(unitId, buildOperation(crypto.randomUUID(), unitId, 1, current.revision,
      [{ id: 'workbook.renamed', sheetId: current.snapshot.sheets[0]!.id, params: { name: trimmed } }]));
    return remoteEntry(await api.getWorkbookSummary(unitId));
  }

  async copy(unitId: string, request: { name?: string; spaceId?: string; folderId?: string; destination?: 'local' | 'remote' } = {}): Promise<WorkbookCatalogEntry> {
    if (request.destination === 'local') throw new WorkbookCatalogError('invalid-input', '副本必须保存到服务器');
    return remoteEntry(await this.requireRemote().copyWorkbook(unitId, request));
  }

  async move(unitId: string, input: WorkbookCatalogMoveInput): Promise<WorkbookCatalogEntry> {
    return remoteEntry(await this.requireRemote().updateWorkbook(unitId, input));
  }

  async grantAccess(unitId: string, subject: string, role: WorkbookRole): Promise<void> {
    if (!subject.trim() || role === 'owner') throw new WorkbookCatalogError('invalid-input', '共享用户或角色无效');
    await this.requireRemote().putWorkbookAcl(unitId, subject.trim(), role);
  }

  async revokeAccess(unitId: string, subject: string): Promise<void> {
    if (!subject.trim()) throw new WorkbookCatalogError('invalid-input', '共享用户不能为空');
    await this.requireRemote().deleteWorkbookAcl(unitId, subject.trim());
  }

  async moveToTrash(unitId: string): Promise<WorkbookCatalogEntry> {
    const api = this.requireRemote();
    await api.moveToTrash(unitId);
    return remoteEntry(await api.getWorkbookSummary(unitId));
  }

  async restore(unitId: string): Promise<WorkbookCatalogEntry> {
    return remoteEntry(await this.requireRemote().restoreFromTrash(unitId));
  }

  async purge(unitId: string): Promise<void> { await this.requireRemote().purgeWorkbook(unitId); }

  async setFavorite(unitId: string, favorite: boolean): Promise<WorkbookCatalogEntry> {
    const api = this.requireRemote();
    const state = await api.getWorkbookUserState(unitId);
    await api.putWorkbookUserState(unitId, { favorite, lastOpenedAt: state.lastOpenedAt });
    return remoteEntry(await api.getWorkbookSummary(unitId));
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

function nativeDocumentMimeType(fileName: string): string {
  const extension = fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  return extension === 'csv' ? 'text/csv'
    : extension === 'txt' || extension === 'prn' || extension === 'dif' || extension === 'slk' ? 'text/plain'
      : extension === 'xml' ? 'application/xml'
        : extension === 'ods' ? 'application/vnd.oasis.opendocument.spreadsheet'
          : extension === 'sjs' ? 'application/zip'
            : extension === 'ssjson' ? 'application/json'
              : 'application/octet-stream';
}
