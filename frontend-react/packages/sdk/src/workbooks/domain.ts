import type { Workbook } from '../workbook/workbook';
import { SdkError } from '../error';
import type { AssetStore } from '../../../spreadsheet-app/src/features/persistence';
import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { NativeDocumentError, type NativeDocumentArtifact } from '@react-sheets/exchange-excel-ooxml';
import {
  ApiRequestError,
  AuthenticationRequiredError,
  MAX_WORKBOOK_NAME_LENGTH,
  isWorkbookRole,
  type OperationEnvelope,
  type SnapshotResponse,
  type WorkbookCatalogQuery as ProtocolWorkbookCatalogQuery,
  type WorkbookCreateMetadata,
  type WorkbookSummary,
} from '@react-sheets/protocol';
import { buildOperation } from '../../../spreadsheet-app/src/collaboration';

import {
  exchangeImportDocument,
  exchangeSaveAsDocument,
  exchangeSaveDocument,
} from '../../../spreadsheet-app/src/features/native-document';
import {
  type WorkspaceRecord,
  type WorkspaceRecordMetadata,
  type WorkspaceUserState,
  WorkspacePersistence,
} from '../../../spreadsheet-app/src/features/persistence';

import { createTemplateSnapshot, createWorkbookUnitId, getWorkbookTemplate } from '../../../spreadsheet-app/src/features/workbook-catalog/templates';
import type { WorkbookCreateOptions } from './contract';
import type {
  WorkbookCatalogEntry,
  WorkbookCatalogExportInput,
  WorkbookCatalogExportResult,
  WorkbookCatalogImportInput,
  WorkbookCatalogImportResult,

  WorkbookCatalogQuery,
  WorkbookCatalogRequestOptions,
  WorkbookCatalogRemoteClient,
  WorkbookRole,
  WorkbookResolution,
} from '../../../spreadsheet-app/src/features/workbook-catalog/types';
import { WorkbookResolver } from '../../../spreadsheet-app/src/features/workbook-catalog/resolver';

import { workbookCapabilities, type WorkbookCapabilities } from '../identity/workbook-capabilities';

export interface CatalogEntry extends WorkbookCatalogEntry { readonly capabilities: WorkbookCapabilities; }
export interface CatalogPage { entries: CatalogEntry[]; nextCursor: string | null; }

export const DEFAULT_NATIVE_IMPORT_MAX_BYTES = 50 * 1024 * 1024;

export class WorkbookCatalogError extends Error {
  readonly code: 'not-found' | 'permission-denied' | 'remote-unavailable' | 'invalid-input' | 'conflict';

  constructor(code: WorkbookCatalogError['code'], message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'WorkbookCatalogError';
    this.code = code;
  }
}

export interface WorkbooksDomainOptions {
  openWorkbook?: (resolution: WorkbookResolution) => Promise<Workbook>;
  persistence?: WorkspacePersistence;
  assetStoreFor?: (unitId: string) => AssetStore;
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
  entry: CatalogEntry;
  committedOperationCount: number;
  revision: number;
}

function requireRole(role: unknown): WorkbookRole {
  if (!isWorkbookRole(role)) throw new WorkbookCatalogError('invalid-input', '服务器目录缺少 canonical 工作簿角色。');
  return role;
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
    role: requireRole(summary.role),
    ownerId: summary.ownerSubject,
    sourceFileName: summary.sourceFileName,
    spaceId: summary.spaceId,
    folderId: summary.folderId,
    locationPath: summary.locationPath?.join(' / '),
    deletedAt: summary.deletedAt,
  };
}

function remoteEntry(summary: WorkbookSummary): CatalogEntry {
  const metadata = metadataFromRemote(summary);
  return {
    unitId: summary.unitId,
    name: summary.name,
    revision: summary.revision,
    updatedAt: summary.updatedAt,
    storage: summary.storageLocation ?? 'remote',
    syncState: summary.syncStatus ?? 'synced',
    role: metadata.role,
    capabilities: workbookCapabilities(metadata.role, metadata.lifecycle),
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

export class WorkbooksDomain {
  private retired = false;
  retire(): void { this.retired = true; }
  readonly #persistence: WorkspacePersistence;
  readonly #openWorkbook?: (resolution: WorkbookResolution) => Promise<Workbook>;
  readonly #remote?: WorkbookCatalogRemoteClient;
  readonly #resolver: WorkbookResolver;
  private readonly now: () => Date;
  private readonly unitIdFactory: () => string;
  readonly #remoteAvailable?: () => boolean;
  readonly #assetStoreFor?: (unitId: string) => AssetStore;

  constructor(options: WorkbooksDomainOptions = {}) {
    this.#openWorkbook = options.openWorkbook;
    this.#persistence = options.persistence ?? new WorkspacePersistence();
    this.#remote = options.remote;
    this.#assetStoreFor = options.assetStoreFor;
    this.now = options.now ?? (() => new Date());
    this.unitIdFactory = options.unitIdFactory ?? (() => createWorkbookUnitId());
    this.#remoteAvailable = options.remoteAvailable;
    this.#resolver = new WorkbookResolver({
      persistence: this.#persistence,
      remote: this.#remote,
      remoteAvailable: () => this.canUseRemote(),
      shareTokenProvider: options.shareTokenProvider,
    });
  }

  private canUseRemote(): boolean {
    return Boolean(!this.retired && this.#remote && (this.#remoteAvailable ? this.#remoteAvailable() : true));
  }

  private requireRemote(): WorkbookCatalogRemoteClient {
    if (!this.#remote || !this.canUseRemote()) throw new WorkbookCatalogError('remote-unavailable', 'Cloud workbook service is unavailable');
    return this.#remote;
  }

  async listPage(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<CatalogPage> {
    if (query.view === 'local') throw new WorkbookCatalogError('invalid-input', '页面内存文件已移除，请使用服务器文件中心');
    const page = await this.requireRemote().listWorkbookPage(query as ProtocolWorkbookCatalogQuery, options);
    return { entries: page.items.map(remoteEntry), nextCursor: page.nextCursor };
  }

  async list(query: WorkbookCatalogQuery = {}, options: WorkbookCatalogRequestOptions = {}): Promise<CatalogEntry[]> {
    const page = await this.listPage(query, options);
    return page.entries;
  }

  async create(input: WorkbookCreateOptions): Promise<CatalogEntry> {
    try {
      if (!input || typeof input.name !== 'string' || !input.name.trim() || input.name.trim().length > MAX_WORKBOOK_NAME_LENGTH
        || [input.spaceId, input.folderId].some(value => value !== undefined && (typeof value !== 'string' || !value.trim()))) {
        throw new SdkError('INVALID_ARGUMENT', 'workbooks.create', '工作簿名称或目录身份无效。', '请提供有效名称与稳定目录 ID。');
      }
      const template = input.template ?? 'blank';
      try { getWorkbookTemplate(template); }
      catch (cause) { throw new SdkError('INVALID_ARGUMENT', 'workbooks.create', '工作簿模板无效。', '请选择公开模板 ID。', { cause }); }
      const snapshot = createTemplateSnapshot(template, this.unitIdFactory(), input.name.trim());
      const response = await this.requireRemote().createWorkbook(snapshot, { spaceId: input.spaceId, folderId: input.folderId });
      const entry = remoteEntry(await this.requireRemote().getWorkbookSummary(response.snapshot.unitId));
      this.requireRemote();
      return entry;
    } catch (cause) {
      if (cause instanceof SdkError) throw cause;
      const code = cause instanceof ApiRequestError && cause.status === 403 ? 'FORBIDDEN' : 'REQUEST_REJECTED';
      throw new SdkError(code, 'workbooks.create', cause instanceof Error ? cause.message : '创建失败。', '请检查认证、目录权限和服务端连接后重试。', { cause });
    }
  }

  async open(unitId: string): Promise<Workbook> {
    if (typeof unitId !== 'string' || !unitId.trim()) throw new SdkError('INVALID_ARGUMENT', 'workbooks.open', '工作簿 ID 无效。', '请使用目录返回的稳定 unitId。');
    this.requireRemote();
    if (!this.#openWorkbook) throw new SdkError('UNSUPPORTED_FEATURE', 'workbooks.open', '目录缺少 canonical Workbook 运行宿主。', '请通过 createSpreadsheetSdk 获取工作簿目录。');
    try {
      const resolution = await this.resolve(unitId);
      this.requireRemote();
      const workbook = await this.#openWorkbook(resolution);
      this.requireRemote();
      return workbook;
    } catch (cause) {
      if (cause instanceof SdkError) throw cause;
      throw new SdkError('REQUEST_REJECTED', 'workbooks.open', `${unitId}: ${cause instanceof Error ? cause.message : '打开失败。'}`, '请检查权限、工作簿身份与服务端连接后重试。', { cause, object: { workbookId: unitId } });
    }
  }

  resolve(unitId: string, options: WorkbookCatalogRequestOptions = {}): Promise<WorkbookResolution> {
    return this.#resolver.resolve(unitId, options);
  }

  async markOpened(resolution: WorkbookResolution): Promise<CatalogEntry> {
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
    this.requireRemote();
    let exported: Awaited<ReturnType<typeof exchangeSaveAsDocument>>;
    try {
      exported = await exchangeSaveAsDocument(resolved.snapshot, {
        ...input, fileName, artifact, assetStore: this.#assetStoreFor?.(unitId),
      });
    } catch (cause) {
      if (cause instanceof NativeDocumentError) throw cause;
      throw new SdkError('REQUEST_REJECTED', 'workbooks.export',
        `工作簿 ${unitId} 导出失败：${cause instanceof Error ? cause.message : String(cause)}`,
        '请确认源文档和引用的资产完整、仍有读取权限，再重试导出。', { cause });
    }
    this.requireRemote();
    if (!exported.buffer || !exported.fileName) throw new WorkbookCatalogError('invalid-input', '导出未生成文件');
    return { unitId, fileName: exported.fileName, buffer: exported.buffer, report: exported.report };
  }

  async syncToServer(unitId: string): Promise<WorkbookCatalogSyncResult> {
    const api = this.requireRemote();
    const checkpoint = await api.checkpointWorkbook(unitId);
    return { entry: remoteEntry(await api.getWorkbookSummary(unitId)), committedOperationCount: 0, revision: checkpoint.revision };
  }

  async rename(unitId: string, name: string): Promise<CatalogEntry> {
    const trimmed = name.trim();
    if (!trimmed || trimmed.length > MAX_WORKBOOK_NAME_LENGTH) throw new WorkbookCatalogError('invalid-input', '工作簿名称无效');
    const api = this.requireRemote();
    const current = await api.getSnapshot(unitId);
    await api.commitOperation(unitId, buildOperation(crypto.randomUUID(), unitId, 1, current.revision,
      [{ id: 'workbook.renamed', sheetId: current.snapshot.sheets[0]!.id, params: { name: trimmed } }]));
    return remoteEntry(await api.getWorkbookSummary(unitId));
  }

  async copy(unitId: string, request: { name?: string; spaceId?: string; folderId?: string; destination?: 'local' | 'remote' } = {}): Promise<CatalogEntry> {
    if (request.destination === 'local') throw new WorkbookCatalogError('invalid-input', '副本必须保存到服务器');
    return remoteEntry(await this.requireRemote().copyWorkbook(unitId, request));
  }

  async move(unitId: string, input: WorkbookCatalogMoveInput): Promise<CatalogEntry> {
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

  async moveToTrash(unitId: string): Promise<CatalogEntry> {
    const api = this.requireRemote();
    await api.moveToTrash(unitId);
    return remoteEntry(await api.getWorkbookSummary(unitId));
  }

  async restore(unitId: string): Promise<CatalogEntry> {
    return remoteEntry(await this.requireRemote().restoreFromTrash(unitId));
  }

  async purge(unitId: string): Promise<void> { await this.requireRemote().purgeWorkbook(unitId); }

  async setFavorite(unitId: string, favorite: boolean): Promise<CatalogEntry> {
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
    return this.requireRemote().putSpaceMember(spaceId, subject.trim(), requireRole(role));
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
