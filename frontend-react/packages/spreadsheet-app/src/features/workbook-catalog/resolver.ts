import {
  ApiRequestError,
  type ApiRequestOptions,
  type ShareTokenProvider,
} from '@react-sheets/protocol';
import type { WorkspacePersistence, WorkspaceRecord } from '../persistence/storage';
import type {
  WorkbookCatalogRemoteClient,
  WorkbookResolution,
} from './types';

export type WorkbookResolutionErrorCode = 'not-found' | 'permission-denied' | 'remote-unavailable' | 'memory-session-reset' | 'invalid-input';

export class WorkbookResolutionError extends Error {
  readonly code: WorkbookResolutionErrorCode;

  constructor(code: WorkbookResolutionErrorCode, message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'WorkbookResolutionError';
    this.code = code;
  }
}

export function isWorkbookResolutionError(error: unknown): error is WorkbookResolutionError {
  if (error instanceof WorkbookResolutionError) return true;
  if (!error || typeof error !== 'object') return false;
  const candidate = error as { name?: unknown; code?: unknown };
  return candidate.name === 'WorkbookResolutionError'
    && (candidate.code === 'not-found'
      || candidate.code === 'permission-denied'
      || candidate.code === 'remote-unavailable'
      || candidate.code === 'memory-session-reset'
      || candidate.code === 'invalid-input');
}

export interface WorkbookResolverOptions {
  persistence: WorkspacePersistence;
  remote?: WorkbookCatalogRemoteClient;
  remoteAvailable?: () => boolean;
  shareTokenProvider?: ShareTokenProvider;
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

function isRemoteUnavailable(error: unknown): boolean {
  if (error instanceof ApiRequestError) return error.status === 408 || error.status === 429 || error.status >= 500;
  if (error && typeof error === 'object' && 'status' in error && typeof error.status === 'number') {
    return error.status === 0 || error.status >= 500;
  }
  return error instanceof TypeError;
}

function toResolutionError(error: unknown, unitId: string): Error {
  if (error instanceof ApiRequestError) {
    if (error.status === 404) return new WorkbookResolutionError('not-found', `Workbook not found: ${unitId}`, { cause: error });
    if (error.status === 401 || error.status === 403) return new WorkbookResolutionError('permission-denied', `Workbook access denied: ${unitId}`, { cause: error });
    if (isRemoteUnavailable(error)) return new WorkbookResolutionError('remote-unavailable', `Cloud workbook service is unavailable: ${unitId}`, { cause: error });
  }
  return error instanceof Error ? error : new Error(`Workbook resolution failed: ${unitId}`);
}

function assertUnitId(unitId: string): string {
  const normalized = unitId.trim();
  if (!normalized) throw new WorkbookResolutionError('invalid-input', 'Workbook unitId is required');
  return normalized;
}

function localResolution(record: WorkspaceRecord, mode: 'local' | 'offline'): WorkbookResolution {
  return {
    schema: 'WorkbookResolution',
    unitId: record.unitId,
    source: mode === 'local' ? 'local' : 'mirrored',
    mode,
    lifecycle: 'active',
    binding: {
      location: record.metadata.location,
      syncMode: record.syncMode,
    },
    snapshot: clone(record.snapshot),
    revision: record.serverRevision,
    access: null,
    localRecord: clone(record),
  };
}

export class WorkbookResolver {
  private readonly persistence: WorkspacePersistence;
  private readonly remote?: WorkbookCatalogRemoteClient;
  private readonly remoteAvailable?: () => boolean;
  private readonly shareTokenProvider?: ShareTokenProvider;

  constructor(options: WorkbookResolverOptions) {
    this.persistence = options.persistence;
    this.remote = options.remote;
    this.remoteAvailable = options.remoteAvailable;
    this.shareTokenProvider = options.shareTokenProvider;
  }

  private canUseRemote(): boolean {
    return Boolean(this.remote && (this.remoteAvailable ? this.remoteAvailable() : true));
  }

  private requireRemote(): WorkbookCatalogRemoteClient {
    if (!this.remote || !this.canUseRemote()) throw new WorkbookResolutionError('remote-unavailable', 'Cloud workbook service is unavailable');
    return this.remote;
  }

  async resolve(unitId: string, options: ApiRequestOptions = {}): Promise<WorkbookResolution> {
    const normalized = assertUnitId(unitId);
    const localRecord = await this.persistence.store.open(normalized);
    if (localRecord?.metadata.lifecycle === 'trashed') {
      throw new WorkbookResolutionError('not-found', `Workbook is in trash: ${normalized}`);
    }

    if (localRecord && (localRecord.syncMode === 'local-only' || localRecord.metadata.location === 'local')) {
      return localResolution(localRecord, 'local');
    }

    if (!this.canUseRemote()) {
      if (localRecord) return localResolution(localRecord, 'offline');
      throw new WorkbookResolutionError(
        'memory-session-reset',
        `The page memory session no longer contains workbook: ${normalized}`,
      );
    }

    try {
      const remote = this.requireRemote();
      const [snapshotResponse, access] = await Promise.all([
        remote.getSnapshot(normalized, options),
        remote.getAccess(normalized, options),
      ]);
      const isShared = Boolean((await this.shareTokenProvider?.())?.trim());
      return {
        schema: 'WorkbookResolution',
        unitId: normalized,
        source: localRecord ? 'mirrored' : isShared ? 'shared' : 'remote',
        mode: 'remote',
        lifecycle: 'active',
        binding: { location: 'remote', syncMode: 'remote' },
        snapshot: clone(snapshotResponse.snapshot),
        revision: snapshotResponse.revision,
        access,
        localRecord: localRecord ? clone(localRecord) : null,
      };
    } catch (error) {
      // A cached mirrored workbook can be opened offline, but an authoritative
      // 401/403/404 must not be hidden by stale local data.
      if (localRecord && isRemoteUnavailable(error)) return localResolution(localRecord, 'offline');
      throw toResolutionError(error, normalized);
    }
  }
}
