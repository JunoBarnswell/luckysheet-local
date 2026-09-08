import {
  ApiRequestError,
  type ApiRequestOptions,
  type ShareTokenProvider,
} from '@react-sheets/protocol';
import type {
  WorkbookCatalogRemoteClient,
  WorkbookResolution,
} from './types';

export type WorkbookResolutionErrorCode = 'not-found' | 'permission-denied' | 'remote-unavailable' | 'invalid-input';

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
      || candidate.code === 'invalid-input');
}

export interface WorkbookResolverOptions {
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
  }
  if (isRemoteUnavailable(error)) return new WorkbookResolutionError('remote-unavailable', `Cloud workbook service is unavailable: ${unitId}`, { cause: error });
  return error instanceof Error ? error : new Error(`Workbook resolution failed: ${unitId}`);
}

function assertUnitId(unitId: string): string {
  const normalized = unitId.trim();
  if (!normalized) throw new WorkbookResolutionError('invalid-input', 'Workbook unitId is required');
  return normalized;
}

async function loadRevisionPages(
  remote: WorkbookCatalogRemoteClient,
  unitId: string,
  manifest: Awaited<ReturnType<WorkbookCatalogRemoteClient['getManifest']>>,
  options: ApiRequestOptions,
) {
  const pages: WorkbookResolution['pages'][number][] = [];
  for (let offset = 0; offset < manifest.pages.length; offset += 4) {
    pages.push(...await Promise.all(manifest.pages.slice(offset, offset + 4).map((page) => remote.getPage({
      unitId,
      revision: manifest.revision,
      sheetId: page.sheetId,
      pageRow: page.pageRow,
      pageColumn: page.pageColumn,
    }, options))));
  }
  return pages;
}

export class WorkbookResolver {
  private readonly remote?: WorkbookCatalogRemoteClient;
  private readonly remoteAvailable?: () => boolean;
  private readonly shareTokenProvider?: ShareTokenProvider;

  constructor(options: WorkbookResolverOptions) {
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
    if (!this.canUseRemote()) {
      if (this.remote) throw new WorkbookResolutionError(
        'remote-unavailable',
        `Authoritative workbook service is unavailable: ${normalized}`,
      );
      throw new WorkbookResolutionError('remote-unavailable', `Cloud workbook service is unavailable: ${normalized}`);
    }

    try {
      const remote = this.requireRemote();
      const [manifest, access] = await Promise.all([
        remote.getManifest(normalized),
        remote.getAccess(normalized, options),
      ]);
      // WorksheetCells and feature resolvers are synchronous projections over
      // the committed revision. Resolve every sparse page before publishing a
      // usable route so no renderer or command can observe a partial model.
      const pages = await loadRevisionPages(remote, normalized, manifest, options);
      const isShared = Boolean((await this.shareTokenProvider?.())?.trim());
      return {
        schema: 'WorkbookResolution',
        unitId: normalized,
        source: isShared ? 'shared' : 'remote',
        mode: 'remote',
        lifecycle: 'active',
        manifest: clone(manifest),
        pages: clone(pages),
        revision: manifest.revision,
        access,
      };
    } catch (error) {
      throw toResolutionError(error, normalized);
    }
  }
}
