import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';
import { WorkbookApiClient } from '@react-sheets/protocol';
import {
  RemoteAssetStore,
  WorkbookCatalogService,
  WorkspacePersistence,
  WorkspaceStorageError,
  isWorkspaceStorageError,
  resolveShareToken,
  type WorkbookSessionOptions,
  type WorkspaceDatabaseState,
} from '@react-sheets/spreadsheet-app';
import type { AuthTokenProvider } from '@react-sheets/protocol';
import { useAuthSession } from './auth/AuthProvider';

export interface StorageReadiness {
  state: WorkspaceDatabaseState | 'warming';
  error: WorkspaceStorageError | null;
}

export interface ApplicationServices {
  catalog: WorkbookCatalogService;
  persistence: WorkspacePersistence;
  ensureStorageReady: () => Promise<IDBDatabase>;
  retryStorage: () => Promise<IDBDatabase>;
  storageReadiness: StorageReadiness;
  workbookApi: WorkbookApiClient;
  createWorkbookSessionOptions: (unitId: string, authTokenProvider: AuthTokenProvider) => WorkbookSessionOptions;
}

const ApplicationServicesContext = createContext<ApplicationServices | null>(null);

function formatStorageError(error: unknown): WorkspaceStorageError {
  if (error instanceof WorkspaceStorageError) return error;
  if (isWorkspaceStorageError(error)) {
    const candidate = error as Partial<WorkspaceStorageError>;
    if (typeof candidate.code === 'string' && typeof candidate.databaseName === 'string' && typeof candidate.operation === 'string') {
      return new WorkspaceStorageError({
        code: candidate.code as WorkspaceStorageError['code'],
        databaseName: candidate.databaseName,
        operation: candidate.operation,
        message: String(candidate.message),
        recovery: typeof candidate.recovery === 'string' ? candidate.recovery : '请保留当前页面并重试。',
        cause: error,
      });
    }
  }
  return new WorkspaceStorageError({
    code: 'STORAGE_UNAVAILABLE',
    databaseName: 'react-sheets-workspaces',
    operation: 'open',
    message: error instanceof Error ? error.message : '本地工作簿存储不可用。',
    recovery: '请完整刷新当前页面后重试。',
    cause: error,
  });
}

export function ApplicationServicesProvider({ children }: { children: ReactNode }) {
  const auth = useAuthSession();
  const [storageReadiness, setStorageReadiness] = useState<StorageReadiness>({ state: 'warming', error: null });
  const core = useMemo(() => {
    const persistence = new WorkspacePersistence();
    const ensureStorageReady = async (): Promise<IDBDatabase> => {
      setStorageReadiness({ state: 'opening', error: null });
      try {
        const database = await persistence.coordinator.open();
        setStorageReadiness({ state: persistence.coordinator.state, error: null });
        return database;
      } catch (cause) {
        const error = formatStorageError(cause);
        setStorageReadiness({ state: 'failed', error });
        throw error;
      }
    };
    const retryStorage = (): Promise<IDBDatabase> => ensureStorageReady();
    const shareTokenProvider = () => resolveShareToken();
    const workbookApi = new WorkbookApiClient({ authTokenProvider: auth.getAccessToken, shareTokenProvider });
    const catalog = new WorkbookCatalogService({
      persistence,
      remote: workbookApi,
      remoteAvailable: () => auth.getSnapshot().phase === 'authenticated' || Boolean(shareTokenProvider()),
      shareTokenProvider,
    });
    const createWorkbookSessionOptions = (unitId: string, authTokenProvider: AuthTokenProvider): WorkbookSessionOptions => ({
      unitId,
      api: workbookApi,
      workspacePersistence: persistence,
      authTokenProvider,
      shareTokenProvider,
      assetStore: new RemoteAssetStore(unitId, workbookApi),
    });
    return { catalog, persistence, ensureStorageReady, retryStorage, workbookApi, createWorkbookSessionOptions };
  }, [auth]);

  useEffect(() => {
    let cancelled = false;
    void core.ensureStorageReady().catch((cause) => {
      if (cancelled) return;
      const error = formatStorageError(cause);
      setStorageReadiness({ state: 'failed', error });
    });
    return () => { cancelled = true; };
  }, [core]);

  useEffect(() => () => { void core.persistence.disposeAsync(); }, [core]);

  const services = useMemo<ApplicationServices>(() => ({
    ...core,
    storageReadiness,
  }), [core, storageReadiness]);

  return <ApplicationServicesContext.Provider value={services}>{children}</ApplicationServicesContext.Provider>;
}

export function useApplicationServices(): ApplicationServices {
  const services = useContext(ApplicationServicesContext);
  if (!services) throw new Error('useApplicationServices must be used inside ApplicationServicesProvider');
  return services;
}
