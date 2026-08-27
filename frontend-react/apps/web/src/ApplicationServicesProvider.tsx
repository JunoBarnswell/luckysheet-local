import { createContext, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { WorkbookApiClient } from '@react-sheets/protocol';
import {
  RemoteAssetStore,
  LocalAssetStore,
  WorkbookCatalogService,
  WorkspacePersistence,
  WorkspaceStorageError,
  isWorkspaceStorageError,
  resolveShareToken,
  type WorkbookSessionOptions,
  type WorkspacePersistenceState,
} from '@react-sheets/spreadsheet-app';
import type { AuthTokenProvider } from '@react-sheets/protocol';
import { useAuthSession } from './auth/AuthProvider';

declare global {
  interface Window {
    reactSheetsDesktopConfig?: Readonly<{ collaborationUrl: string }>;
  }
}

export interface StorageReadiness {
  state: WorkspacePersistenceState | 'warming' | 'failed';
  error: WorkspaceStorageError | null;
}

export interface ApplicationServices {
  catalog: WorkbookCatalogService;
  persistence: WorkspacePersistence;
  ensureStorageReady: () => Promise<void>;
  retryStorage: () => Promise<void>;
  storageReadiness: StorageReadiness;
  workbookApi: WorkbookApiClient;
  createWorkbookSessionOptions: (unitId: string, authTokenProvider: AuthTokenProvider, useLocalAssets?: boolean) => WorkbookSessionOptions;
}

const ApplicationServicesContext = createContext<ApplicationServices | null>(null);

function formatStorageError(error: unknown): WorkspaceStorageError {
  if (isWorkspaceStorageError(error)) return error;
  return new WorkspaceStorageError({
    code: 'STORAGE_MEMORY_TRANSACTION_FAILED',
    operation: 'open',
    message: error instanceof Error ? error.message : '内存工作簿存储不可用。',
    recovery: '请重新开始页面内存会话后重试。',
    cause: error,
  });
}

function resolveDesktopCollaborationUrl(): string | undefined {
  if (typeof window === 'undefined' || window.location.protocol !== 'app:') return undefined;
  const candidate = window.reactSheetsDesktopConfig?.collaborationUrl;
  if (!candidate) throw new Error('Desktop runtime is missing its collaboration endpoint');
  const parsed = new URL(candidate);
  if (!['ws:', 'wss:'].includes(parsed.protocol) || parsed.username || parsed.password) {
    throw new Error('Desktop collaboration endpoint is invalid');
  }
  return parsed.toString();
}

export function ApplicationServicesProvider({ children }: { children: ReactNode }) {
  const auth = useAuthSession();
  const [storageReadiness, setStorageReadiness] = useState<StorageReadiness>({ state: 'warming', error: null });
  const disposeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const persistenceRef = useRef<WorkspacePersistence | null>(null);
  const core = useMemo(() => {
    const persistence = persistenceRef.current ?? (persistenceRef.current = new WorkspacePersistence());
    const ensureStorageReady = async (): Promise<void> => {
      setStorageReadiness({ state: 'warming', error: null });
      try {
        await persistence.ensureReady();
        setStorageReadiness({ state: persistence.state, error: null });
      } catch (cause) {
        const error = formatStorageError(cause);
        setStorageReadiness({ state: 'failed', error });
        throw error;
      }
    };
    const retryStorage = (): Promise<void> => ensureStorageReady();
    const shareTokenProvider = () => resolveShareToken();
    const workbookApi = new WorkbookApiClient({ authTokenProvider: auth.getAccessToken, shareTokenProvider });
    const catalog = new WorkbookCatalogService({
      persistence,
      remote: workbookApi,
      remoteAvailable: () => auth.getSnapshot().phase === 'authenticated' || Boolean(shareTokenProvider()),
      shareTokenProvider,
    });
    const createWorkbookSessionOptions = (unitId: string, authTokenProvider: AuthTokenProvider, useLocalAssets = false): WorkbookSessionOptions => ({
      unitId,
      api: workbookApi,
      workspacePersistence: persistence,
      authTokenProvider,
      shareTokenProvider,
      collaborationUrl: resolveDesktopCollaborationUrl(),
      assetStore: useLocalAssets ? new LocalAssetStore(unitId, persistence.coordinator) : new RemoteAssetStore(unitId, workbookApi),
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

  useEffect(() => {
    if (disposeTimer.current !== null) {
      clearTimeout(disposeTimer.current);
      disposeTimer.current = null;
    }
    return () => {
      // React StrictMode probes effect cleanup during development. Defer the
      // actual dispose by one task so the probe's immediate remount can cancel
      // it; a real provider unmount still releases the memory session.
      disposeTimer.current = setTimeout(() => {
        disposeTimer.current = null;
        void core.persistence.disposeAsync();
      }, 0);
    };
  }, [core]);

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
