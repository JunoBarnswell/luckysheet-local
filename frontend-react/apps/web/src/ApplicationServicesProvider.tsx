import { createContext, useContext, useMemo, type ReactNode } from 'react';
import { WorkbookApiClient } from '@react-sheets/protocol';
import { RemoteAssetStore, WorkbookCatalogService, WorkspacePersistence, type WorkbookSessionOptions } from '@react-sheets/spreadsheet-app';
import type { AuthTokenProvider } from '@react-sheets/protocol';
import { useAuthSession } from './auth/AuthProvider';

export interface ApplicationServices {
  catalog: WorkbookCatalogService;
  persistence: WorkspacePersistence;
  workbookApi: WorkbookApiClient;
  createWorkbookSessionOptions: (unitId: string, authTokenProvider: AuthTokenProvider) => WorkbookSessionOptions;
}

const ApplicationServicesContext = createContext<ApplicationServices | null>(null);

export function ApplicationServicesProvider({ children }: { children: ReactNode }) {
  const auth = useAuthSession();
  const services = useMemo<ApplicationServices>(() => {
    const persistence = new WorkspacePersistence();
    const workbookApi = new WorkbookApiClient({ authTokenProvider: auth.getAccessToken });
    const catalog = new WorkbookCatalogService({
      persistence,
      remote: workbookApi,
      remoteAvailable: () => auth.getSnapshot().phase === 'authenticated',
    });
    const createWorkbookSessionOptions = (unitId: string, authTokenProvider: AuthTokenProvider): WorkbookSessionOptions => ({
      unitId,
      workspacePersistence: persistence,
      authTokenProvider,
      assetStore: new RemoteAssetStore(unitId, workbookApi),
    });
    return { catalog, persistence, workbookApi, createWorkbookSessionOptions };
  }, [auth]);
  return <ApplicationServicesContext.Provider value={services}>{children}</ApplicationServicesContext.Provider>;
}

export function useApplicationServices(): ApplicationServices {
  const services = useContext(ApplicationServicesContext);
  if (!services) throw new Error('useApplicationServices must be used inside ApplicationServicesProvider');
  return services;
}
