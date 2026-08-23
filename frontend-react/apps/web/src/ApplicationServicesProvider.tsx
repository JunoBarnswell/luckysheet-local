import { createContext, useContext, useMemo, type ReactNode } from 'react';
import { WorkbookApiClient } from '@react-sheets/protocol';
import { WorkbookCatalogService, WorkspacePersistence } from '@react-sheets/spreadsheet-app';
import { useAuthSession } from './auth/AuthProvider';

export interface ApplicationServices {
  catalog: WorkbookCatalogService;
  persistence: WorkspacePersistence;
  workbookApi: WorkbookApiClient;
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
    return { catalog, persistence, workbookApi };
  }, [auth]);
  return <ApplicationServicesContext.Provider value={services}>{children}</ApplicationServicesContext.Provider>;
}

export function useApplicationServices(): ApplicationServices {
  const services = useContext(ApplicationServicesContext);
  if (!services) throw new Error('useApplicationServices must be used inside ApplicationServicesProvider');
  return services;
}
