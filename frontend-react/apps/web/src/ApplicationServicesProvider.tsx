import { createContext, useContext, type ReactNode } from 'react';
import { useSdkServices } from '@react-sheets/sdk';
import { sdk } from './sdk';

type ApplicationServices = ReturnType<typeof useSdkServices>;
const ApplicationServicesContext = createContext<ApplicationServices | null>(null);

/** UI context host; runtime construction and lifecycle are owned by SDK. */
export function ApplicationServicesProvider({ children }: { children: ReactNode }) {
  const services = useSdkServices(sdk);
  return <ApplicationServicesContext.Provider value={services}>{children}</ApplicationServicesContext.Provider>;
}
export function useApplicationServices(): ApplicationServices {
  const services = useContext(ApplicationServicesContext);
  if (!services) throw new Error('useApplicationServices must be used inside ApplicationServicesProvider');
  return services;
}
