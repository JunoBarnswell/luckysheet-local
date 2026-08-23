import { createContext, useContext, useEffect, useSyncExternalStore, type ReactNode } from 'react';
import { getAuthSession, type AuthSession, type AuthSnapshot } from './oidc';

const AuthContext = createContext<AuthSession | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const session = getAuthSession();
  useEffect(() => {
    void session.initialize();
  }, [session]);
  return <AuthContext.Provider value={session}>{children}</AuthContext.Provider>;
}

export function useAuthSession(): AuthSession {
  const session = useContext(AuthContext);
  if (!session) throw new Error('useAuthSession must be used inside AuthProvider');
  return session;
}

export function useAuthSnapshot(): AuthSnapshot {
  const session = useAuthSession();
  return useSyncExternalStore(session.subscribe, session.getSnapshot, session.getSnapshot);
}
