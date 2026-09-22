import { Box, Button, Text, TextInput, DataTable } from '@react-sheets/ui-system';
import { createContext, useContext, useEffect, useSyncExternalStore, type ReactNode } from 'react';
import type { AuthSession, AuthSnapshot } from './oidc';
import { getAuthSession } from './session';
import { LoginPage } from './LoginPage';

const AuthContext = createContext<AuthSession | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const session = getAuthSession();
  const snapshot = useSyncExternalStore(session.subscribe, session.getSnapshot, session.getSnapshot);
  useEffect(() => {
    void session.initialize();
  }, [session]);
  const shared = Boolean(new URLSearchParams(window.location.search).get('share'));
  if (snapshot.phase === 'loading') return <main className="p-8" role="status">正在连接工作簿服务…</main>;
  if (snapshot.phase === 'error') return <main className="p-8"><p role="alert">{snapshot.error}</p><Button onClick={() => void session.initialize()}>重新连接</Button></main>;
  if (snapshot.mode === 'local' && snapshot.phase !== 'authenticated' && !shared) return <LoginPage snapshot={snapshot} />;
  return <AuthContext.Provider value={session}><Box key={snapshot.subject ?? 'guest'}>{children}</Box></AuthContext.Provider>;
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
