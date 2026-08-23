import { useEffect, useRef, useSyncExternalStore } from 'react';
import { WorkbookSession, type WorkbookSessionOptions, type UiSnapshot } from './workbook-session';

export interface UseWorkbookSessionResult {
  session: WorkbookSession;
  snapshot: UiSnapshot;
}

export function useWorkbookSession(options: WorkbookSessionOptions = {}): UseWorkbookSessionResult {
  const sessionRef = useRef<WorkbookSession | null>(null);
  if (!sessionRef.current) {
    sessionRef.current = new WorkbookSession(options);
  }
  const session = sessionRef.current;

  useEffect(() => {
    session.start();
    return () => session.dispose();
  }, [session]);

  const snapshot = useSyncExternalStore(session.subscribe, session.getUiSnapshot, session.getUiSnapshot);

  return { session, snapshot };
}
