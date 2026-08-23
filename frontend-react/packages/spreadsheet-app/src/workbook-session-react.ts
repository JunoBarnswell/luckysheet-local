import { useEffect, useRef, useSyncExternalStore } from 'react';
import { WorkbookSession, type WorkbookSessionOptions, type UiSnapshot } from './workbook-session';

export interface UseWorkbookSessionResult {
  session: WorkbookSession;
  snapshot: UiSnapshot;
}

/**
 * The application shell owns one factory and passes the selected workbook id
 * explicitly. A factory prevents a session from deriving identity from
 * `window.location`, while still allowing tests to construct a local session
 * directly with `new WorkbookSession()`.
 */
export interface WorkbookSessionFactory {
  create(unitId: string, options?: Omit<WorkbookSessionOptions, 'unitId'>): WorkbookSession;
}

export function createWorkbookSessionFactory(defaults: Omit<WorkbookSessionOptions, 'unitId'> = {}): WorkbookSessionFactory {
  return {
    create: (unitId, options = {}) => {
      const normalized = unitId.trim();
      if (!normalized) throw new Error('Workbook session unitId is required');
      return new WorkbookSession({ ...defaults, ...options, unitId: normalized });
    },
  };
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
