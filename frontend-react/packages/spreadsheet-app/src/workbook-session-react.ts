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
  const disposeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  if (!sessionRef.current) {
    sessionRef.current = new WorkbookSession(options);
  }
  const session = sessionRef.current;

  useEffect(() => {
    if (disposeTimer.current !== null) {
      clearTimeout(disposeTimer.current);
      disposeTimer.current = null;
    }
    session.start();
    return () => {
      // React StrictMode probes effect cleanup during development. Defer the
      // real dispose so the immediate probe remount can reuse the in-flight
      // persistence initialization instead of opening a second Writer.
      disposeTimer.current = setTimeout(() => {
        disposeTimer.current = null;
        session.dispose();
      }, 0);
    };
  }, [session]);

  const snapshot = useSyncExternalStore(session.subscribe, session.getUiSnapshot, session.getUiSnapshot);

  return { session, snapshot };
}
