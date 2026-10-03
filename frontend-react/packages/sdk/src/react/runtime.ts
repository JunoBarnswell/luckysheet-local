import { useEffect, useRef, useSyncExternalStore } from 'react';
import type { WorkbookResolution, WorkbookSession } from '@react-sheets/spreadsheet-app';
import { runtimeFor, type SpreadsheetSdk } from '../sdk';

export function useSdkServices(sdk: SpreadsheetSdk) {
  const runtime = runtimeFor(sdk);
  const storageReadiness = useSyncExternalStore(runtime.subscribe, runtime.getSnapshot, runtime.getSnapshot);
  useEffect(() => {
    const release = runtime.acquire();
    void runtime.ensureStorageReady().catch(() => { /* The observable storage snapshot owns this failure. */ });
    return release;
  }, [runtime]);
  return { catalog: sdk.workbooks, ensureStorageReady: runtime.ensureStorageReady, retryStorage: runtime.retryStorage, storageReadiness };
}

/** SDK owns construction, credentials, persistence and browser worker policy. */
export function useWorkbook(sdk: SpreadsheetSdk, resolution: WorkbookResolution) {
  const runtime = runtimeFor(sdk);
  const sessionRef = useRef<WorkbookSession | null>(null);
  const disposeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  if (!sessionRef.current) sessionRef.current = runtime.createSession(resolution);
  const session = sessionRef.current;
  useEffect(() => {
    if (disposeTimer.current !== null) { clearTimeout(disposeTimer.current); disposeTimer.current = null; }
    session.start();
    return () => { disposeTimer.current = setTimeout(() => { disposeTimer.current = null; runtime.closeSession(session); }, 0); };
  }, [runtime, session]);
  const snapshot = useSyncExternalStore(session.subscribe, session.getUiSnapshot, session.getUiSnapshot);
  return { session, snapshot, data: runtime.dataActions(session), dimensions: runtime.dimensionActions(session) };
}
