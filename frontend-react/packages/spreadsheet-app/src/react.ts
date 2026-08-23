import { useEffect, useRef, useSyncExternalStore } from 'react';
import { SpreadsheetApplication, type SpreadsheetApplicationOptions, type UiSnapshot } from './application';

export interface UseSpreadsheetAppResult {
  app: SpreadsheetApplication;
  snapshot: UiSnapshot;
}

export function useSpreadsheetApp(options: SpreadsheetApplicationOptions = {}): UseSpreadsheetAppResult {
  const appRef = useRef<SpreadsheetApplication | null>(null);
  if (!appRef.current) {
    appRef.current = new SpreadsheetApplication(options);
  }
  const app = appRef.current;

  useEffect(() => {
    app.start();
    return () => app.dispose();
  }, [app]);

  const snapshot = useSyncExternalStore(app.subscribe, app.getUiSnapshot, app.getUiSnapshot);

  return { app, snapshot };
}
