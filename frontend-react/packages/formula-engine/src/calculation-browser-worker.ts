import {
  installBrowserCalculationWorkerEntry,
  type BrowserCalculationWorkerScope,
} from './calculation-worker-entry';

// This module is emitted as a dedicated browser Worker by Vite through the
// direct new Worker(new URL(...)) call in calculation-browser-task-port.ts.
installBrowserCalculationWorkerEntry(self as unknown as BrowserCalculationWorkerScope);
