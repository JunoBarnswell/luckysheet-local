import {
  installBrowserCalculationWorkerEntry,
  type BrowserCalculationWorkerScope,
} from './calculation-worker-entry';

// This module is emitted as a dedicated browser Worker by Vite through the
// direct new Worker(new URL(...)) call in calculation-browser-task-port.ts.
// The entry keeps the FormulaEngine alive for the whole workbook session.
installBrowserCalculationWorkerEntry(self as unknown as BrowserCalculationWorkerScope);
