import { installXlsxWorkerEntry } from './worker-entry';

installXlsxWorkerEntry(self as unknown as import('./worker-entry').XlsxWorkerScope);
