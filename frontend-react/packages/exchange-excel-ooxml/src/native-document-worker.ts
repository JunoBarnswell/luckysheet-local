import { installNativeDocumentWorkerEntry } from './worker-entry';

installNativeDocumentWorkerEntry(self as unknown as import('./worker-entry').NativeDocumentWorkerScope);
