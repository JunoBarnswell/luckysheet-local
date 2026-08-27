import { installPivotWorkerEntry, type PivotWorkerScope } from './task-worker-entry';

installPivotWorkerEntry(self as unknown as PivotWorkerScope);
