import type {
  DefinedNameModel,
  DataSourceManifest,
  SheetId,
  SheetSnapshot,
  WorkbookModel,
  WorkbookTableModel,
  UnitId,
} from './index';
import type { PrintDocumentSnapshot, QueryDefinitionSnapshot } from './workbook-state';
import { WorkbookModel as WorkbookModelClass } from './index';

/**
 * The single persisted/transport snapshot contract. Floating objects are
 * represented only by the canonical drawing collection and payload map.
 * There is deliberately no versioned sibling or legacy union in production.
 */
export interface WorkbookSnapshot {
  schema: 'WorkbookSnapshot';
  /** Canonical persisted schema revision. Non-matching snapshots are rejected. */
  version: 2;
  unitId: UnitId;
  name: string;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
  tables?: WorkbookTableModel[];
  /** Metadata only; large source bytes live in the local data-block store. */
  dataSources: DataSourceManifest[];
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  sheets: SheetSnapshot[];
}

export const WORKBOOK_SNAPSHOT_SCHEMA_REVISION = 2 as const;

/**
 * The application reads one Snapshot contract only. A non-canonical record
 * must be repaired before it enters the workbook runtime.
 */
export function assertCanonicalWorkbookSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  if (snapshot.version !== WORKBOOK_SNAPSHOT_SCHEMA_REVISION) {
    throw new Error(`Unsupported workbook snapshot version: ${String(snapshot.version)}`);
  }
  if (!Array.isArray(snapshot.dataSources)) throw new Error('Workbook snapshot dataSources must be an array');
  return structuredClone(snapshot);
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(assertCanonicalWorkbookSnapshot(snapshot));
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
