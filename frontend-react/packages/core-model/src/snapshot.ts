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
  /** Missing values are legacy v1 snapshots and are migrated on load. */
  version?: 1 | 2;
  unitId: UnitId;
  name: string;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
  tables?: WorkbookTableModel[];
  /** Metadata only; large source bytes live in the local data-block store. */
  dataSources?: DataSourceManifest[];
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  sheets: SheetSnapshot[];
}

export const CURRENT_WORKBOOK_SNAPSHOT_VERSION = 2 as const;

/**
 * Keep migration at the snapshot boundary so callers never need a legacy
 * runtime branch. Feature-specific migrations are performed by their model
 * normalizers while this function guarantees the envelope shape.
 */
export function migrateWorkbookSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const version = snapshot.version ?? 1;
  if (version !== 1 && version !== CURRENT_WORKBOOK_SNAPSHOT_VERSION) {
    throw new Error(`Unsupported workbook snapshot version: ${String(snapshot.version)}`);
  }
  return {
    ...structuredClone(snapshot),
    version: CURRENT_WORKBOOK_SNAPSHOT_VERSION,
    dataSources: structuredClone(snapshot.dataSources ?? []),
  };
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(migrateWorkbookSnapshot(snapshot));
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
