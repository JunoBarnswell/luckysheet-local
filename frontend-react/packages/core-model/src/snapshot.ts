import type {
  DefinedNameModel,
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
  unitId: UnitId;
  name: string;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
  tables?: WorkbookTableModel[];
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  sheets: SheetSnapshot[];
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(snapshot);
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
