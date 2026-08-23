import type {
  BandedRule,
  CellData,
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  FilterModel,
  FloatingImage,
  FreezeModel,
  MergeSpan,
  PivotModel,
  ShapeModel,
  SheetId,
  SparklineModel,
  UnitId,
  WorkbookTableModel,
} from './index';
import type { CommentThread, CellNote, DrawingObject, ProtectionRule, SheetTableModel, SpillRange } from './domain';
import { WorkbookModel } from './index';

export interface SheetSnapshotV2 {
  id: SheetId;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, Record<string, CellData>>;
  merges: MergeSpan[];
  freeze: FreezeModel;
  charts: ChartModel[];
  pivots: PivotModel[];
  shapes: ShapeModel[];
  sparklines: SparklineModel[];
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
  rowHeights?: Record<number, number>;
  columnWidths?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  tabColor?: string;
  images?: FloatingImage[];
  bandedRule?: BandedRule;
  filter?: FilterModel;
  sheetTables?: SheetTableModel[];
  drawings?: DrawingObject[];
  notes?: Array<{ row: number; column: number; note: CellNote }>;
  commentThreads?: CommentThread[];
  spillRanges?: SpillRange[];
  protectionRules?: ProtectionRule[];
}

export interface WorkbookSnapshotV2 {
  schema: 'WorkbookSnapshotV2';
  schemaVersion: 2;
  unitId: UnitId;
  name: string;
  activeSheetId: SheetId;
  definedNames?: Record<string, string>;
  tables?: WorkbookTableModel[];
  protectionRules?: ProtectionRule[];
  sheets: SheetSnapshotV2[];
}

/** @deprecated use WorkbookSnapshotV2 */
export interface WorkbookSnapshotV1 {
  schema: 'WorkbookSnapshotV1';
  unitId: UnitId;
  name: string;
  activeSheetId: SheetId;
  definedNames?: Record<string, string>;
  tables?: WorkbookTableModel[];
  sheets: Array<Omit<SheetSnapshotV2, 'sheetTables' | 'drawings' | 'notes' | 'commentThreads' | 'spillRanges' | 'protectionRules'>>;
}

export type AnyWorkbookSnapshot = WorkbookSnapshotV1 | WorkbookSnapshotV2;

export function migrateSnapshot(snapshot: AnyWorkbookSnapshot): WorkbookSnapshotV2 {
  if (snapshot.schema === 'WorkbookSnapshotV2') return snapshot;
  return {
    schema: 'WorkbookSnapshotV2',
    schemaVersion: 2,
    unitId: snapshot.unitId,
    name: snapshot.name,
    activeSheetId: snapshot.activeSheetId,
    definedNames: snapshot.definedNames ? { ...snapshot.definedNames } : undefined,
    tables: snapshot.tables?.map((t) => structuredClone(t)),
    sheets: snapshot.sheets.map((sheet) => ({
      ...structuredClone(sheet),
      sheetTables: [],
      drawings: [],
      notes: [],
      commentThreads: [],
      spillRanges: [],
      protectionRules: [],
    })),
  };
}

export function loadWorkbookFromSnapshot(snapshot: AnyWorkbookSnapshot): WorkbookModel {
  const v2 = migrateSnapshot(snapshot);
  const workbook = WorkbookModel.fromSnapshot({
    schema: 'WorkbookSnapshotV1',
    unitId: v2.unitId,
    name: v2.name,
    activeSheetId: v2.activeSheetId,
    definedNames: v2.definedNames,
    tables: v2.tables,
    sheets: v2.sheets,
  });
  return workbook;
}

export function createWorkbookSnapshotV2(workbook: WorkbookModel): WorkbookSnapshotV2 {
  const v1 = workbook.snapshot();
  return migrateSnapshot(v1);
}
