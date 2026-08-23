import type {
  DefinedNameModel,
  RangeRef,
  SheetSnapshotV1,
  SheetId,
  WorkbookModel,
  WorkbookSnapshotV1,
  WorkbookTableModel,
  UnitId,
} from './index';
import type { PrintDocumentSnapshot, QueryDefinitionSnapshot } from './workbook-state';
import type {
  DrawingObject,
  DrawingPayload,
} from './domain';
import { WorkbookModel as WorkbookModelClass } from './index';

/**
 * The persisted snapshot format. Floating objects are represented only by
 * the canonical drawing collection and payload map. Per-kind collections are
 * intentionally absent from this type so new code cannot recreate the old
 * dual model.
 */
export interface SheetSnapshotV2 extends SheetSnapshotV1 {}

export interface WorkbookSnapshotV2 {
  schema: 'WorkbookSnapshotV2';
  schemaVersion: 2;
  unitId: UnitId;
  name: string;
  activeSheetId: SheetId;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
  tables?: WorkbookTableModel[];
  protectionRules?: import('./domain').ProtectionRule[];
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  sheets: SheetSnapshotV2[];
}

/**
 * Input-only shapes for the one-time migration of snapshots written before
 * the canonical DrawingObject/DrawingPayload model existed. These types are
 * never emitted by the core model.
 */
interface LegacyChartSnapshot {
  id: string;
  sheetId: SheetId;
  pivotId?: string;
  type: 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'combo';
  title?: string;
  sourceRanges: RangeRef[];
  series?: Array<{ name: string; range: RangeRef; color?: string }>;
  categoryRange?: RangeRef;
  bounds: { x: number; y: number; width: number; height: number };
  legendPosition?: 'top' | 'bottom' | 'left' | 'right' | 'none';
  showDataLabels?: boolean;
}

interface LegacyShapeSnapshot {
  id: string;
  sheetId: SheetId;
  type: 'rectangle' | 'rounded-rectangle' | 'ellipse' | 'line' | 'arrow' | 'callout' | 'star';
  bounds: { x: number; y: number; width: number; height: number };
  fill: string;
  stroke: string;
  strokeWidth?: number;
  text?: string;
  textColor?: string;
  fontSize?: number;
  rotation?: number;
}

interface LegacyImageSnapshot {
  id: string;
  sheetId: SheetId;
  name?: string;
  src: string;
  bounds: { x: number; y: number; width: number; height: number };
}

type LegacySheetSnapshotV1 = Omit<SheetSnapshotV1, 'drawings' | 'drawingPayloads'> & {
  drawings?: DrawingObject[];
  drawingPayloads?: Record<string, DrawingPayload>;
  charts?: LegacyChartSnapshot[];
  shapes?: LegacyShapeSnapshot[];
  images?: LegacyImageSnapshot[];
};

interface LegacyWorkbookSnapshotV1 extends Omit<WorkbookSnapshotV1, 'sheets'> {
  sheets: LegacySheetSnapshotV1[];
}

export type AnyWorkbookSnapshot = WorkbookSnapshotV1 | LegacyWorkbookSnapshotV1 | WorkbookSnapshotV2;

function absoluteAnchor() {
  return { kind: 'absolute' as const };
}

function absoluteTransform(bounds: { x: number; y: number; width: number; height: number }, rotation?: number) {
  return {
    x: bounds.x,
    y: bounds.y,
    width: bounds.width,
    height: bounds.height,
    ...(rotation === undefined ? {} : { rotation }),
  };
}

function addLegacyDrawing(
  drawings: DrawingObject[],
  payloads: Record<string, DrawingPayload>,
  drawing: DrawingObject,
  payload: DrawingPayload,
): void {
  // Canonical payloads/drawings win when a partially migrated snapshot has
  // both representations. This makes migration idempotent.
  if (drawings.some((entry) => entry.id === drawing.id)) return;
  if (!payloads[drawing.payloadId]) payloads[drawing.payloadId] = structuredClone(payload);
  drawings.push(structuredClone(drawing));
}

function migrateLegacyCollections(input: LegacySheetSnapshotV1): SheetSnapshotV1 {
  const {
    charts,
    shapes,
    images,
    drawings: sourceDrawings,
    drawingPayloads: sourcePayloads,
    ...canonicalFields
  } = structuredClone(input);
  const drawings = sourceDrawings ?? [];
  const drawingPayloads = sourcePayloads ?? {};

  for (const chart of charts ?? []) {
    addLegacyDrawing(
      drawings,
      drawingPayloads,
      {
        id: chart.id,
        sheetId: chart.sheetId,
        kind: 'chart',
        anchor: absoluteAnchor(),
        transform: absoluteTransform(chart.bounds),
        zIndex: drawings.length,
        payloadId: chart.id,
      },
      {
        kind: 'chart',
        chartId: chart.id,
        chartType: chart.type,
        title: chart.title,
        pivotId: chart.pivotId,
        sourceRanges: chart.sourceRanges,
        series: chart.series,
        categoryRange: chart.categoryRange,
        legendPosition: chart.legendPosition,
        showDataLabels: chart.showDataLabels,
      },
    );
  }

  for (const shape of shapes ?? []) {
    addLegacyDrawing(
      drawings,
      drawingPayloads,
      {
        id: shape.id,
        sheetId: shape.sheetId,
        kind: 'shape',
        anchor: absoluteAnchor(),
        transform: absoluteTransform(shape.bounds, shape.rotation),
        zIndex: drawings.length,
        payloadId: shape.id,
      },
      {
        kind: 'shape',
        type: shape.type,
        fill: shape.fill,
        stroke: shape.stroke,
        strokeWidth: shape.strokeWidth,
        text: shape.text,
        textColor: shape.textColor,
        fontSize: shape.fontSize,
      },
    );
  }

  for (const image of images ?? []) {
    addLegacyDrawing(
      drawings,
      drawingPayloads,
      {
        id: image.id,
        sheetId: image.sheetId,
        kind: 'image',
        name: image.name,
        anchor: absoluteAnchor(),
        transform: absoluteTransform(image.bounds),
        zIndex: drawings.length,
        payloadId: image.id,
      },
      {
        kind: 'image',
        src: image.src,
        name: image.name,
      },
    );
  }

  for (const drawing of drawings) {
    if (!drawingPayloads[drawing.payloadId]) {
      throw new Error(`Snapshot drawing ${drawing.id} has no payload ${drawing.payloadId}`);
    }
  }

  return {
    ...canonicalFields,
    drawings,
    drawingPayloads,
  };
}

/** Convert any supported input into the canonical V2 snapshot exactly once. */
export function migrateSnapshot(snapshot: AnyWorkbookSnapshot): WorkbookSnapshotV2 {
  if (snapshot.schema === 'WorkbookSnapshotV2') {
    return {
      schema: 'WorkbookSnapshotV2',
      schemaVersion: 2,
      unitId: snapshot.unitId,
      name: snapshot.name,
      activeSheetId: snapshot.activeSheetId,
      definedNames: snapshot.definedNames ? { ...snapshot.definedNames } : undefined,
      definedNameModels: snapshot.definedNameModels?.map((entry) => structuredClone(entry)),
      tables: snapshot.tables?.map((table) => structuredClone(table)),
      protectionRules: snapshot.protectionRules?.map((rule) => structuredClone(rule)),
      printDocuments: snapshot.printDocuments?.map((document) => structuredClone(document)),
      queryDefinitions: snapshot.queryDefinitions?.map((definition) => structuredClone(definition)),
      sheets: snapshot.sheets.map((sheet) => migrateLegacyCollections(sheet)),
    };
  }

  const legacy = snapshot as LegacyWorkbookSnapshotV1;
  return {
    schema: 'WorkbookSnapshotV2',
    schemaVersion: 2,
    unitId: legacy.unitId,
    name: legacy.name,
    activeSheetId: legacy.activeSheetId,
    definedNames: legacy.definedNames ? { ...legacy.definedNames } : undefined,
    definedNameModels: legacy.definedNameModels?.map((entry) => structuredClone(entry)),
    tables: legacy.tables?.map((table) => structuredClone(table)),
    printDocuments: legacy.printDocuments?.map((document) => structuredClone(document)),
    queryDefinitions: legacy.queryDefinitions?.map((definition) => structuredClone(definition)),
    sheets: legacy.sheets.map((sheet) => migrateLegacyCollections(sheet)),
  };
}

export function loadWorkbookFromSnapshot(snapshot: AnyWorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(snapshot);
}

export function createWorkbookSnapshotV2(workbook: WorkbookModel): WorkbookSnapshotV2 {
  return migrateSnapshot(workbook.snapshot());
}
