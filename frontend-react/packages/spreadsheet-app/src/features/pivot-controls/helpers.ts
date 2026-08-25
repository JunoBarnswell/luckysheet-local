import type {
  DrawingObject,
  DrawingTransform,
  PivotControlFilter,
  PivotControlConnection,
  PivotControlStyle,
  PivotSlicerSettings,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  PivotTimelinePeriod,
  PivotModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  isPivotSlicerDrawingPayload,
  isPivotTimelineDrawingPayload,
  normalizePivotTimelinePeriod,
  pivotSourceIdentity,
} from '@react-sheets/core-model';

type DrawingAnchor = DrawingObject['anchor'];

export type PivotControlPayload = PivotSlicerDrawingPayload | PivotTimelineDrawingPayload;

export interface PivotControlRecord {
  drawing: DrawingObject;
  payload: PivotControlPayload;
}

export const DEFAULT_PIVOT_CONTROL_STYLE: PivotControlStyle = {
  theme: 'accent',
  fill: '#eff6ff',
  border: '#2563eb',
  textColor: '#0f172a',
  accentColor: '#2563eb',
  selectedFill: '#bfdbfe',
  fontSize: 12,
};

export const DEFAULT_PIVOT_SLICER_SETTINGS: PivotSlicerSettings = {
  showHeader: true,
  caption: 'Slicer',
  multiSelect: true,
  sort: 'ascending',
  showNoDataItems: true,
  noDataItemsLast: true,
  showNoDataStyle: true,
  columnCount: 1,
  itemHeight: 20,
};

export function createPivotSlicerSettings(overrides: Partial<PivotSlicerSettings> = {}): PivotSlicerSettings {
  const settings = { ...DEFAULT_PIVOT_SLICER_SETTINGS, ...structuredClone(overrides) };
  if (!Number.isSafeInteger(settings.columnCount) || settings.columnCount < 1 || settings.columnCount > 32) throw new Error('Pivot Slicer column count must be between 1 and 32');
  if (!Number.isFinite(settings.itemHeight) || settings.itemHeight < 16 || settings.itemHeight > 96) throw new Error('Pivot Slicer item height must be between 16 and 96');
  if (!settings.caption.trim()) throw new Error('Pivot Slicer caption must not be empty');
  return settings;
}

export function createPivotControlStyle(overrides: Partial<PivotControlStyle> = {}): PivotControlStyle {
  return { ...DEFAULT_PIVOT_CONTROL_STYLE, ...structuredClone(overrides) };
}

export function createPivotControlFilter(overrides: Partial<PivotControlFilter> = {}): PivotControlFilter {
  return {
    mode: overrides.mode ?? 'all',
    memberKeys: structuredClone(overrides.memberKeys ?? []),
  };
}

export function createPivotTimelinePeriod(overrides: PivotTimelinePeriod = {}): PivotTimelinePeriod {
  normalizePivotTimelinePeriod(overrides);
  return structuredClone(overrides);
}

export interface PivotControlDrawingInput {
  drawingId: string;
  payloadId: string;
  sheetId: string;
  pivotId: string;
  fieldId: string;
  transform: DrawingTransform;
  zIndex: number;
  anchor?: DrawingAnchor;
  name?: string;
  connections?: PivotControlConnection[];
}

export function buildPivotSlicerDrawing(
  input: PivotControlDrawingInput & { filter?: PivotControlFilter; style?: PivotControlStyle; settings?: Partial<PivotSlicerSettings> },
): { drawing: DrawingObject; payload: PivotSlicerDrawingPayload } {
  return {
    drawing: {
      id: input.drawingId,
      sheetId: input.sheetId,
      kind: 'slicer',
      name: input.name,
      anchor: structuredClone(input.anchor ?? { kind: 'absolute' }),
      transform: structuredClone(input.transform),
      zIndex: input.zIndex,
      payloadId: input.payloadId,
    },
    payload: {
      kind: 'slicer',
      pivotId: input.pivotId,
      fieldId: input.fieldId,
      filter: createPivotControlFilter(input.filter),
      style: createPivotControlStyle(input.style),
      settings: createPivotSlicerSettings(input.settings),
      ...(input.connections?.length ? { connections: structuredClone(input.connections) } : {}),
    },
  };
}

export function buildPivotTimelineDrawing(
  input: PivotControlDrawingInput & {
    period?: PivotTimelinePeriod;
    style?: PivotControlStyle;
    level?: PivotTimelineDrawingPayload['level'];
    selectionLevel?: PivotTimelineDrawingPayload['selectionLevel'];
    showHeader?: boolean;
    showSelectionLabel?: boolean;
    showTimeLevel?: boolean;
    showHorizontalScrollbar?: boolean;
    scrollPosition?: string;
    bounds?: PivotTimelinePeriod;
    filterType?: PivotTimelineDrawingPayload['filterType'];
    caption?: string;
    styleName?: string;
  },
): { drawing: DrawingObject; payload: PivotTimelineDrawingPayload } {
  return {
    drawing: {
      id: input.drawingId,
      sheetId: input.sheetId,
      kind: 'timeline',
      name: input.name,
      anchor: structuredClone(input.anchor ?? { kind: 'absolute' }),
      transform: structuredClone(input.transform),
      zIndex: input.zIndex,
      payloadId: input.payloadId,
    },
    payload: {
      kind: 'timeline',
      pivotId: input.pivotId,
      fieldId: input.fieldId,
      period: createPivotTimelinePeriod(input.period),
      level: input.level ?? 'months',
      selectionLevel: input.selectionLevel ?? input.level ?? 'months',
      showHeader: input.showHeader ?? true,
      showSelectionLabel: input.showSelectionLabel ?? true,
      showTimeLevel: input.showTimeLevel ?? true,
      showHorizontalScrollbar: input.showHorizontalScrollbar ?? true,
      ...(input.scrollPosition === undefined ? {} : { scrollPosition: input.scrollPosition }),
      bounds: createPivotTimelinePeriod(input.bounds),
      filterType: input.filterType ?? 'unknown',
      ...(input.caption === undefined ? {} : { caption: input.caption }),
      styleName: input.styleName ?? 'TimelineStyleLight2',
      style: createPivotControlStyle(input.style),
      ...(input.connections?.length ? { connections: structuredClone(input.connections) } : {}),
    },
  };
}

function readPivotControlRecord(sheet: WorksheetModel, drawing: DrawingObject): PivotControlRecord {
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload) throw new Error(`Missing Pivot control payload: ${drawing.payloadId}`);
  if (drawing.kind === 'slicer' && payload.kind === 'slicer' && isPivotSlicerDrawingPayload(payload)) {
    return { drawing, payload };
  }
  if (drawing.kind === 'timeline' && payload.kind === 'timeline' && isPivotTimelineDrawingPayload(payload)) {
    return { drawing, payload };
  }
  throw new Error(`Pivot control drawing/payload mismatch: ${drawing.id}`);
}

/** Enumerate the canonical floating controls. PivotModel arrays are not read. */
export function listPivotControlRecords(sheet: WorksheetModel): PivotControlRecord[] {
  return sheet.drawings
    .filter((drawing) => drawing.kind === 'slicer' || drawing.kind === 'timeline')
    .map((drawing) => readPivotControlRecord(sheet, drawing));
}

export function findPivotControlRecord(sheet: WorksheetModel, drawingId: string): PivotControlRecord | undefined {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  if (!drawing || (drawing.kind !== 'slicer' && drawing.kind !== 'timeline')) return undefined;
  return readPivotControlRecord(sheet, drawing);
}

export function linkedPivotIds(payload: PivotControlPayload): string[] {
  return [...new Set([payload.pivotId, ...(payload.connections ?? []).map((connection) => connection.pivotId)])];
}

export function listPivotControlsForPivot(sheet: WorksheetModel, pivotId: string): PivotControlRecord[] {
  return listPivotControlRecords(sheet).filter((record) => linkedPivotIds(record.payload).includes(pivotId));
}

function pivotById(workbook: import('@react-sheets/core-model').WorkbookModel, pivotId: string): PivotModel {
  const pivot = workbook.getSheets().flatMap((sheet) => sheet.pivots).find((candidate) => candidate.id === pivotId);
  if (!pivot) throw new Error(`Unknown PivotTable: ${pivotId}`);
  return pivot;
}

function fieldFor(pivot: PivotModel, fieldId: string) {
  const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === fieldId);
  if (!field) throw new Error(`Unknown Pivot field ${fieldId} on ${pivot.id}`);
  return field;
}

/** Return only Pivots with the same complete source/cache and source dimension. */
export function compatiblePivotControlConnections(
  workbook: import('@react-sheets/core-model').WorkbookModel,
  pivotId: string,
  fieldId: string,
  kind: 'slicer' | 'timeline',
): PivotControlConnection[] {
  const primary = pivotById(workbook, pivotId);
  const primaryField = fieldFor(primary, fieldId);
  if (kind === 'timeline' && primaryField.dataType !== 'date') throw new Error(`Timeline field is not date-semantic: ${fieldId}`);
  const sourceKey = pivotSourceIdentity(primary.source);
  return workbook.getSheets().flatMap((sheet) => sheet.pivots).filter((candidate) => candidate.id !== primary.id).flatMap((candidate) => {
    if (pivotSourceIdentity(candidate.source) !== sourceKey) return [];
    const targetField = candidate.fieldCatalog.fields.find((field) => field.ordinal === primaryField.ordinal && field.name === primaryField.name && field.dataType === primaryField.dataType);
    if (!targetField || (kind === 'timeline' && targetField.dataType !== 'date')) return [];
    return [{ pivotId: candidate.id, sourceKey, fieldId: targetField.fieldId }];
  });
}

/** Validate an explicit Report Connections set atomically against workbook state. */
export function validatePivotControlConnections(
  workbook: import('@react-sheets/core-model').WorkbookModel,
  payload: PivotControlPayload,
  connections: readonly PivotControlConnection[],
): PivotControlConnection[] {
  const primary = pivotById(workbook, payload.pivotId);
  const primaryField = fieldFor(primary, payload.fieldId);
  if (payload.kind === 'timeline' && primaryField.dataType !== 'date') throw new Error(`Timeline field is not date-semantic: ${payload.fieldId}`);
  const sourceKey = pivotSourceIdentity(primary.source);
  const seen = new Set<string>();
  return connections.map((connection) => {
    if (!connection.pivotId.trim() || connection.pivotId === payload.pivotId || seen.has(connection.pivotId)) throw new Error(`Duplicate or primary Pivot connection: ${connection.pivotId}`);
    seen.add(connection.pivotId);
    const target = pivotById(workbook, connection.pivotId);
    if (connection.sourceKey !== sourceKey || pivotSourceIdentity(target.source) !== sourceKey) throw new Error(`Pivot connection source/cache is incompatible: ${connection.pivotId}`);
    const targetField = fieldFor(target, connection.fieldId);
    if (targetField.ordinal !== primaryField.ordinal || targetField.name !== primaryField.name || targetField.dataType !== primaryField.dataType) throw new Error(`Pivot connection field is incompatible: ${connection.fieldId}`);
    if (payload.kind === 'timeline' && targetField.dataType !== 'date') throw new Error(`Timeline connection field is not date-semantic: ${connection.fieldId}`);
    return { pivotId: connection.pivotId, sourceKey, fieldId: targetField.fieldId };
  });
}
