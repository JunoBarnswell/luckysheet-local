import type {
  DrawingObject,
  DrawingTransform,
  PivotControlFilter,
  PivotControlStyle,
  PivotSlicerSettings,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  PivotTimelinePeriod,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  isPivotSlicerDrawingPayload,
  isPivotTimelineDrawingPayload,
  normalizePivotTimelinePeriod,
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
  connectedPivotIds?: string[];
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
      ...(input.connectedPivotIds?.length ? { connectedPivotIds: [...new Set(input.connectedPivotIds)] } : {}),
    },
  };
}

export function buildPivotTimelineDrawing(
  input: PivotControlDrawingInput & { period?: PivotTimelinePeriod; style?: PivotControlStyle },
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
      style: createPivotControlStyle(input.style),
      ...(input.connectedPivotIds?.length ? { connectedPivotIds: [...new Set(input.connectedPivotIds)] } : {}),
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
  return [...new Set([payload.pivotId, ...(payload.connectedPivotIds ?? [])])];
}

export function listPivotControlsForPivot(sheet: WorksheetModel, pivotId: string): PivotControlRecord[] {
  return listPivotControlRecords(sheet).filter((record) => linkedPivotIds(record.payload).includes(pivotId));
}
