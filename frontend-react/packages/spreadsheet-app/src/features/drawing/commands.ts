import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type {
  DrawingObject,
  DrawingPayload,
  ConnectorDrawingPayload,
  DrawingGroup,
  DrawingTransform,
  ImageCrop,
  ImageDrawingPayload,
  ImageEffects,
  TextBoxDrawingPayload,
  TextBoxTextFrame,
  WorksheetModel,
  WorksheetSnapSettings,
} from '@react-sheets/core-model';
import {
  assertCanonicalConnector,
  isAssetRef,
  isDrawingConnectorPayload,
  isShapeDrawingPayload,
  isDrawingGroup,
  isFormControlDrawingPayload,
  isPivotSlicerDrawingPayload,
  isPivotTimelineDrawingPayload,
  isWorksheetSnapSettings,
  planConnectorRoute,
  recomputeConnectorRoutes,
} from '@react-sheets/core-model';
import { DrawingRuntime, reorderDrawing } from './runtime';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function removeById<T extends { id: string }>(items: T[], id: string): T | undefined {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return undefined;
  return items.splice(index, 1)[0];
}

export interface DrawingSelectParams {
  sheetId: string;
  drawingIds: string[];
  mode?: 'replace' | 'add' | 'toggle';
}

export interface DrawingTransformParams {
  sheetId: string;
  drawingId: string;
  transform: DrawingTransform;
  connectorRoutes?: DrawingConnectorRouteChange[];
}

/** One pointer gesture: preview frames are transient and only this commit is persisted. */
export interface DrawingTransformCommitParams {
  sheetId: string;
  drawingId: string;
  before: DrawingTransform;
  after: DrawingTransform;
}

export interface DrawingTransformBatchEntry {
  drawingId: string;
  before: DrawingTransform;
  after: DrawingTransform;
}

export interface DrawingTransformBatchParams {
  sheetId: string;
  entries: DrawingTransformBatchEntry[];
  connectorRoutes?: DrawingConnectorRouteChange[];
}

export interface DrawingConnectorRouteState {
  payload: ConnectorDrawingPayload;
  transform: DrawingTransform;
}

export interface DrawingConnectorRouteChange {
  drawingId: string;
  before: DrawingConnectorRouteState;
  after: DrawingConnectorRouteState;
}

export interface DrawingConnectorAddParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: ConnectorDrawingPayload;
}

export interface DrawingConnectorUpdateParams {
  sheetId: string;
  drawingId: string;
  before: ConnectorDrawingPayload;
  after: ConnectorDrawingPayload;
}

export interface DrawingGroupParams {
  sheetId: string;
  group: DrawingGroup;
}

export interface DrawingUngroupParams {
  sheetId: string;
  groupId: string;
}

export interface DrawingSnapSettingsParams {
  sheetId: string;
  before: WorksheetSnapSettings;
  after: WorksheetSnapSettings;
}

export interface DrawingAnchorParams {
  sheetId: string;
  drawingId: string;
  anchor: DrawingObject['anchor'];
}

export interface DrawingAddParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: DrawingPayload;
}

export interface DrawingRemoveParams {
  sheetId: string;
  drawingId: string;
}

export interface DrawingPayloadUpdateParams {
  sheetId: string;
  payloadId: string;
  before: DrawingPayload;
  after: DrawingPayload;
}

export interface DrawingTextBoxUpdateParams {
  sheetId: string;
  drawingId: string;
  payload: TextBoxDrawingPayload;
}

export interface DrawingZOrderParams {
  sheetId: string;
  drawingId: string;
  direction: 'forward' | 'backward' | 'front' | 'back';
}

interface DrawingZOrderRestoreParams {
  sheetId: string;
  entries: Array<{ drawingId: string; zIndex: number }>;
}

export interface DrawingAlignParams {
  sheetId: string;
  drawingIds: string[];
  alignment: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';
}

export interface DrawingDistributeParams {
  sheetId: string;
  drawingIds: string[];
  axis: 'horizontal' | 'vertical';
}

export interface DrawingCopyParams {
  sheetId: string;
  sourceDrawingId: string;
  drawingId: string;
  payloadId: string;
  offset: { x: number; y: number };
}

export type DrawingImageCrop = ImageCrop;

export interface DrawingImageCropParams {
  sheetId: string;
  drawingId: string;
  crop: DrawingImageCrop;
}

export interface DrawingImageAltTextParams {
  sheetId: string;
  drawingId: string;
  altText: string;
}

export interface DrawingImageEffectsParams {
  sheetId: string;
  drawingId: string;
  effects: ImageEffects;
}

const objectParams = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const hasSheetId = (value: unknown): value is { sheetId: string } => objectParams(value) && typeof value.sheetId === 'string' && value.sheetId.length > 0;
const hasDrawingId = (value: unknown): value is { drawingId: string } => objectParams(value) && typeof value.drawingId === 'string' && value.drawingId.length > 0;
const hasPayloadId = (value: unknown): value is { payloadId: string } => objectParams(value) && typeof value.payloadId === 'string' && value.payloadId.length > 0;
const isFiniteNumber = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value);

function isTransform(value: unknown): value is DrawingTransform {
  if (!objectParams(value)) return false;
  return isFiniteNumber(value.x)
    && isFiniteNumber(value.y)
    && isFiniteNumber(value.width)
    && isFiniteNumber(value.height)
    && value.width >= 0
    && value.height >= 0
    && (value.rotation === undefined || isFiniteNumber(value.rotation));
}

function isAnchor(value: unknown): value is DrawingObject['anchor'] {
  if (!objectParams(value) || !['absolute', 'one-cell', 'two-cell'].includes(String(value.kind))) return false;
  const anchor = value as Record<string, unknown>;
  if (anchor.kind === 'absolute') return true;
  if (!Number.isInteger(anchor.row) || !Number.isInteger(anchor.column) || (anchor.row as number) < 0 || (anchor.column as number) < 0) return false;
  if (anchor.kind === 'two-cell') {
    return Number.isInteger(anchor.endRow) && Number.isInteger(anchor.endColumn)
      && (anchor.endRow as number) >= (anchor.row as number) && (anchor.endColumn as number) >= (anchor.column as number);
  }
  return true;
}

function isDrawingPayload(value: unknown): value is DrawingPayload {
  if (!objectParams(value)) return false;
  if (value.kind === 'image') return isImagePayload(value);
  if (value.kind === 'connector') return isDrawingConnectorPayload(value);
  if (value.kind === 'form-control') return isFormControlDrawingPayload(value);
  if (value.kind === 'textbox') return isTextBoxPayload(value);
  if (value.kind === 'shape') return isShapeDrawingPayload(value);
  if (['chart', 'data-chart', 'camera'].includes(String(value.kind))) return true;
  return isPivotSlicerDrawingPayload(value) || isPivotTimelineDrawingPayload(value);
}

function isTextBoxTextFrame(value: unknown): value is TextBoxTextFrame {
  if (!objectParams(value)) return false;
  const margin = value.margin;
  return typeof value.fontFamily === 'string' && value.fontFamily.length > 0
    && isFiniteNumber(value.fontSize) && value.fontSize > 0
    && typeof value.bold === 'boolean' && typeof value.italic === 'boolean' && typeof value.underline === 'boolean'
    && typeof value.textColor === 'string' && value.textColor.length > 0
    && ['left', 'center', 'right'].includes(String(value.horizontalAlignment))
    && ['top', 'middle', 'bottom'].includes(String(value.verticalAlignment))
    && ['horizontal', 'vertical'].includes(String(value.direction))
    && objectParams(margin)
    && isFiniteNumber(margin.top) && margin.top >= 0
    && isFiniteNumber(margin.right) && margin.right >= 0
    && isFiniteNumber(margin.bottom) && margin.bottom >= 0
    && isFiniteNumber(margin.left) && margin.left >= 0
    && typeof value.wrap === 'boolean'
    && ['none', 'shrink-text', 'resize-shape'].includes(String(value.autofit));
}

function isTextBoxPayload(value: unknown): value is TextBoxDrawingPayload {
  return objectParams(value) && value.kind === 'textbox' && typeof value.text === 'string' && isTextBoxTextFrame(value.textFrame);
}

function isDrawing(value: unknown): value is DrawingObject {
  if (!objectParams(value)) return false;
  return typeof value.id === 'string'
    && typeof value.sheetId === 'string'
    && typeof value.kind === 'string'
    && typeof value.payloadId === 'string'
    && isAnchor(value.anchor)
    && isTransform(value.transform)
    && isFiniteNumber(value.zIndex);
}

function isDrawingAddParams(value: unknown): value is DrawingAddParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isDrawing(params.drawing) && isDrawingPayload(params.payload);
}

function isConnectorRouteState(value: unknown): value is DrawingConnectorRouteState {
  if (!objectParams(value)) return false;
  return isDrawingConnectorPayload(value.payload) && isTransform(value.transform);
}

function isConnectorRouteChange(value: unknown): value is DrawingConnectorRouteChange {
  if (!objectParams(value)) return false;
  return typeof value.drawingId === 'string' && value.drawingId.length > 0
    && isConnectorRouteState(value.before) && isConnectorRouteState(value.after);
}

function isConnectorRoutes(value: unknown): value is DrawingConnectorRouteChange[] | undefined {
  return value === undefined || (Array.isArray(value) && value.every(isConnectorRouteChange));
}

function isConnectorAddParams(value: unknown): value is DrawingConnectorAddParams {
  if (!isDrawingAddParams(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isDrawing(params.drawing) && params.drawing.kind === 'connector' && isDrawingConnectorPayload(params.payload);
}

function isConnectorUpdateParams(value: unknown): value is DrawingConnectorUpdateParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isDrawingConnectorPayload(params.before) && isDrawingConnectorPayload(params.after);
}

function isGroupParams(value: unknown): value is DrawingGroupParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  return isDrawingGroup((value as Record<string, unknown>).group);
}

function isUngroupParams(value: unknown): value is DrawingUngroupParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  const groupId = (value as Record<string, unknown>).groupId;
  return typeof groupId === 'string' && groupId.trim().length > 0;
}

function isSnapSettingsParams(value: unknown): value is DrawingSnapSettingsParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isWorksheetSnapSettings(params.before) && isWorksheetSnapSettings(params.after);
}

function isDrawingRemoveParams(value: unknown): value is DrawingRemoveParams {
  return hasSheetId(value) && hasDrawingId(value);
}

function isTransformParams(value: unknown): value is DrawingTransformParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isTransform(params.transform) && isConnectorRoutes(params.connectorRoutes);
}

function isTransformCommitParams(value: unknown): value is DrawingTransformCommitParams {
  if (!objectParams(value) || !hasSheetId(value) || !hasDrawingId(value)) return false;
  const params = value as Record<string, unknown>;
  return isTransform(params.before) && isTransform(params.after);
}

function isTransformBatchParams(value: unknown): value is DrawingTransformBatchParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  const entries = (value as Record<string, unknown>).entries;
  return Array.isArray(entries) && entries.every((entry) => {
    if (!objectParams(entry)) return false;
    const item = entry as Record<string, unknown>;
    return typeof item.drawingId === 'string' && isTransform(item.before) && isTransform(item.after);
  }) && isConnectorRoutes((value as Record<string, unknown>).connectorRoutes);
}

function isAnchorParams(value: unknown): value is DrawingAnchorParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return isAnchor((value as Record<string, unknown>).anchor);
}

function isPayloadUpdateParams(value: unknown): value is DrawingPayloadUpdateParams {
  if (!hasSheetId(value) || !hasPayloadId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isDrawingPayload(params.before) && isDrawingPayload(params.after);
}

function isTextBoxUpdateParams(value: unknown): value is DrawingTextBoxUpdateParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  const params = value as Record<string, unknown>;
  return isTextBoxPayload(params.payload);
}

function isZOrderParams(value: unknown): value is DrawingZOrderParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return ['forward', 'backward', 'front', 'back'].includes(String((value as Record<string, unknown>).direction));
}

function isZOrderRestoreParams(value: unknown): value is DrawingZOrderRestoreParams {
  if (!hasSheetId(value) || !objectParams(value)) return false;
  const entries = (value as Record<string, unknown>).entries;
  return Array.isArray(entries) && entries.every((entry) => {
    if (!objectParams(entry)) return false;
    const item = entry as Record<string, unknown>;
    return typeof item.drawingId === 'string' && isFiniteNumber(item.zIndex);
  });
}

function isImageCrop(value: unknown): value is DrawingImageCrop {
  return objectParams(value)
    && isFiniteNumber(value.left) && isFiniteNumber(value.top)
    && isFiniteNumber(value.right) && isFiniteNumber(value.bottom)
    && value.left >= 0 && value.top >= 0 && value.right >= 0 && value.bottom >= 0
    && value.left + value.right < 1 && value.top + value.bottom < 1;
}

function isImageEffects(value: unknown): value is ImageEffects {
  if (!objectParams(value)) return false;
  return (value.brightness === undefined || (isFiniteNumber(value.brightness) && value.brightness >= -1 && value.brightness <= 1))
    && (value.contrast === undefined || (isFiniteNumber(value.contrast) && value.contrast >= -1 && value.contrast <= 1))
    && (value.transparency === undefined || (isFiniteNumber(value.transparency) && value.transparency >= 0 && value.transparency <= 1));
}

function isImagePayload(value: unknown): value is ImageDrawingPayload {
  return objectParams(value)
    && value.kind === 'image'
    && isAssetRef(value.asset)
    && (value.altText === undefined || typeof value.altText === 'string')
    && (value.name === undefined || typeof value.name === 'string')
    && (value.crop === undefined || isImageCrop(value.crop))
    && (value.effects === undefined || isImageEffects(value.effects));
}

function isImageCropParams(value: unknown): value is DrawingImageCropParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return isImageCrop((value as Record<string, unknown>).crop);
}

function isImageAltTextParams(value: unknown): value is DrawingImageAltTextParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return typeof (value as Record<string, unknown>).altText === 'string';
}

function isImageEffectsParams(value: unknown): value is DrawingImageEffectsParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return isImageEffects((value as Record<string, unknown>).effects);
}

function isDrawingPairValid(drawing: DrawingObject, payload: DrawingPayload): void {
  if (!drawing.id || !drawing.sheetId || !drawing.payloadId) throw new Error(`Drawing identity is required: ${drawing.id}`);
  if (!isAnchor(drawing.anchor)) throw new Error(`Invalid drawing anchor: ${drawing.id}`);
  if (!isTransform(drawing.transform)) throw new Error(`Invalid drawing transform: ${drawing.id}`);
  if (!isDrawingPayload(payload)) throw new Error(`Invalid drawing payload: ${drawing.payloadId}`);
  if (drawing.kind !== payload.kind) throw new Error(`Drawing payload kind mismatch: ${drawing.id}`);
  if (payload.kind === 'chart' && payload.chartId !== drawing.payloadId) {
    throw new Error(`Drawing payload identity mismatch: ${drawing.payloadId}`);
  }
}

function canonicalConnectorAggregate(sheet: WorksheetModel, drawing: DrawingObject, payload: ConnectorDrawingPayload): { drawing: DrawingObject; payload: ConnectorDrawingPayload } {
  const planned = planConnectorRoute(sheet, drawing, payload);
  return { drawing: { ...structuredClone(drawing), transform: planned.transform }, payload: planned.payload };
}

function addDrawing(sheet: WorksheetModel, drawing: DrawingObject, payload: DrawingPayload): void {
  if (payload.kind === 'connector') {
    const canonical = canonicalConnectorAggregate(sheet, drawing, payload);
    if (JSON.stringify(canonical.drawing.transform) !== JSON.stringify(drawing.transform)
      || JSON.stringify(canonical.payload.route) !== JSON.stringify(payload.route)) {
      throw new Error(`Connector aggregate is not canonical: ${drawing.id}`);
    }
  }
  isDrawingPairValid(drawing, payload);
  if (drawing.sheetId !== sheet.id) throw new Error(`Drawing sheet mismatch: ${drawing.id}`);
  if (sheet.drawings.some((entry) => entry.id === drawing.id)) throw new Error(`Drawing already exists: ${drawing.id}`);
  if (sheet.drawingPayloads.has(drawing.payloadId)) throw new Error(`Drawing payload already exists: ${drawing.payloadId}`);
  sheet.drawings.push(structuredClone(drawing));
  sheet.drawingPayloads.set(drawing.payloadId, structuredClone(payload));
}

function addGroup(sheet: WorksheetModel, group: DrawingGroup): void {
  if (!isDrawingGroup(group) || group.sheetId !== sheet.id) throw new Error(`Drawing group is invalid: ${group.id}`);
  if (sheet.drawingGroups.some((entry) => entry.id === group.id)) throw new Error(`Drawing group already exists: ${group.id}`);
  const ids = new Set(group.memberDrawingIds);
  if (ids.size !== group.memberDrawingIds.length) throw new Error(`Drawing group contains duplicate members: ${group.id}`);
  for (const id of ids) {
    findDrawing(sheet, id);
    if (sheet.drawingGroups.some((entry) => entry.memberDrawingIds.includes(id))) throw new Error(`Drawing belongs to another group: ${id}`);
  }
  sheet.drawingGroups.push(structuredClone(group));
}

function removeGroup(sheet: WorksheetModel, groupId: string): DrawingGroup {
  const index = sheet.drawingGroups.findIndex((entry) => entry.id === groupId);
  if (index < 0) throw new Error(`Unknown drawing group: ${groupId}`);
  return structuredClone(sheet.drawingGroups.splice(index, 1)[0]!);
}

function applyConnectorRouteState(sheet: WorksheetModel, drawingId: string, state: DrawingConnectorRouteState, expectedBefore?: DrawingConnectorRouteState): void {
  const drawing = findDrawing(sheet, drawingId);
  if (drawing.kind !== 'connector') throw new Error(`Drawing is not a connector: ${drawingId}`);
  const current = sheet.drawingPayloads.get(drawing.payloadId);
  if (!current || current.kind !== 'connector') throw new Error(`Missing connector payload: ${drawing.payloadId}`);
  if (expectedBefore && (JSON.stringify(current) !== JSON.stringify(expectedBefore.payload) || JSON.stringify(drawing.transform) !== JSON.stringify(expectedBefore.transform))) {
    throw new Error(`Connector route changed before update: ${drawingId}`);
  }
  const canonical = canonicalConnectorAggregate(sheet, drawing, state.payload);
  if (JSON.stringify(canonical.drawing.transform) !== JSON.stringify(state.transform)
    || JSON.stringify(canonical.payload) !== JSON.stringify(state.payload)) throw new Error(`Connector route is not canonical: ${drawingId}`);
  sheet.drawingPayloads.set(drawing.payloadId, structuredClone(state.payload));
  drawing.transform = structuredClone(state.transform);
}

function applyTransformMutation(sheet: WorksheetModel, entries: readonly DrawingTransformBatchEntry[], connectorRoutes: readonly DrawingConnectorRouteChange[] | undefined): void {
  const overrides = entries.map((entry) => ({ drawingId: entry.drawingId, transform: entry.after }));
  const expectedRoutes = recomputeConnectorRoutes(sheet, overrides);
  const expectedById = new Map(expectedRoutes.map((entry) => [entry.drawingId, entry]));
  const provided = connectorRoutes ?? [];
  if (provided.length !== expectedRoutes.length || provided.some((entry) => {
    const expected = expectedById.get(entry.drawingId);
    return !expected || JSON.stringify(entry.after) !== JSON.stringify(expected.after) || JSON.stringify(entry.before) !== JSON.stringify(expected.before);
  })) throw new Error('Connector route changes do not match the canonical transform plan');
  for (const entry of entries) {
    const drawing = findDrawing(sheet, entry.drawingId);
    if (JSON.stringify(drawing.transform) !== JSON.stringify(entry.before)) throw new Error(`Drawing transform changed before update: ${entry.drawingId}`);
  }
  for (const route of provided) {
    const drawing = findDrawing(sheet, route.drawingId);
    const current = sheet.drawingPayloads.get(drawing.payloadId);
    if (!current || current.kind !== 'connector' || JSON.stringify(current) !== JSON.stringify(route.before.payload)
      || JSON.stringify(drawing.transform) !== JSON.stringify(route.before.transform)) throw new Error(`Connector route changed before transform: ${route.drawingId}`);
  }
  for (const entry of entries) findDrawing(sheet, entry.drawingId).transform = structuredClone(entry.after);
  for (const route of provided) applyConnectorRouteState(sheet, route.drawingId, route.after);
}

function removeDrawing(sheet: WorksheetModel, drawingId: string): { drawing: DrawingObject; payload: DrawingPayload } {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  if (!drawing) throw new Error(`Unknown drawing: ${drawingId}`);
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload) throw new Error(`Missing drawing payload: ${drawing.payloadId}`);
  removeById(sheet.drawings, drawingId);
  sheet.drawingPayloads.delete(drawing.payloadId);
  return { drawing: structuredClone(drawing), payload: structuredClone(payload) };
}

function findDrawing(sheet: WorksheetModel, drawingId: string): DrawingObject {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  if (!drawing) throw new Error(`Unknown drawing: ${drawingId}`);
  return drawing;
}

function updatePayload(sheet: WorksheetModel, params: DrawingPayloadUpdateParams): void {
  const drawing = sheet.drawings.find((entry) => entry.payloadId === params.payloadId);
  if (!drawing) throw new Error(`Unknown drawing payload: ${params.payloadId}`);
  isDrawingPairValid(drawing, params.after);
  const current = sheet.drawingPayloads.get(params.payloadId);
  if (!current) throw new Error(`Missing drawing payload: ${params.payloadId}`);
  if (current.kind !== params.before.kind) throw new Error(`Drawing payload kind mismatch: ${params.payloadId}`);
  if (JSON.stringify(current) !== JSON.stringify(params.before)) throw new Error(`Drawing payload changed before update: ${params.payloadId}`);
  sheet.drawingPayloads.set(params.payloadId, structuredClone(params.after));
}

function restoreZOrder(sheet: WorksheetModel, params: DrawingZOrderRestoreParams): void {
  for (const entry of params.entries) findDrawing(sheet, entry.drawingId).zIndex = entry.zIndex;
}

function rangesForParams(params: unknown): ReturnType<typeof sheetRange> {
  return hasSheetId(params) ? sheetRange(params.sheetId) : [];
}

function executeConnectorAdd(params: DrawingConnectorAddParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const canonical = canonicalConnectorAggregate(sheet, params.drawing, params.payload);
  const next: DrawingConnectorAddParams = { sheetId: params.sheetId, drawing: canonical.drawing, payload: canonical.payload };
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.connector.add',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: next,
    affectedRanges,
    inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: canonical.drawing.id }, affectedRanges }],
    apply: () => addDrawing(context.workbook.getSheet(params.sheetId), canonical.drawing, canonical.payload),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function connectorRouteChanges(sheet: WorksheetModel, entries: readonly DrawingTransformBatchEntry[]): DrawingConnectorRouteChange[] {
  return recomputeConnectorRoutes(sheet, entries.map((entry) => ({ drawingId: entry.drawingId, transform: entry.after }))).map((entry) => ({
    drawingId: entry.drawingId,
    before: entry.before,
    after: entry.after,
  }));
}

function assertProvidedRoutes(expected: readonly DrawingConnectorRouteChange[], provided: readonly DrawingConnectorRouteChange[] | undefined): void {
  if (provided !== undefined && JSON.stringify(provided) !== JSON.stringify(expected)) throw new Error('Connector route changes do not match the canonical transform plan');
}

function executeAdd(params: DrawingAddParams, context: CommandContext, kind?: DrawingObject['kind']): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const drawing = kind ? { ...params.drawing, kind } : params.drawing;
  if (drawing.kind === 'connector' && params.payload.kind === 'connector') return executeConnectorAdd({ sheetId: params.sheetId, drawing, payload: params.payload }, context);
  const next = { sheetId: params.sheetId, drawing, payload: params.payload };
  isDrawingPairValid(next.drawing, next.payload);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.add',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: next,
    affectedRanges,
    inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: drawing.id }, affectedRanges }],
    apply: () => addDrawing(context.workbook.getSheet(params.sheetId), drawing, params.payload),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeTransform(params: DrawingTransformParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const drawing = findDrawing(sheet, params.drawingId);
  const previous = structuredClone(drawing.transform);
  const entries = [{ drawingId: params.drawingId, before: previous, after: structuredClone(params.transform) }];
  const connectorRoutes = connectorRouteChanges(sheet, entries);
  assertProvidedRoutes(connectorRoutes, params.connectorRoutes);
  const canonicalParams: DrawingTransformParams = { ...params, connectorRoutes };
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.transform',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: canonicalParams,
    affectedRanges,
    inverse: [{ id: 'drawing.transform', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...canonicalParams, transform: previous, connectorRoutes: connectorRoutes.map((entry) => ({ drawingId: entry.drawingId, before: entry.after, after: entry.before })) }, affectedRanges }],
    apply: () => applyTransformMutation(context.workbook.getSheet(params.sheetId), entries, connectorRoutes),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeTransformCommit(params: DrawingTransformCommitParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const drawing = findDrawing(sheet, params.drawingId);
  if (JSON.stringify(drawing.transform) !== JSON.stringify(params.before)) throw new Error(`Drawing transform changed before pointer commit: ${params.drawingId}`);
  const entries = [{ drawingId: params.drawingId, before: structuredClone(params.before), after: structuredClone(params.after) }];
  const connectorRoutes = connectorRouteChanges(sheet, entries);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.transform',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: { sheetId: params.sheetId, drawingId: params.drawingId, transform: structuredClone(params.after), connectorRoutes },
    affectedRanges,
    inverse: [{ id: 'drawing.transform', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawingId, transform: structuredClone(params.before), connectorRoutes: connectorRoutes.map((entry) => ({ drawingId: entry.drawingId, before: entry.after, after: entry.before })) }, affectedRanges }],
    apply: () => applyTransformMutation(sheet, entries, connectorRoutes),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeTransformBatch(params: DrawingTransformBatchParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  if (params.entries.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
  const sheet = context.workbook.getSheet(params.sheetId);
  for (const entry of params.entries) {
    const drawing = findDrawing(sheet, entry.drawingId);
    if (JSON.stringify(drawing.transform) !== JSON.stringify(entry.before)) throw new Error(`Drawing transform changed before batch: ${entry.drawingId}`);
  }
  const connectorRoutes = connectorRouteChanges(sheet, params.entries);
  assertProvidedRoutes(connectorRoutes, params.connectorRoutes);
  const canonicalParams: DrawingTransformBatchParams = { ...params, connectorRoutes };
  const affectedRanges = sheetRange(params.sheetId);
  const inverse: DrawingTransformBatchParams = {
    sheetId: params.sheetId,
    entries: params.entries.map((entry) => ({ drawingId: entry.drawingId, before: entry.after, after: entry.before })),
    connectorRoutes: connectorRoutes.map((entry) => ({ drawingId: entry.drawingId, before: entry.after, after: entry.before })),
  };
  context.applyMutation({
    id: 'drawing.transform.batch',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: canonicalParams,
    affectedRanges,
    inverse: [{ id: 'drawing.transform.batch', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges }],
    apply: () => {
      applyTransformMutation(context.workbook.getSheet(params.sheetId), params.entries, connectorRoutes);
    },
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function prepareConnectorUpdate(params: DrawingConnectorUpdateParams, context: CommandContext): { canonicalParams: DrawingConnectorUpdateParams; current: ConnectorDrawingPayload; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const drawing = findDrawing(sheet, params.drawingId);
  if (drawing.kind !== 'connector') throw new Error(`Drawing is not a connector: ${params.drawingId}`);
  const current = sheet.drawingPayloads.get(drawing.payloadId);
  if (!current || current.kind !== 'connector' || JSON.stringify(current) !== JSON.stringify(params.before)) throw new Error(`Connector payload changed before update: ${params.drawingId}`);
  const canonical = canonicalConnectorAggregate(sheet, drawing, params.after);
  const canonicalParams: DrawingConnectorUpdateParams = { ...params, after: canonical.payload, before: structuredClone(current) };
  return { canonicalParams, current, affectedRanges: sheetRange(params.sheetId) };
}

function applyPreparedConnectorUpdate(params: DrawingConnectorUpdateParams, context: CommandContext): void {
  const target = context.workbook.getSheet(params.sheetId);
  const targetDrawing = findDrawing(target, params.drawingId);
  const targetPayload = target.drawingPayloads.get(targetDrawing.payloadId);
  if (!targetPayload || targetDrawing.kind !== 'connector' || targetPayload.kind !== 'connector' || JSON.stringify(targetPayload) !== JSON.stringify(params.before)) throw new Error(`Connector payload changed before update: ${params.drawingId}`);
  const planned = canonicalConnectorAggregate(target, targetDrawing, params.after);
  if (JSON.stringify(planned.payload) !== JSON.stringify(params.after)) throw new Error(`Connector route is not canonical: ${params.drawingId}`);
  target.drawingPayloads.set(targetDrawing.payloadId, structuredClone(params.after));
  targetDrawing.transform = structuredClone(planned.drawing.transform);
}

function executeConnectorUpdate(params: DrawingConnectorUpdateParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const prepared = prepareConnectorUpdate(params, context);
  const canonicalParams = prepared.canonicalParams;
  const affectedRanges = prepared.affectedRanges;
  context.applyMutation({
    id: 'drawing.connector.update',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: canonicalParams,
    affectedRanges,
    inverse: [{ id: 'drawing.connector.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawingId, before: canonicalParams.after, after: canonicalParams.before }, affectedRanges }],
    apply: () => applyPreparedConnectorUpdate(canonicalParams, context),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeConnectorRoute(params: DrawingConnectorUpdateParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const prepared = prepareConnectorUpdate(params, context);
  const canonicalParams = prepared.canonicalParams;
  const affectedRanges = prepared.affectedRanges;
  context.applyMutation({
    id: 'drawing.connector.route',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: canonicalParams,
    affectedRanges,
    inverse: [{ id: 'drawing.connector.route', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawingId, before: canonicalParams.after, after: canonicalParams.before }, affectedRanges }],
    apply: () => applyPreparedConnectorUpdate(canonicalParams, context),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeGroup(params: DrawingGroupParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const group = structuredClone(params.group);
  addGroup(sheet, group);
  removeGroup(sheet, group.id);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.group',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: { sheetId: params.sheetId, group },
    affectedRanges,
    inverse: [{ id: 'drawing.ungroup', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, groupId: group.id }, affectedRanges }],
    apply: () => addGroup(context.workbook.getSheet(params.sheetId), group),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeUngroup(params: DrawingUngroupParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const group = structuredClone(sheet.drawingGroups.find((entry) => entry.id === params.groupId));
  if (!group) throw new Error(`Unknown drawing group: ${params.groupId}`);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.ungroup',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{ id: 'drawing.group', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, group }, affectedRanges }],
    apply: () => removeGroup(context.workbook.getSheet(params.sheetId), params.groupId),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeSnapSettings(params: DrawingSnapSettingsParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  if (JSON.stringify(sheet.snapSettings) !== JSON.stringify(params.before)) throw new Error('Worksheet snap settings changed before update');
  if (!isWorksheetSnapSettings(params.after)) throw new Error('Worksheet snap settings are invalid');
  const canonical: DrawingSnapSettingsParams = { sheetId: params.sheetId, before: structuredClone(sheet.snapSettings), after: structuredClone(params.after) };
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.snapSettings',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: canonical,
    affectedRanges,
    inverse: [{ id: 'drawing.snapSettings', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, before: canonical.after, after: canonical.before }, affectedRanges }],
    apply: () => { context.workbook.getSheet(params.sheetId).snapSettings = structuredClone(canonical.after); },
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function handleConnectorUpdate(item: { params: DrawingConnectorUpdateParams }, context: CommandContext): void {
  const sheet = context.workbook.getSheet(item.params.sheetId);
  const drawing = findDrawing(sheet, item.params.drawingId);
  const current = sheet.drawingPayloads.get(drawing.payloadId);
  if (drawing.kind !== 'connector' || !current || current.kind !== 'connector' || JSON.stringify(current) !== JSON.stringify(item.params.before)) throw new Error(`Connector payload changed before update: ${item.params.drawingId}`);
  const canonical = canonicalConnectorAggregate(sheet, drawing, item.params.after);
  if (JSON.stringify(canonical.payload) !== JSON.stringify(item.params.after)) throw new Error(`Connector route is not canonical: ${item.params.drawingId}`);
  sheet.drawingPayloads.set(drawing.payloadId, structuredClone(item.params.after));
  drawing.transform = structuredClone(canonical.drawing.transform);
}

export function registerDrawingCommands(runtime: CommandRuntime, drawingRuntime: DrawingRuntime): string[] {
  runtime.registry.registerMutation<DrawingAddParams>({
      id: 'drawing.add',
      handler: (item, context) => {
    addDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawing, item.params.payload);
  },
      metadata: {
    schema: { name: 'DrawingAddParams', validate: isDrawingAddParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingConnectorAddParams>({
    id: 'drawing.connector.add',
    handler: (item, context) => addDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawing, item.params.payload),
    metadata: {
      schema: { name: 'DrawingConnectorAddParams', validate: isConnectorAddParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.remove'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingRemoveParams>({
      id: 'drawing.remove',
      handler: (item, context) => {
    removeDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId);
  },
      metadata: {
    schema: { name: 'DrawingRemoveParams', validate: isDrawingRemoveParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.add'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingTransformParams>({
      id: 'drawing.transform',
      handler: (item, context) => {
    applyTransformMutation(context.workbook.getSheet(item.params.sheetId), [{ drawingId: item.params.drawingId, before: findDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId).transform, after: item.params.transform }], item.params.connectorRoutes);
  },
      metadata: {
    schema: { name: 'DrawingTransformParams', validate: isTransformParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.transform'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingTransformBatchParams>({
      id: 'drawing.transform.batch',
      handler: (item, context) => {
    applyTransformMutation(context.workbook.getSheet(item.params.sheetId), item.params.entries, item.params.connectorRoutes);
  },
      metadata: {
    schema: { name: 'DrawingTransformBatchParams', validate: isTransformBatchParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.transform.batch'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingAnchorParams>({
      id: 'drawing.anchor',
      handler: (item, context) => {
    findDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId).anchor = structuredClone(item.params.anchor);
  },
      metadata: {
    schema: { name: 'DrawingAnchorParams', validate: isAnchorParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.anchor'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingPayloadUpdateParams>({
      id: 'drawing.payload.update',
      handler: (item, context) => {
    updatePayload(context.workbook.getSheet(item.params.sheetId), item.params);
  },
      metadata: {
    schema: { name: 'DrawingPayloadUpdateParams', validate: isPayloadUpdateParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.payload.update'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingConnectorUpdateParams>({
    id: 'drawing.connector.update',
    handler: (item, context) => handleConnectorUpdate(item, context),
    metadata: {
      schema: { name: 'DrawingConnectorUpdateParams', validate: isConnectorUpdateParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.connector.update'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingConnectorUpdateParams>({
    id: 'drawing.connector.route',
    handler: (item, context) => handleConnectorUpdate(item, context),
    metadata: {
      schema: { name: 'DrawingConnectorRouteParams', validate: isConnectorUpdateParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.connector.route'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingGroupParams>({
    id: 'drawing.group',
    handler: (item, context) => addGroup(context.workbook.getSheet(item.params.sheetId), item.params.group),
    metadata: {
      schema: { name: 'DrawingGroupParams', validate: isGroupParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.ungroup'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingUngroupParams>({
    id: 'drawing.ungroup',
    handler: (item, context) => removeGroup(context.workbook.getSheet(item.params.sheetId), item.params.groupId),
    metadata: {
      schema: { name: 'DrawingUngroupParams', validate: isUngroupParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.group'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingSnapSettingsParams>({
    id: 'drawing.snapSettings',
    handler: (item, context) => {
      const sheet = context.workbook.getSheet(item.params.sheetId);
      if (JSON.stringify(sheet.snapSettings) !== JSON.stringify(item.params.before)) throw new Error('Worksheet snap settings changed before update');
      if (!isWorksheetSnapSettings(item.params.after)) throw new Error('Worksheet snap settings are invalid');
      sheet.snapSettings = structuredClone(item.params.after);
    },
    metadata: {
      schema: { name: 'DrawingSnapSettingsParams', validate: isSnapSettingsParams },
      permission: { capability: 'drawing.edit' },
      affectedRanges: { resolve: rangesForParams, mode: 'declared' },
      inversePolicy: { allowedMutationIds: ['drawing.snapSettings'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<DrawingZOrderParams>({
      id: 'drawing.zorder',
      handler: (item, context) => {
    reorderDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId, item.params.direction);
  },
      metadata: {
    schema: { name: 'DrawingZOrderParams', validate: isZOrderParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.zorder.restore'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<DrawingZOrderRestoreParams>({
      id: 'drawing.zorder.restore',
      handler: (item, context) => {
    restoreZOrder(context.workbook.getSheet(item.params.sheetId), item.params);
  },
      metadata: {
    schema: { name: 'DrawingZOrderRestoreParams', validate: isZOrderRestoreParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.zorder.restore'], minCount: 1, maxCount: 1 },
  },
    });

  const commandIds: string[] = [];
  runtime.registry.registerCommand<DrawingSelectParams>({
    id: 'drawing.select',
    execute: (params, context) => {
      drawingRuntime.select(params.sheetId, params.drawingIds, params.mode ?? 'replace');
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
    },
  });
  commandIds.push('drawing.select');
  runtime.registry.registerCommand<{ sheetId: string; drawingIds?: string[] }>({
    id: 'drawing.deselect',
    execute: (params, context) => {
      drawingRuntime.deselect(params.sheetId, params.drawingIds);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
    },
  });
  commandIds.push('drawing.deselect');

  runtime.registry.registerCommand<DrawingAddParams>({ id: 'drawing.add', execute: (params, context) => executeAdd(params, context) });
  commandIds.push('drawing.add');
  for (const kind of ['image', 'shape', 'data-chart', 'camera', 'textbox', 'form-control', 'slicer', 'timeline'] as const) {
    const id = `drawing.add.${kind}`;
    runtime.registry.registerCommand<DrawingAddParams>({ id, execute: (params, context) => executeAdd(params, context, kind) });
    commandIds.push(id);
  }
  runtime.registry.registerCommand<DrawingConnectorAddParams>({ id: 'drawing.connector.add', execute: (params, context) => executeConnectorAdd(params, context) });
  commandIds.push('drawing.connector.add');
  runtime.registry.registerCommand<DrawingConnectorUpdateParams>({ id: 'drawing.connector.update', execute: (params, context) => executeConnectorUpdate(params, context) });
  commandIds.push('drawing.connector.update');
  runtime.registry.registerCommand<DrawingConnectorUpdateParams>({ id: 'drawing.connector.route', execute: (params, context) => executeConnectorRoute(params, context) });
  commandIds.push('drawing.connector.route');
  runtime.registry.registerCommand<DrawingGroupParams>({ id: 'drawing.group', execute: (params, context) => executeGroup(params, context) });
  commandIds.push('drawing.group');
  runtime.registry.registerCommand<DrawingUngroupParams>({ id: 'drawing.ungroup', execute: (params, context) => executeUngroup(params, context) });
  commandIds.push('drawing.ungroup');
  runtime.registry.registerCommand<DrawingSnapSettingsParams>({ id: 'drawing.snapSettings.set', execute: (params, context) => executeSnapSettings(params, context) });
  commandIds.push('drawing.snapSettings.set');

  runtime.registry.registerCommand<DrawingRemoveParams>({
    id: 'drawing.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      const payload = sheet.drawingPayloads.get(drawing.payloadId);
      if (!payload) throw new Error(`Missing drawing payload: ${drawing.payloadId}`);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawing: structuredClone(drawing), payload: structuredClone(payload) }, affectedRanges }],
        apply: () => removeDrawing(context.workbook.getSheet(params.sheetId), params.drawingId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.remove');

  for (const id of ['drawing.move', 'drawing.resize', 'drawing.rotate'] as const) {
    runtime.registry.registerCommand<DrawingTransformParams>({ id, execute: (params, context) => executeTransform(params, context) });
    commandIds.push(id);
  }
  runtime.registry.registerCommand<DrawingTransformCommitParams>({ id: 'drawing.transform.commit', execute: (params, context) => executeTransformCommit(params, context) });
  commandIds.push('drawing.transform.commit');
  runtime.registry.registerCommand<DrawingTransformBatchParams>({ id: 'drawing.transform.batch', execute: (params, context) => executeTransformBatch(params, context) });
  commandIds.push('drawing.transform.batch');

  runtime.registry.registerCommand<DrawingAnchorParams>({
    id: 'drawing.anchor.set',
    execute: (params, context) => {
      const drawing = findDrawing(context.workbook.getSheet(params.sheetId), params.drawingId);
      const previous = structuredClone(drawing.anchor);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.anchor',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.anchor', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, anchor: previous }, affectedRanges }],
        apply: () => { findDrawing(context.workbook.getSheet(params.sheetId), params.drawingId).anchor = structuredClone(params.anchor); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.anchor.set');

  runtime.registry.registerCommand<DrawingPayloadUpdateParams>({
    id: 'drawing.payload.update',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const current = sheet.drawingPayloads.get(params.payloadId);
      if (!current) throw new Error(`Missing drawing payload: ${params.payloadId}`);
      if (current.kind !== params.before.kind) throw new Error(`Drawing payload changed before update: ${params.payloadId}`);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.payload.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: params.payloadId, before: params.after, after: params.before }, affectedRanges }],
        apply: () => updatePayload(context.workbook.getSheet(params.sheetId), params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.payload.update');

  runtime.registry.registerCommand<DrawingTextBoxUpdateParams>({
    id: 'drawing.textbox.update',
    execute: (params, context) => {
      if (!isTextBoxUpdateParams(params)) throw new Error(`Invalid textbox update payload`);
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      if (drawing.kind !== 'textbox') throw new Error(`Drawing is not a textbox: ${params.drawingId}`);
      const current = sheet.drawingPayloads.get(drawing.payloadId);
      if (!current || current.kind !== 'textbox') throw new Error(`Missing textbox payload: ${drawing.payloadId}`);
      isDrawingPairValid(drawing, params.payload);
      const affectedRanges = sheetRange(params.sheetId);
      const payloadId = drawing.payloadId;
      const before = structuredClone(current);
      context.applyMutation({
        id: 'drawing.payload.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, payloadId, before, after: structuredClone(params.payload) },
        affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId, before: structuredClone(params.payload), after: before }, affectedRanges }],
        apply: () => updatePayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, payloadId, before, after: structuredClone(params.payload) }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.textbox.update');

  runtime.registry.registerCommand<DrawingZOrderParams>({
    id: 'drawing.zorder',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      const previous = sheet.drawings.map((entry) => ({ drawingId: entry.id, zIndex: entry.zIndex }));
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.zorder',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'drawing.zorder.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, entries: previous }, affectedRanges }],
        apply: () => { reorderDrawing(context.workbook.getSheet(params.sheetId), drawing.id, params.direction); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.zorder');

  runtime.registry.registerCommand<DrawingAlignParams>({
    id: 'drawing.align',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawings = params.drawingIds.map((id) => findDrawing(sheet, id));
      if (drawings.length < 2) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const minX = Math.min(...drawings.map((entry) => entry.transform.x));
      const maxRight = Math.max(...drawings.map((entry) => entry.transform.x + entry.transform.width));
      const minY = Math.min(...drawings.map((entry) => entry.transform.y));
      const maxBottom = Math.max(...drawings.map((entry) => entry.transform.y + entry.transform.height));
      const centerX = minX + (maxRight - minX) / 2;
      const centerY = minY + (maxBottom - minY) / 2;
      const entries = drawings.map((drawing) => {
        const before = structuredClone(drawing.transform);
        const after = structuredClone(before);
        if (params.alignment === 'left') after.x = minX;
        if (params.alignment === 'center') after.x = centerX - after.width / 2;
        if (params.alignment === 'right') after.x = maxRight - after.width;
        if (params.alignment === 'top') after.y = minY;
        if (params.alignment === 'middle') after.y = centerY - after.height / 2;
        if (params.alignment === 'bottom') after.y = maxBottom - after.height;
        return { drawingId: drawing.id, before, after };
      });
      return executeTransformBatch({ sheetId: params.sheetId, entries }, context);
    },
  });
  commandIds.push('drawing.align');

  runtime.registry.registerCommand<DrawingDistributeParams>({
    id: 'drawing.distribute',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawings = params.drawingIds.map((id) => findDrawing(sheet, id)).sort((left, right) => params.axis === 'horizontal' ? left.transform.x - right.transform.x : left.transform.y - right.transform.y);
      if (drawings.length < 3) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const first = params.axis === 'horizontal' ? drawings[0]!.transform.x : drawings[0]!.transform.y;
      const lastDrawing = drawings[drawings.length - 1]!;
      const last = params.axis === 'horizontal' ? lastDrawing.transform.x : lastDrawing.transform.y;
      const step = (last - first) / (drawings.length - 1);
      const entries = drawings.map((drawing, index) => {
        const before = structuredClone(drawing.transform);
        const after = structuredClone(before);
        if (params.axis === 'horizontal') after.x = first + step * index;
        else after.y = first + step * index;
        return { drawingId: drawing.id, before, after };
      });
      return executeTransformBatch({ sheetId: params.sheetId, entries }, context);
    },
  });
  commandIds.push('drawing.distribute');

  runtime.registry.registerCommand<DrawingCopyParams>({
    id: 'drawing.copy',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const source = findDrawing(sheet, params.sourceDrawingId);
      const sourcePayload = sheet.drawingPayloads.get(source.payloadId);
      if (!sourcePayload) throw new Error(`Missing drawing payload: ${source.payloadId}`);
      const drawing: DrawingObject = {
        ...structuredClone(source),
        id: params.drawingId,
        payloadId: params.payloadId,
        transform: {
          ...source.transform,
          x: source.transform.x + params.offset.x,
          y: source.transform.y + params.offset.y,
        },
      };
      const payload: DrawingPayload = sourcePayload.kind === 'chart'
        ? { ...structuredClone(sourcePayload), chartId: params.payloadId }
        : structuredClone(sourcePayload);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, drawing, payload },
        affectedRanges,
        inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: drawing.id }, affectedRanges }],
        apply: () => addDrawing(context.workbook.getSheet(params.sheetId), drawing, payload),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.copy');

  runtime.registry.registerCommand<DrawingImageCropParams>({
    id: 'drawing.image.crop',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      if (drawing.kind !== 'image') throw new Error(`Drawing is not an image: ${params.drawingId}`);
      const current = sheet.drawingPayloads.get(drawing.payloadId);
      if (!current || current.kind !== 'image') throw new Error(`Missing image payload: ${drawing.payloadId}`);
      const before = structuredClone(current);
      const after: ImageDrawingPayload = { ...structuredClone(current), crop: structuredClone(params.crop) };
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.payload.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after },
        affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before: after, after: before }, affectedRanges }],
        apply: () => updatePayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.image.crop');

  runtime.registry.registerCommand<DrawingImageAltTextParams>({
    id: 'drawing.image.altText',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      if (drawing.kind !== 'image') throw new Error(`Drawing is not an image: ${params.drawingId}`);
      const current = sheet.drawingPayloads.get(drawing.payloadId);
      if (!current || current.kind !== 'image') throw new Error(`Missing image payload: ${drawing.payloadId}`);
      const before = structuredClone(current);
      const after = { ...structuredClone(current), altText: params.altText };
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.payload.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after },
        affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before: after, after: before }, affectedRanges }],
        apply: () => updatePayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.image.altText');

  runtime.registry.registerCommand<DrawingImageEffectsParams>({
    id: 'drawing.image.effects',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = findDrawing(sheet, params.drawingId);
      if (drawing.kind !== 'image') throw new Error(`Drawing is not an image: ${params.drawingId}`);
      const current = sheet.drawingPayloads.get(drawing.payloadId);
      if (!current || current.kind !== 'image') throw new Error(`Missing image payload: ${drawing.payloadId}`);
      if (!isImageEffects(params.effects)) throw new Error('Image effects are invalid');
      const before = structuredClone(current);
      const after: ImageDrawingPayload = { ...structuredClone(current), effects: structuredClone(params.effects) };
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.payload.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after },
        affectedRanges,
        inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: drawing.payloadId, before: after, after: before }, affectedRanges }],
        apply: () => updatePayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, payloadId: drawing.payloadId, before, after }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.image.effects');

  return commandIds;
}

export const DRAWING_MUTATION_IDS = [
  'drawing.add',
  'drawing.connector.add',
  'drawing.connector.update',
  'drawing.connector.route',
  'drawing.remove',
  'drawing.transform',
  'drawing.transform.batch',
  'drawing.anchor',
  'drawing.payload.update',
  'drawing.zorder',
  'drawing.zorder.restore',
  'drawing.group',
  'drawing.ungroup',
  'drawing.snapSettings',
] as const;
