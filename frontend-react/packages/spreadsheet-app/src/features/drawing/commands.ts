import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type {
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  WorksheetModel,
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

export interface DrawingImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

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

type ImagePayloadWithCrop = Extract<DrawingPayload, { kind: 'image' }> & { crop?: DrawingImageCrop };

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
  return objectParams(value) && ['image', 'shape', 'chart', 'textbox'].includes(String(value.kind));
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

function isDrawingRemoveParams(value: unknown): value is DrawingRemoveParams {
  return hasSheetId(value) && hasDrawingId(value);
}

function isTransformParams(value: unknown): value is DrawingTransformParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return isTransform((value as Record<string, unknown>).transform);
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
  });
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

function isImageCropParams(value: unknown): value is DrawingImageCropParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return isImageCrop((value as Record<string, unknown>).crop);
}

function isImageAltTextParams(value: unknown): value is DrawingImageAltTextParams {
  if (!hasSheetId(value) || !hasDrawingId(value) || !objectParams(value)) return false;
  return typeof (value as Record<string, unknown>).altText === 'string';
}

function isDrawingPairValid(drawing: DrawingObject, payload: DrawingPayload): void {
  if (!drawing.id || !drawing.sheetId || !drawing.payloadId) throw new Error(`Drawing identity is required: ${drawing.id}`);
  if (!isAnchor(drawing.anchor)) throw new Error(`Invalid drawing anchor: ${drawing.id}`);
  if (!isTransform(drawing.transform)) throw new Error(`Invalid drawing transform: ${drawing.id}`);
  if (drawing.kind !== payload.kind) throw new Error(`Drawing payload kind mismatch: ${drawing.id}`);
  if (payload.kind === 'chart' && payload.chartId !== drawing.payloadId) {
    throw new Error(`Drawing payload identity mismatch: ${drawing.payloadId}`);
  }
}

function addDrawing(sheet: WorksheetModel, drawing: DrawingObject, payload: DrawingPayload): void {
  isDrawingPairValid(drawing, payload);
  if (drawing.sheetId !== sheet.id) throw new Error(`Drawing sheet mismatch: ${drawing.id}`);
  if (sheet.drawings.some((entry) => entry.id === drawing.id)) throw new Error(`Drawing already exists: ${drawing.id}`);
  if (sheet.drawingPayloads.has(drawing.payloadId)) throw new Error(`Drawing payload already exists: ${drawing.payloadId}`);
  sheet.drawings.push(structuredClone(drawing));
  sheet.drawingPayloads.set(drawing.payloadId, structuredClone(payload));
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

function executeAdd(params: DrawingAddParams, context: CommandContext, kind?: DrawingObject['kind']): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const drawing = kind ? { ...params.drawing, kind } : params.drawing;
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
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.transform',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{ id: 'drawing.transform', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, transform: previous }, affectedRanges }],
    apply: () => { findDrawing(context.workbook.getSheet(params.sheetId), params.drawingId).transform = structuredClone(params.transform); },
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeTransformCommit(params: DrawingTransformCommitParams, context: CommandContext): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const drawing = findDrawing(context.workbook.getSheet(params.sheetId), params.drawingId);
  if (JSON.stringify(drawing.transform) !== JSON.stringify(params.before)) throw new Error(`Drawing transform changed before pointer commit: ${params.drawingId}`);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.transform',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: { sheetId: params.sheetId, drawingId: params.drawingId, transform: structuredClone(params.after) },
    affectedRanges,
    inverse: [{ id: 'drawing.transform', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawingId, transform: structuredClone(params.before) }, affectedRanges }],
    apply: () => { findDrawing(context.workbook.getSheet(params.sheetId), params.drawingId).transform = structuredClone(params.after); },
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
  const affectedRanges = sheetRange(params.sheetId);
  const inverse: DrawingTransformBatchParams = {
    sheetId: params.sheetId,
    entries: params.entries.map((entry) => ({ drawingId: entry.drawingId, before: entry.after, after: entry.before })),
  };
  context.applyMutation({
    id: 'drawing.transform.batch',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{ id: 'drawing.transform.batch', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges }],
    apply: () => {
      const targetSheet = context.workbook.getSheet(params.sheetId);
      for (const entry of params.entries) findDrawing(targetSheet, entry.drawingId).transform = structuredClone(entry.after);
    },
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

export function registerDrawingCommands(runtime: CommandRuntime, drawingRuntime: DrawingRuntime): string[] {
  runtime.registry.registerMutation<DrawingAddParams>('drawing.add', (item, context) => {
    addDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawing, item.params.payload);
  }, {
    schema: { name: 'DrawingAddParams', validate: isDrawingAddParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.remove'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingRemoveParams>('drawing.remove', (item, context) => {
    removeDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId);
  }, {
    schema: { name: 'DrawingRemoveParams', validate: isDrawingRemoveParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.add'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingTransformParams>('drawing.transform', (item, context) => {
    findDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId).transform = structuredClone(item.params.transform);
  }, {
    schema: { name: 'DrawingTransformParams', validate: isTransformParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.transform'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingTransformBatchParams>('drawing.transform.batch', (item, context) => {
    for (const entry of item.params.entries) findDrawing(context.workbook.getSheet(item.params.sheetId), entry.drawingId).transform = structuredClone(entry.after);
  }, {
    schema: { name: 'DrawingTransformBatchParams', validate: isTransformBatchParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.transform.batch'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingAnchorParams>('drawing.anchor', (item, context) => {
    findDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId).anchor = structuredClone(item.params.anchor);
  }, {
    schema: { name: 'DrawingAnchorParams', validate: isAnchorParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.anchor'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingPayloadUpdateParams>('drawing.payload.update', (item, context) => {
    updatePayload(context.workbook.getSheet(item.params.sheetId), item.params);
  }, {
    schema: { name: 'DrawingPayloadUpdateParams', validate: isPayloadUpdateParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.payload.update'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingZOrderParams>('drawing.zorder', (item, context) => {
    reorderDrawing(context.workbook.getSheet(item.params.sheetId), item.params.drawingId, item.params.direction);
  }, {
    schema: { name: 'DrawingZOrderParams', validate: isZOrderParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.zorder.restore'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<DrawingZOrderRestoreParams>('drawing.zorder.restore', (item, context) => {
    restoreZOrder(context.workbook.getSheet(item.params.sheetId), item.params);
  }, {
    schema: { name: 'DrawingZOrderRestoreParams', validate: isZOrderRestoreParams },
    permission: { capability: 'drawing.edit' },
    affectedRanges: { resolve: rangesForParams, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['drawing.zorder.restore'], minCount: 1, maxCount: 1 },
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
  for (const kind of ['image', 'shape', 'textbox'] as const) {
    const id = `drawing.add.${kind}`;
    runtime.registry.registerCommand<DrawingAddParams>({ id, execute: (params, context) => executeAdd(params, context, kind) });
    commandIds.push(id);
  }

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
      const after: ImagePayloadWithCrop = { ...structuredClone(current), crop: structuredClone(params.crop) };
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

  return commandIds;
}

export const DRAWING_MUTATION_IDS = [
  'drawing.add',
  'drawing.remove',
  'drawing.transform',
  'drawing.transform.batch',
  'drawing.anchor',
  'drawing.payload.update',
  'drawing.zorder',
  'drawing.zorder.restore',
] as const;
