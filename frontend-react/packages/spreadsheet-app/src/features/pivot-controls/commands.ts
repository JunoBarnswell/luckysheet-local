import type { CommandContext, CommandResult, CommandRuntime } from '@react-sheets/command-runtime';
import type {
  DrawingObject,
  DrawingPayload,
  PivotControlFilter,
  PivotControlStyle,
  PivotSlicerDrawingPayload,
  PivotSlicerSettings,
  PivotTimelineDrawingPayload,
  PivotTimelinePeriod,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  isPivotSlicerDrawingPayload,
  isPivotTimelineDrawingPayload,
} from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import type { DrawingAddParams, DrawingPayloadUpdateParams } from '../drawing/commands';
import {
  createPivotControlFilter,
  createPivotControlStyle,
  createPivotSlicerSettings,
  createPivotTimelinePeriod,
  findPivotControlRecord,
} from './helpers';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

export interface PivotSlicerCreateParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: PivotSlicerDrawingPayload;
}

export interface PivotTimelineCreateParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: PivotTimelineDrawingPayload;
}

export interface PivotSlicerFilterSetParams {
  sheetId: string;
  drawingId: string;
  filter: PivotControlFilter;
}

export interface PivotTimelinePeriodSetParams {
  sheetId: string;
  drawingId: string;
  period: PivotTimelinePeriod;
}

export interface PivotControlStyleSetParams {
  sheetId: string;
  drawingId: string;
  style: PivotControlStyle;
}

export interface PivotSlicerSettingsSetParams {
  sheetId: string;
  drawingId: string;
  settings: PivotSlicerSettings;
}

export interface PivotControlConnectionsSetParams {
  sheetId: string;
  drawingId: string;
  connectedPivotIds: string[];
}

function isPivotControlPayload(payload: DrawingPayload): payload is PivotSlicerDrawingPayload | PivotTimelineDrawingPayload {
  return isPivotSlicerDrawingPayload(payload) || isPivotTimelineDrawingPayload(payload);
}

function validateCreatePair(sheet: WorksheetModel, params: PivotSlicerCreateParams | PivotTimelineCreateParams): void {
  if (params.sheetId !== sheet.id || params.drawing.sheetId !== sheet.id) {
    throw new Error(`Pivot control sheet mismatch: ${params.drawing.id}`);
  }
  if (params.drawing.kind !== params.payload.kind || !isPivotControlPayload(params.payload)) {
    throw new Error(`Pivot control drawing/payload mismatch: ${params.drawing.id}`);
  }
  if (sheet.drawings.some((entry) => entry.id === params.drawing.id)) {
    throw new Error(`Drawing already exists: ${params.drawing.id}`);
  }
  if (sheet.drawingPayloads.has(params.drawing.payloadId)) {
    throw new Error(`Drawing payload already exists: ${params.drawing.payloadId}`);
  }
}

function insertDrawing(sheet: WorksheetModel, params: DrawingAddParams): void {
  if (sheet.drawings.some((entry) => entry.id === params.drawing.id)) {
    throw new Error(`Drawing already exists: ${params.drawing.id}`);
  }
  if (sheet.drawingPayloads.has(params.drawing.payloadId)) {
    throw new Error(`Drawing payload already exists: ${params.drawing.payloadId}`);
  }
  sheet.drawings.push(structuredClone(params.drawing));
  sheet.drawingPayloads.set(params.drawing.payloadId, structuredClone(params.payload));
}

function updateDrawingPayload(sheet: WorksheetModel, params: DrawingPayloadUpdateParams): void {
  const current = sheet.drawingPayloads.get(params.payloadId);
  if (!current) throw new Error(`Missing drawing payload: ${params.payloadId}`);
  if (JSON.stringify(current) !== JSON.stringify(params.before)) {
    throw new Error(`Drawing payload changed before update: ${params.payloadId}`);
  }
  sheet.drawingPayloads.set(params.payloadId, structuredClone(params.after));
}

function executeCreate(
  params: PivotSlicerCreateParams | PivotTimelineCreateParams,
  context: CommandContext,
): CommandResult {
  const sheet = context.workbook.getSheet(params.sheetId);
  validateCreatePair(sheet, params);
  const mutationParams: DrawingAddParams = {
    sheetId: params.sheetId,
    drawing: structuredClone(params.drawing),
    payload: structuredClone(params.payload),
  };
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.add',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: mutationParams,
    affectedRanges,
    inverse: [{
      id: 'drawing.remove',
      unitId: context.workbook.unitId,
      sheetId: params.sheetId,
      params: { sheetId: params.sheetId, drawingId: params.drawing.id },
      affectedRanges,
    }],
    apply: () => insertDrawing(context.workbook.getSheet(params.sheetId), mutationParams),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executePayloadUpdate(
  sheetId: string,
  drawingId: string,
  context: CommandContext,
  update: (payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload) => PivotSlicerDrawingPayload | PivotTimelineDrawingPayload,
): CommandResult {
  const sheet = context.workbook.getSheet(sheetId);
  const record = findPivotControlRecord(sheet, drawingId);
  if (!record) throw new Error(`Unknown Pivot control drawing: ${drawingId}`);
  const before = structuredClone(record.payload);
  const after = update(structuredClone(record.payload));
  if (!isPivotControlPayload(after) || after.kind !== before.kind) {
    throw new Error(`Invalid Pivot control payload update: ${drawingId}`);
  }
  const mutationParams: DrawingPayloadUpdateParams = {
    sheetId,
    payloadId: record.drawing.payloadId,
    before,
    after,
  };
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({
    id: 'drawing.payload.update',
    unitId: context.workbook.unitId,
    sheetId,
    params: mutationParams,
    affectedRanges,
    inverse: [{
      id: 'drawing.payload.update',
      unitId: context.workbook.unitId,
      sheetId,
      params: {
        sheetId,
        payloadId: record.drawing.payloadId,
        before: after,
        after: before,
      },
      affectedRanges,
    }],
    apply: () => updateDrawingPayload(context.workbook.getSheet(sheetId), mutationParams),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

export const PIVOT_CONTROL_COMMAND_IDS = [
  'pivot.control.slicer.create',
  'pivot.control.timeline.create',
  'pivot.control.slicer.filter.set',
  'pivot.control.timeline.period.set',
  'pivot.control.style.set',
  'pivot.control.slicer.settings.set',
  'pivot.control.connections.set',
] as const;

export function registerPivotControlCommands(runtime: CommandRuntime): string[] {
  if (!runtime.registry.hasMutation('drawing.add') || !runtime.registry.hasMutation('drawing.payload.update')) {
    throw new Error('Pivot controls require the drawing feature to be registered first');
  }

  runtime.registry.registerCommand<PivotSlicerCreateParams>({
    id: 'pivot.control.slicer.create',
    execute: (params, context) => executeCreate(params, context),
  });
  runtime.registry.registerCommand<PivotTimelineCreateParams>({
    id: 'pivot.control.timeline.create',
    execute: (params, context) => executeCreate(params, context),
  });
  runtime.registry.registerCommand<PivotSlicerFilterSetParams>({
    id: 'pivot.control.slicer.filter.set',
    execute: (params, context) => executePayloadUpdate(params.sheetId, params.drawingId, context, (payload) => {
      if (payload.kind !== 'slicer') throw new Error(`Drawing is not a slicer: ${params.drawingId}`);
      return { ...payload, filter: createPivotControlFilter(params.filter) };
    }),
  });
  runtime.registry.registerCommand<PivotTimelinePeriodSetParams>({
    id: 'pivot.control.timeline.period.set',
    execute: (params, context) => executePayloadUpdate(params.sheetId, params.drawingId, context, (payload) => {
      if (payload.kind !== 'timeline') throw new Error(`Drawing is not a timeline: ${params.drawingId}`);
      return { ...payload, period: createPivotTimelinePeriod(params.period) };
    }),
  });
  runtime.registry.registerCommand<PivotControlStyleSetParams>({
    id: 'pivot.control.style.set',
    execute: (params, context) => executePayloadUpdate(params.sheetId, params.drawingId, context, (payload) => ({
      ...payload,
      style: createPivotControlStyle(params.style),
    })),
  });
  runtime.registry.registerCommand<PivotSlicerSettingsSetParams>({
    id: 'pivot.control.slicer.settings.set',
    execute: (params, context) => executePayloadUpdate(params.sheetId, params.drawingId, context, (payload) => {
      if (payload.kind !== 'slicer') throw new Error(`Drawing is not a slicer: ${params.drawingId}`);
      return { ...payload, settings: createPivotSlicerSettings(params.settings) };
    }),
  });
  runtime.registry.registerCommand<PivotControlConnectionsSetParams>({
    id: 'pivot.control.connections.set',
    execute: (params, context) => executePayloadUpdate(params.sheetId, params.drawingId, context, (payload) => {
      const connectedPivotIds = [...new Set(params.connectedPivotIds.map((id) => id.trim()).filter(Boolean))];
      return connectedPivotIds.length > 0
        ? { ...payload, connectedPivotIds }
        : (() => {
          const next = { ...payload };
          delete next.connectedPivotIds;
          return next;
        })();
    }),
  });

  return [...PIVOT_CONTROL_COMMAND_IDS];
}

export function registerPivotControlFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'pivot-controls',
    version: '1.0.0',
    dependencies: ['drawing', 'pivot'],
    commandIds: registerPivotControlCommands(runtime),
    permissions: ['drawing.edit'],
  };
}
