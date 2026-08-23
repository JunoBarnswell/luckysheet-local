import type { CommandRuntime } from '@react-sheets/command-runtime';
import type {
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  WorksheetModel,
} from '@react-sheets/core-model';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from '../../command-helpers';
import { DrawingRuntime, nextZIndex, reorderDrawing } from './runtime';

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

export interface DrawingZOrderParams {
  sheetId: string;
  drawingId: string;
  direction: 'forward' | 'backward' | 'front' | 'back';
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

interface DrawingZOrderRestoreParams {
  sheetId: string;
  entries: Array<{ drawingId: string; zIndex: number }>;
}

/**
 * Drawing is the only persisted floating-object aggregate.
 *
 * The worksheet still exposes deprecated projection arrays while the rest of
 * the application migrates, but command mutations must never populate or
 * synchronize them. Keeping that policy here prevents a second write source
 * from reappearing when a new drawing kind is added.
 */
function addDrawing(sheet: WorksheetModel, drawing: DrawingObject, payload: DrawingPayload): void {
  if (drawing.sheetId !== sheet.id) throw new Error(`Drawing sheet mismatch: ${drawing.id}`);
  if (payload.kind !== drawing.kind) throw new Error(`Drawing payload kind mismatch: ${drawing.id}`);
  if (payload.kind === 'chart' && payload.chartId !== drawing.payloadId) {
    throw new Error(`Drawing payload identity mismatch: ${drawing.payloadId}`);
  }
  if (sheet.drawings.some((entry) => entry.id === drawing.id)) {
    throw new Error(`Drawing already exists: ${drawing.id}`);
  }
  if (sheet.drawingPayloads.has(drawing.payloadId)) {
    throw new Error(`Drawing payload already exists: ${drawing.payloadId}`);
  }
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

function restoreZOrder(sheet: WorksheetModel, params: DrawingZOrderRestoreParams): void {
  for (const entry of params.entries) {
    const drawing = sheet.drawings.find((item) => item.id === entry.drawingId);
    if (!drawing) throw new Error(`Unknown drawing: ${entry.drawingId}`);
    drawing.zIndex = entry.zIndex;
  }
}

export function registerDrawingCommands(runtime: CommandRuntime, drawingRuntime: DrawingRuntime): string[] {
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

  registerMutationHandler<DrawingAddParams>(runtime, 'drawing.add', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    addDrawing(sheet, params.drawing, params.payload);
  });
  registerMutationHandler<DrawingRemoveParams>(runtime, 'drawing.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    removeDrawing(sheet, params.drawingId);
  });
  registerMutationHandler<DrawingTransformParams>(runtime, 'drawing.transform', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
    if (!drawing) throw new Error(`Unknown drawing: ${params.drawingId}`);
    drawing.transform = structuredClone(params.transform);
  });
  registerMutationHandler<DrawingZOrderParams>(runtime, 'drawing.zorder', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    if (!sheet.drawings.some((entry) => entry.id === params.drawingId)) throw new Error(`Unknown drawing: ${params.drawingId}`);
    reorderDrawing(sheet, params.drawingId, params.direction);
  });
  registerMutationHandler<{ sheetId: string; drawingId: string; zIndex: number }>(runtime, 'drawing.zindex.set', (params, context) => {
    const drawing = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
    if (!drawing) throw new Error(`Unknown drawing: ${params.drawingId}`);
    drawing.zIndex = params.zIndex;
  });
  registerMutationHandler<DrawingZOrderRestoreParams>(runtime, 'drawing.zorder.restore', (params, context) => {
    restoreZOrder(context.workbook.getSheet(params.sheetId), params);
  });

  const registerTransformCommand = (commandId: string, mutationId: string): void => {
    runtime.registry.registerCommand<DrawingTransformParams>({
      id: commandId,
      execute: (params, context) => {
        const sheet = context.workbook.getSheet(params.sheetId);
        const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
        if (!drawing) throw new Error(`Unknown drawing: ${params.drawingId}`);
        const previous = structuredClone(drawing.transform);
        const affectedRanges = sheetRange(params.sheetId);
        applyTrackedMutation(context, {
          id: mutationId,
          sheetId: params.sheetId,
          params,
          inverseParams: { ...params, transform: previous },
          affectedRanges,
          apply: () => {
            const target = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
            if (!target) throw new Error(`Unknown drawing: ${params.drawingId}`);
            target.transform = structuredClone(params.transform);
          },
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
    commandIds.push(commandId);
  };

  registerTransformCommand('drawing.move', 'drawing.transform');
  registerTransformCommand('drawing.resize', 'drawing.transform');
  registerTransformCommand('drawing.rotate', 'drawing.transform');

  runtime.registry.registerCommand<DrawingZOrderParams>({
    id: 'drawing.zorder',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
      if (!drawing) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const previous = sheet.drawings.map((entry) => ({ drawingId: entry.id, zIndex: entry.zIndex }));
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<DrawingZOrderParams, DrawingZOrderRestoreParams>(context, {
        id: 'drawing.zorder',
        sheetId: params.sheetId,
        params,
        inverseId: 'drawing.zorder.restore',
        inverseParams: { sheetId: params.sheetId, entries: previous },
        affectedRanges,
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
          if (!target) throw new Error(`Unknown drawing: ${params.drawingId}`);
          reorderDrawing(context.workbook.getSheet(params.sheetId), params.drawingId, params.direction);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.zorder');

  runtime.registry.registerCommand<DrawingAddParams>({
    id: 'drawing.add',
    execute: (params, context) => {
      const drawing = structuredClone(params.drawing);
      if (!drawing.zIndex) drawing.zIndex = nextZIndex(context.workbook.getSheet(params.sheetId));
      const payload = structuredClone(params.payload);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'drawing.add',
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, drawing, payload },
        inverseParams: { sheetId: params.sheetId, drawingId: drawing.id },
        affectedRanges,
        apply: () => addDrawing(context.workbook.getSheet(params.sheetId), drawing, payload),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.add');

  runtime.registry.registerCommand<DrawingRemoveParams>({
    id: 'drawing.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
      const payload = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
      if (!drawing) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      if (!payload) throw new Error(`Missing drawing payload: ${drawing.payloadId}`);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'drawing.remove',
        sheetId: params.sheetId,
        params,
        inverseId: 'drawing.add',
        inverseParams: { sheetId: params.sheetId, drawing: structuredClone(drawing), payload: structuredClone(payload) },
        affectedRanges,
        apply: () => removeDrawing(context.workbook.getSheet(params.sheetId), params.drawingId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('drawing.remove');

  const registerKindAdd = (commandId: string, kind: DrawingObject['kind']): void => {
    runtime.registry.registerCommand<DrawingAddParams & { kind?: typeof kind }>({
      id: commandId,
      execute: (params, context) => runtime.registry.getCommand<DrawingAddParams>('drawing.add').execute({ ...params, drawing: { ...params.drawing, kind } }, context),
    });
    commandIds.push(commandId);
  };

  registerKindAdd('drawing.add.image', 'image');
  registerKindAdd('drawing.add.shape', 'shape');
  registerKindAdd('drawing.add.textbox', 'textbox');

  return commandIds;
}

export const DRAWING_MUTATION_IDS = ['drawing.add', 'drawing.remove', 'drawing.transform', 'drawing.zorder', 'drawing.zindex.set', 'drawing.zorder.restore'] as const;
