import type { CommandRuntime } from '@react-sheets/command-runtime';
import type {
  ChartModel,
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  FloatingImage,
  ShapeModel,
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

function syncLegacyCollections(sheet: WorksheetModel, drawing: DrawingObject, payload: DrawingPayload, mode: 'upsert' | 'remove'): void {
  if (payload.kind === 'chart') {
    const bounds = drawing.transform;
    const chart: ChartModel = {
      id: drawing.payloadId,
      sheetId: drawing.sheetId,
      pivotId: payload.pivotId,
      type: payload.chartType === 'combo' ? 'column' : payload.chartType,
      title: payload.title,
      sourceRanges: structuredClone(payload.sourceRanges),
      series: payload.series?.map((entry) => ({ name: entry.name, range: entry.range, color: entry.color })),
      categoryRange: payload.categoryRange ? structuredClone(payload.categoryRange) : undefined,
      bounds: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
      legendPosition: payload.legendPosition,
      showDataLabels: payload.showDataLabels,
    };
    if (mode === 'remove') removeById(sheet.charts, chart.id);
    else {
      removeById(sheet.charts, chart.id);
      sheet.charts.push(chart);
    }
    return;
  }
  if (payload.kind === 'shape' || payload.kind === 'textbox') {
    const shape: ShapeModel = {
      id: drawing.payloadId,
      sheetId: drawing.sheetId,
      type: payload.kind === 'textbox' ? 'callout' : payload.type,
      bounds: { ...drawing.transform },
      fill: payload.kind === 'textbox' ? '#ffffff' : payload.fill,
      stroke: payload.kind === 'textbox' ? '#64748b' : payload.stroke,
      strokeWidth: payload.kind === 'shape' ? payload.strokeWidth : 1,
      text: payload.text,
      textColor: payload.textColor,
      fontSize: payload.fontSize,
      rotation: drawing.transform.rotation,
    };
    if (mode === 'remove') removeById(sheet.shapes, shape.id);
    else {
      removeById(sheet.shapes, shape.id);
      sheet.shapes.push(shape);
    }
    return;
  }
  if (payload.kind === 'image') {
    const image: FloatingImage = {
      id: drawing.payloadId,
      sheetId: drawing.sheetId,
      name: payload.name,
      src: payload.src,
      bounds: { ...drawing.transform },
    };
    if (mode === 'remove') removeById(sheet.images, image.id);
    else {
      removeById(sheet.images, image.id);
      sheet.images.push(image);
    }
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
    sheet.drawings.push(structuredClone(params.drawing));
    sheet.drawingPayloads.set(params.drawing.payloadId, structuredClone(params.payload));
    syncLegacyCollections(sheet, params.drawing, params.payload, 'upsert');
  });
  registerMutationHandler<DrawingRemoveParams>(runtime, 'drawing.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const drawing = removeById(sheet.drawings, params.drawingId);
    const payload = drawing ? sheet.drawingPayloads.get(drawing.payloadId) : undefined;
    if (drawing?.payloadId) sheet.drawingPayloads.delete(drawing.payloadId);
    if (drawing && payload) syncLegacyCollections(sheet, drawing, payload, 'remove');
  });
  registerMutationHandler<DrawingTransformParams>(runtime, 'drawing.transform', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
    if (!drawing) return;
    drawing.transform = structuredClone(params.transform);
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (payload) syncLegacyCollections(sheet, drawing, payload, 'upsert');
  });
  registerMutationHandler<DrawingZOrderParams>(runtime, 'drawing.zorder', (params, context) => {
    reorderDrawing(context.workbook.getSheet(params.sheetId), params.drawingId, params.direction);
  });
  registerMutationHandler<{ sheetId: string; drawingId: string; zIndex: number }>(runtime, 'drawing.zindex.set', (params, context) => {
    const drawing = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
    if (drawing) drawing.zIndex = params.zIndex;
  });

  const registerTransformCommand = (commandId: string, mutationId: string): void => {
    runtime.registry.registerCommand<DrawingTransformParams>({
      id: commandId,
      execute: (params, context) => {
        const sheet = context.workbook.getSheet(params.sheetId);
        const drawing = sheet.drawings.find((entry) => entry.id === params.drawingId);
        const previous = drawing ? structuredClone(drawing.transform) : structuredClone(params.transform);
        const affectedRanges = sheetRange(params.sheetId);
        applyTrackedMutation(context, {
          id: mutationId,
          sheetId: params.sheetId,
          params,
          inverseParams: { ...params, transform: previous },
          affectedRanges,
          apply: () => {
            const target = context.workbook.getSheet(params.sheetId).drawings.find((entry) => entry.id === params.drawingId);
            if (!target) return;
            target.transform = structuredClone(params.transform);
            const payload = context.workbook.getSheet(params.sheetId).drawingPayloads.get(target.payloadId);
            if (payload) syncLegacyCollections(context.workbook.getSheet(params.sheetId), target, payload, 'upsert');
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
      const previous = drawing ? drawing.zIndex : 0;
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<DrawingZOrderParams, { sheetId: string; drawingId: string; zIndex: number }>(context, {
        id: 'drawing.zorder',
        sheetId: params.sheetId,
        params,
        inverseId: 'drawing.zindex.set',
        inverseParams: { sheetId: params.sheetId, drawingId: params.drawingId, zIndex: previous },
        affectedRanges,
        apply: () => reorderDrawing(context.workbook.getSheet(params.sheetId), params.drawingId, params.direction),
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
        apply: () => {
          const sheet = context.workbook.getSheet(params.sheetId);
          sheet.drawings.push(structuredClone(drawing));
          sheet.drawingPayloads.set(drawing.payloadId, structuredClone(payload));
          syncLegacyCollections(sheet, drawing, payload, 'upsert');
        },
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
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'drawing.remove',
        sheetId: params.sheetId,
        params,
        inverseParams: drawing && payload ? { sheetId: params.sheetId, drawing: structuredClone(drawing), payload: structuredClone(payload) } : params,
        affectedRanges,
        apply: () => {
          const current = context.workbook.getSheet(params.sheetId);
          const removed = removeById(current.drawings, params.drawingId);
          const removedPayload = removed ? current.drawingPayloads.get(removed.payloadId) : undefined;
          if (removed?.payloadId) current.drawingPayloads.delete(removed.payloadId);
          if (removed && removedPayload) syncLegacyCollections(current, removed, removedPayload, 'remove');
        },
      });
      return { operationId: context.operationId, mutationCount: drawing ? 1 : 0, affectedRanges };
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

export const DRAWING_MUTATION_IDS = ['drawing.add', 'drawing.remove', 'drawing.transform', 'drawing.zorder'] as const;
