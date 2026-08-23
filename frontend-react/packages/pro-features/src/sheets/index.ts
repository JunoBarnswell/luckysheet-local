import type { ChartModel, PivotChartReference, PivotLayout, PivotModel, PivotSlicer, PivotTimeline, RangeRef, ShapeModel, SparklineModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

export * from './chart-renderer';
export * from './sparkline-renderer';
export * from './shape-renderer';
export * from './pivot-engine';
export * from './pivot-write';
export * from './pivot-panel-state';
export * from './chart-commands';
export * from './pivot-commands-ext';
export * from './sparkline-commands';

import { registerChartDrawingCommands } from './chart-commands';
import { registerExtendedPivotCommands } from './pivot-commands-ext';
import { registerSparklineCommands } from './sparkline-commands';

export interface AddChartParams extends ChartModel {}
export interface AddPivotParams extends PivotModel {}
export interface AddShapeParams extends ShapeModel {}
export interface AddSparklineParams extends SparklineModel {}

function range(sheetId: string): RangeRef[] {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function removeById<T extends { id: string }>(items: T[], id: string): T | undefined {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return undefined;
  return items.splice(index, 1)[0];
}

export function registerProSheetCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation('chart.add', (item, context) =>
    context.workbook.getSheet(item.sheetId).charts.push(item.params as ChartModel),
  );
  runtime.registry.registerMutation('chart.remove', (item, context) => {
    removeById(context.workbook.getSheet(item.sheetId).charts, item.params as string);
  });
  runtime.registry.registerMutation('pivot.add', (item, context) =>
    context.workbook.getSheet(item.sheetId).pivots.push(item.params as PivotModel),
  );
  runtime.registry.registerMutation('pivot.remove', (item, context) => {
    removeById(context.workbook.getSheet(item.sheetId).pivots, item.params as string);
  });
  runtime.registry.registerMutation('pivot.update', (item, context) => {
    const params = item.params as { pivotId: string; sheetId: string; sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[] };
    const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
    if (pivot) {
      if (params.layout) pivot.layout = structuredClone(params.layout);
      if (params.sourceRange) pivot.sourceRange = structuredClone(params.sourceRange);
      if (params.sourceRange) pivot.fieldCatalog = undefined;
      if (params.slicers) pivot.slicers = structuredClone(params.slicers);
      if (params.timelines) pivot.timelines = structuredClone(params.timelines);
      if (params.chartReferences) pivot.chartReferences = structuredClone(params.chartReferences);
    }
  });
  runtime.registry.registerMutation('pivot.refresh', (item, context) => {
    const params = item.params as { pivotId: string; sheetId: string; refreshRevision: number; lastRefreshedAt: string };
    const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
    if (pivot) {
      pivot.refreshRevision = params.refreshRevision;
      pivot.lastRefreshedAt = params.lastRefreshedAt;
    }
  });
  runtime.registry.registerMutation('shape.add', (item, context) =>
    context.workbook.getSheet(item.sheetId).shapes.push(item.params as ShapeModel),
  );
  runtime.registry.registerMutation('shape.remove', (item, context) => {
    removeById(context.workbook.getSheet(item.sheetId).shapes, item.params as string);
  });

  registerAddCommand(
    runtime,
    'pro.chart.add',
    'chart.add',
    'chart.remove',
    (params: AddChartParams) => params,
    (sheet) => sheet.charts,
  );
  runtime.registry.registerCommand<{ pivotId: string; sheetId: string; sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[] }>({
    id: 'pro.pivot.update',
    execute: (input, context) => {
      const params = input as { pivotId: string; sheetId: string; sourceRange?: RangeRef; layout?: PivotLayout; slicers?: PivotSlicer[]; timelines?: PivotTimeline[]; chartReferences?: PivotChartReference[] };
      const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
      if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = { sourceRange: structuredClone(pivot.sourceRange), layout: structuredClone(pivot.layout), slicers: structuredClone(pivot.slicers), timelines: structuredClone(pivot.timelines), chartReferences: structuredClone(pivot.chartReferences) };
      const affectedRanges = range(params.sheetId);
      context.applyMutation({
        id: 'pivot.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { pivotId: params.pivotId, sheetId: params.sheetId, ...previous }, affectedRanges }],
        apply: () => {
          if (params.layout) pivot.layout = structuredClone(params.layout);
          if (params.sourceRange) pivot.sourceRange = structuredClone(params.sourceRange);
          if (params.sourceRange) pivot.fieldCatalog = undefined;
          if (params.slicers) pivot.slicers = structuredClone(params.slicers);
          if (params.timelines) pivot.timelines = structuredClone(params.timelines);
          if (params.chartReferences) pivot.chartReferences = structuredClone(params.chartReferences);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ pivotId: string; sheetId: string }>({
    id: 'pro.pivot.refresh',
    execute: (params, context) => {
      const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
      if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const next = {
        pivotId: params.pivotId,
        sheetId: params.sheetId,
        refreshRevision: (pivot.refreshRevision ?? 0) + 1,
        lastRefreshedAt: new Date().toISOString(),
      };
      const previous = {
        pivotId: params.pivotId,
        sheetId: params.sheetId,
        refreshRevision: pivot.refreshRevision ?? 0,
        lastRefreshedAt: pivot.lastRefreshedAt ?? '',
      };
      const affectedRanges = range(params.sheetId);
      context.applyMutation({
        id: 'pivot.refresh',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: next,
        affectedRanges,
        inverse: [{ id: 'pivot.refresh', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => {
          pivot.refreshRevision = next.refreshRevision;
          pivot.lastRefreshedAt = next.lastRefreshedAt;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  registerAddCommand(
    runtime,
    'pro.pivot.add',
    'pivot.add',
    'pivot.remove',
    (params: AddPivotParams) => params,
    (sheet) => sheet.pivots,
  );
  registerAddCommand(
    runtime,
    'pro.shape.add',
    'shape.add',
    'shape.remove',
    (params: AddShapeParams) => params,
    (sheet) => sheet.shapes,
  );

  // 浮动对象移动/缩放(图表与形状共用 update 语义)
  const registerBoundsCommand = <T extends { id: string; sheetId: string; bounds: { x: number; y: number; width: number; height: number } }>(
    commandId: string,
    mutationId: string,
    collection: (sheet: ReturnType<CommandRuntime['workbook']['getSheet']>) => Array<{ id: string; bounds: { x: number; y: number; width: number; height: number } }>,
  ): void => {
    runtime.registry.registerMutation(mutationId, (item, context) => {
      const params = item.params as T;
      const target = collection(context.workbook.getSheet(params.sheetId)).find((entry) => entry.id === params.id);
      if (target) target.bounds = { ...params.bounds };
    });
    runtime.registry.registerCommand<T>({
      id: commandId,
      execute: (input, context) => {
        const params = input as T;
        const entry = collection(context.workbook.getSheet(params.sheetId)).find((item) => item.id === params.id);
        const previousBounds = entry ? { ...entry.bounds } : params.bounds;
        const affectedRanges = range(params.sheetId);
        context.applyMutation({
          id: mutationId,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params,
          affectedRanges,
          inverse: [
            {
              id: mutationId,
              unitId: context.workbook.unitId,
              sheetId: params.sheetId,
              params: { ...params, bounds: previousBounds },
              affectedRanges,
            },
          ],
          apply: () => {
            const target = collection(context.workbook.getSheet(params.sheetId)).find((item) => item.id === params.id);
            if (target) target.bounds = { ...params.bounds };
          },
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
  };

  registerBoundsCommand<{ id: string; sheetId: string; bounds: { x: number; y: number; width: number; height: number } }>(
    'pro.chart.move',
    'chart.update',
    (sheet) => sheet.charts,
  );
  registerBoundsCommand<{ id: string; sheetId: string; bounds: { x: number; y: number; width: number; height: number } }>(
    'pro.shape.move',
    'shape.update',
    (sheet) => sheet.shapes,
  );

  registerChartDrawingCommands(runtime);
  registerExtendedPivotCommands(runtime);
  registerSparklineCommands(runtime);
}

function registerAddCommand<T extends { id: string; sheetId: string }>(
  runtime: CommandRuntime,
  commandId: string,
  mutationId: string,
  removeMutationId: string,
  normalize: (params: T) => T,
  collection: (sheet: ReturnType<CommandRuntime['workbook']['getSheet']>) => unknown[],
): void {
  runtime.registry.registerCommand<T>({
    id: commandId,
    execute: (input, context) => {
      const params = normalize(input);
      const affectedRanges = range(params.sheetId);
      context.applyMutation({
        id: mutationId,
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: removeMutationId,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: params.id,
            affectedRanges,
          },
        ],
        apply: () => collection(context.workbook.getSheet(params.sheetId)).push(structuredClone(params)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}
