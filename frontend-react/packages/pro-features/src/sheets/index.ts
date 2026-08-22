import type { ChartModel, PivotModel, RangeRef, ShapeModel, SparklineModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

export * from './chart-renderer';
export * from './sparkline-renderer';
export * from './shape-renderer';
export * from './pivot-engine';

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
  runtime.registry.registerMutation('shape.add', (item, context) =>
    context.workbook.getSheet(item.sheetId).shapes.push(item.params as ShapeModel),
  );
  runtime.registry.registerMutation('shape.remove', (item, context) => {
    removeById(context.workbook.getSheet(item.sheetId).shapes, item.params as string);
  });
  runtime.registry.registerMutation('sparkline.add', (item, context) =>
    context.workbook.getSheet(item.sheetId).sparklines.push(item.params as SparklineModel),
  );
  runtime.registry.registerMutation('sparkline.remove', (item, context) => {
    removeById(context.workbook.getSheet(item.sheetId).sparklines, item.params as string);
  });

  registerAddCommand(
    runtime,
    'pro.chart.add',
    'chart.add',
    'chart.remove',
    (params: AddChartParams) => params,
    (sheet) => sheet.charts,
  );
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
  registerAddCommand(
    runtime,
    'pro.sparkline.add',
    'sparkline.add',
    'sparkline.remove',
    (params: AddSparklineParams) => params,
    (sheet) => sheet.sparklines,
  );
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
