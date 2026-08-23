import type { ChartModel, PivotModel, RangeRef, ShapeModel, SparklineModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { buildPivotWriteback } from './pivot-write';

export * from './chart-renderer';
export * from './sparkline-renderer';
export * from './shape-renderer';
export * from './pivot-engine';
export * from './pivot-write';

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

  // 数据透视写回
  runtime.registry.registerMutation('pro.pivot.write', (item, context) => {
    const params = item.params as {
      sheetId: string;
      pivotId: string;
      targetStartRow: number;
      targetStartColumn: number;
      values: Array<Array<{ value: string | number | boolean | null }>>;
    };
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let r = 0; r < params.values.length; r++) {
      const rowValues = params.values[r]!;
      for (let c = 0; c < rowValues.length; c++) {
        sheet.cells.set(params.targetStartRow + r, params.targetStartColumn + c, structuredClone(rowValues[c]!));
      }
    }
  });
  runtime.registry.registerCommand<{ sheetId: string; pivotId: string }>({
    id: 'pro.pivot.write',
    execute: (input, context) => {
      const params = input as { sheetId: string; pivotId: string };
      const sheet = context.workbook.getSheet(params.sheetId);
      const pivot = sheet.pivots.find((entry) => entry.id === params.pivotId);
      if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const writeback = buildPivotWriteback(pivot, sheet);
      const affectedRanges: RangeRef[] = [];
      for (let r = 0; r < writeback.values.length; r++) {
        affectedRanges.push({
          sheetId: params.sheetId,
          startRow: writeback.targetStartRow + r,
          endRow: writeback.targetStartRow + r,
          startColumn: writeback.targetStartColumn,
          endColumn: writeback.targetStartColumn + writeback.values[r]!.length - 1,
        });
      }
      context.applyMutation({
        id: 'pro.pivot.write',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ...writeback },
        affectedRanges,
        inverse: [
          {
            id: 'range.clear' as never,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: {
              sheetId: params.sheetId,
              range: {
                sheetId: params.sheetId,
                startRow: writeback.targetStartRow,
                endRow: writeback.targetStartRow + writeback.values.length - 1,
                startColumn: writeback.targetStartColumn,
                endColumn:
                  writeback.targetStartColumn +
                  Math.max(...writeback.values.map((row: { length: number }) => row.length)) -
                  1,
              },
              mode: 'contents',
            },
            affectedRanges,
          },
        ],
        apply: () => {
          for (let r = 0; r < writeback.values.length; r++) {
            const rowValues = writeback.values[r]!;
            for (let c = 0; c < rowValues.length; c++) {
              sheet.cells.set(writeback.targetStartRow + r, writeback.targetStartColumn + c, structuredClone(rowValues[c]!));
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
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
