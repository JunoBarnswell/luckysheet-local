import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { ChartDrawingPayload, ChartModel, DrawingObject, RangeRef } from '@react-sheets/core-model';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from './command-helpers';

export type P1ChartType = ChartDrawingPayload['chartType'];

export interface ChartInsertParams {
  sheetId: string;
  chartId: string;
  drawingId: string;
  bounds: { x: number; y: number; width: number; height: number };
  payload: ChartDrawingPayload;
}

export interface ChartUpdateParams {
  sheetId: string;
  chartId: string;
  payload: Partial<ChartDrawingPayload>;
}

export interface ChartSetTypeParams {
  sheetId: string;
  chartId: string;
  chartType: P1ChartType;
  stacked?: ChartDrawingPayload['stacked'];
}

export interface ChartSetLegendParams {
  sheetId: string;
  chartId: string;
  legendPosition: NonNullable<ChartDrawingPayload['legendPosition']>;
}

export interface ChartSetDataLabelsParams {
  sheetId: string;
  chartId: string;
  showDataLabels: boolean;
}

export interface ChartSetSeriesParams {
  sheetId: string;
  chartId: string;
  sourceRanges: RangeRef[];
  series?: ChartDrawingPayload['series'];
  categoryRange?: RangeRef;
}

function findChartDrawing(sheet: ReturnType<CommandRuntime['workbook']['getSheet']>, chartId: string): { drawing: DrawingObject; payload: ChartDrawingPayload } | undefined {
  const drawing = sheet.drawings.find((entry) => entry.kind === 'chart' && entry.payloadId === chartId);
  if (!drawing) return undefined;
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload || payload.kind !== 'chart') return undefined;
  return { drawing, payload };
}

function upsertLegacyChart(sheet: ReturnType<CommandRuntime['workbook']['getSheet']>, chartId: string, drawing: DrawingObject, payload: ChartDrawingPayload): void {
  const chart: ChartModel = {
    id: chartId,
    sheetId: drawing.sheetId,
    pivotId: payload.pivotId,
    type: payload.chartType === 'combo' ? 'column' : payload.chartType,
    title: payload.title,
    sourceRanges: structuredClone(payload.sourceRanges),
    series: payload.series?.map((entry) => ({ name: entry.name, range: entry.range, color: entry.color })),
    categoryRange: payload.categoryRange ? structuredClone(payload.categoryRange) : undefined,
    bounds: { ...drawing.transform },
    legendPosition: payload.legendPosition,
    showDataLabels: payload.showDataLabels,
  };
  removeById(sheet.charts, chartId);
  sheet.charts.push(chart);
}

export function registerChartDrawingCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<ChartInsertParams>(runtime, 'chart.drawing.add', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const drawing: DrawingObject = {
      id: params.drawingId,
      sheetId: params.sheetId,
      kind: 'chart',
      payloadId: params.chartId,
      anchor: { kind: 'absolute' },
      transform: { ...params.bounds, rotation: 0 },
      zIndex: sheet.drawings.length + 1,
    };
    sheet.drawings.push(drawing);
    sheet.drawingPayloads.set(params.chartId, structuredClone(params.payload));
    upsertLegacyChart(sheet, params.chartId, drawing, params.payload);
  });

  registerMutationHandler<{ sheetId: string; chartId: string; payload: ChartDrawingPayload }>(runtime, 'chart.drawing.update', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const current = findChartDrawing(sheet, params.chartId);
    if (!current) return;
    const next = { ...current.payload, ...params.payload, kind: 'chart' as const };
    sheet.drawingPayloads.set(params.chartId, structuredClone(next));
    upsertLegacyChart(sheet, params.chartId, current.drawing, next);
  });

  registerMutationHandler<{ sheetId: string; chartId: string }>(runtime, 'chart.drawing.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const current = findChartDrawing(sheet, params.chartId);
    if (!current) return;
    removeById(sheet.drawings, current.drawing.id);
    sheet.drawingPayloads.delete(params.chartId);
    removeById(sheet.charts, params.chartId);
  });

  runtime.registry.registerCommand<ChartInsertParams>({
    id: 'chart.insert',
    execute: (params, context) => {
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<ChartInsertParams, { sheetId: string; chartId: string }>(context, {
        id: 'chart.drawing.add',
        sheetId: params.sheetId,
        params,
        inverseId: 'chart.drawing.remove',
        inverseParams: { sheetId: params.sheetId, chartId: params.chartId },
        affectedRanges,
        apply: () => runtime.registry.getMutation('chart.drawing.add')({ id: 'chart.drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('chart.insert');

  const registerPatch = <P extends { sheetId: string; chartId: string }>(
    commandId: string,
    patch: (payload: ChartDrawingPayload, params: P) => ChartDrawingPayload,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const sheet = context.workbook.getSheet(params.sheetId);
        const current = findChartDrawing(sheet, params.chartId);
        if (!current) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
        const nextPayload = patch(current.payload, params);
        const affectedRanges = sheetRange(params.sheetId);
        applyTrackedMutation(context, {
          id: 'chart.drawing.update',
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, chartId: params.chartId, payload: nextPayload },
          inverseParams: { sheetId: params.sheetId, chartId: params.chartId, payload: current.payload },
          affectedRanges,
          apply: () => runtime.registry.getMutation('chart.drawing.update')({ id: 'chart.drawing.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, chartId: params.chartId, payload: nextPayload }, affectedRanges }, context),
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
    commandIds.push(commandId);
  };

  registerPatch<ChartUpdateParams>('chart.update', (payload, params) => ({ ...payload, ...params.payload, kind: 'chart' }));
  registerPatch<ChartSetTypeParams>('chart.setType', (payload, params) => ({ ...payload, chartType: params.chartType, stacked: params.stacked ?? payload.stacked }));
  registerPatch<ChartSetLegendParams>('chart.setLegend', (payload, params) => ({ ...payload, legendPosition: params.legendPosition }));
  registerPatch<ChartSetDataLabelsParams>('chart.setDataLabels', (payload, params) => ({ ...payload, showDataLabels: params.showDataLabels }));
  registerPatch<ChartSetSeriesParams>('chart.setSeries', (payload, params) => ({
    ...payload,
    sourceRanges: structuredClone(params.sourceRanges),
    series: params.series ? structuredClone(params.series) : payload.series,
    categoryRange: params.categoryRange ? structuredClone(params.categoryRange) : payload.categoryRange,
  }));

  runtime.registry.registerCommand<{ sheetId: string; chartId: string }>({
    id: 'chart.remove',
    execute: (params, context) => {
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'chart.drawing.remove',
        sheetId: params.sheetId,
        params,
        inverseParams: params,
        affectedRanges,
        apply: () => runtime.registry.getMutation('chart.drawing.remove')({ id: 'chart.drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('chart.remove');

  for (const chartType of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'] as const) {
    runtime.registry.registerCommand<{ sheetId: string; chartId: string; drawingId: string; bounds: ChartInsertParams['bounds']; sourceRanges: RangeRef[]; title?: string }>({
      id: `chart.insert.${chartType}`,
      execute: (params, context) =>
        runtime.registry.getCommand<ChartInsertParams>('chart.insert').execute(
          {
            sheetId: params.sheetId,
            chartId: params.chartId,
            drawingId: params.drawingId,
            bounds: params.bounds,
            payload: {
              kind: 'chart',
              chartId: params.chartId,
              chartType,
              sourceRanges: params.sourceRanges,
              title: params.title,
              legendPosition: 'bottom',
              showDataLabels: false,
            },
          },
          context,
        ),
    });
    commandIds.push(`chart.insert.${chartType}`);
  }

  return commandIds;
}

export const P1_CHART_COMMAND_IDS = [
  'chart.insert',
  'chart.update',
  'chart.remove',
  'chart.setType',
  'chart.setLegend',
  'chart.setDataLabels',
  'chart.setSeries',
  'chart.insert.column',
  'chart.insert.bar',
  'chart.insert.line',
  'chart.insert.area',
  'chart.insert.pie',
  'chart.insert.doughnut',
  'chart.insert.scatter',
  'chart.insert.combo',
] as const;
