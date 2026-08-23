import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { ChartDrawingPayload, DrawingObject, RangeRef, WorksheetModel } from '@react-sheets/core-model';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from './command-helpers';

export type P1ChartType = ChartDrawingPayload['chartType'];

export interface ChartInsertParams {
  sheetId: string;
  chartId: string;
  drawingId: string;
  bounds: { x: number; y: number; width: number; height: number };
  payload: ChartDrawingPayload;
  /** Restore metadata. Normal insert callers omit this field. */
  zIndex?: number;
  rotation?: number;
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

interface ChartRemoveParams {
  sheetId: string;
  chartId: string;
}

/** Resolve only the unified Drawing aggregate; no legacy chart projection is consulted. */
function findChartDrawing(sheet: WorksheetModel, chartId: string): { drawing: DrawingObject; payload: ChartDrawingPayload } | undefined {
  const drawing = sheet.drawings.find((entry) => entry.kind === 'chart' && entry.payloadId === chartId);
  if (!drawing) return undefined;
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload || payload.kind !== 'chart') throw new Error(`Missing chart payload: ${chartId}`);
  if (payload.chartId !== chartId) throw new Error(`Chart payload identity mismatch: ${chartId}`);
  return { drawing, payload };
}

function addChartDrawing(sheet: WorksheetModel, params: ChartInsertParams): void {
  if (params.payload.kind !== 'chart' || params.payload.chartId !== params.chartId) {
    throw new Error(`Chart payload identity mismatch: ${params.chartId}`);
  }
  if (params.sheetId !== sheet.id) throw new Error(`Chart sheet mismatch: ${params.chartId}`);
  if (sheet.drawings.some((entry) => entry.id === params.drawingId)) {
    throw new Error(`Drawing already exists: ${params.drawingId}`);
  }
  if (sheet.drawingPayloads.has(params.chartId)) {
    throw new Error(`Chart payload already exists: ${params.chartId}`);
  }
  sheet.drawings.push({
    id: params.drawingId,
    sheetId: params.sheetId,
    kind: 'chart',
    payloadId: params.chartId,
    anchor: { kind: 'absolute' },
    transform: { ...params.bounds, rotation: params.rotation ?? 0 },
    zIndex: params.zIndex ?? (sheet.drawings.length + 1),
  });
  sheet.drawingPayloads.set(params.chartId, structuredClone(params.payload));
}

function removeChartDrawing(sheet: WorksheetModel, chartId: string): { drawing: DrawingObject; payload: ChartDrawingPayload } {
  const current = findChartDrawing(sheet, chartId);
  if (!current) throw new Error(`Unknown chart: ${chartId}`);
  removeById(sheet.drawings, current.drawing.id);
  sheet.drawingPayloads.delete(chartId);
  return { drawing: structuredClone(current.drawing), payload: structuredClone(current.payload) };
}

function updateChartDrawing(sheet: WorksheetModel, params: { sheetId: string; chartId: string; payload: ChartDrawingPayload }): void {
  const current = findChartDrawing(sheet, params.chartId);
  if (!current) throw new Error(`Unknown chart: ${params.chartId}`);
  if (params.payload.kind !== 'chart' || params.payload.chartId !== params.chartId) {
    throw new Error(`Chart payload identity mismatch: ${params.chartId}`);
  }
  sheet.drawingPayloads.set(params.chartId, structuredClone(params.payload));
}

export function registerChartDrawingCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<ChartInsertParams>(runtime, 'chart.drawing.add', (params, context) => {
    addChartDrawing(context.workbook.getSheet(params.sheetId), params);
  });

  registerMutationHandler<{ sheetId: string; chartId: string; payload: ChartDrawingPayload }>(runtime, 'chart.drawing.update', (params, context) => {
    updateChartDrawing(context.workbook.getSheet(params.sheetId), params);
  });

  registerMutationHandler<ChartRemoveParams>(runtime, 'chart.drawing.remove', (params, context) => {
    removeChartDrawing(context.workbook.getSheet(params.sheetId), params.chartId);
  });

  runtime.registry.registerCommand<ChartInsertParams>({
    id: 'chart.insert',
    execute: (params, context) => {
      if (params.payload.chartId !== params.chartId) throw new Error(`Chart payload identity mismatch: ${params.chartId}`);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation<ChartInsertParams, ChartRemoveParams>(context, {
        id: 'chart.drawing.add',
        sheetId: params.sheetId,
        params,
        inverseId: 'chart.drawing.remove',
        inverseParams: { sheetId: params.sheetId, chartId: params.chartId },
        affectedRanges,
        apply: () => addChartDrawing(context.workbook.getSheet(params.sheetId), params),
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
        if (nextPayload.kind !== 'chart' || nextPayload.chartId !== params.chartId) {
          throw new Error(`Chart payload identity mismatch: ${params.chartId}`);
        }
        const affectedRanges = sheetRange(params.sheetId);
        const updateParams = { sheetId: params.sheetId, chartId: params.chartId, payload: nextPayload };
        const inverseParams = { sheetId: params.sheetId, chartId: params.chartId, payload: current.payload };
        applyTrackedMutation(context, {
          id: 'chart.drawing.update',
          sheetId: params.sheetId,
          params: updateParams,
          inverseParams,
          affectedRanges,
          apply: () => updateChartDrawing(context.workbook.getSheet(params.sheetId), updateParams),
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
      const sheet = context.workbook.getSheet(params.sheetId);
      const current = findChartDrawing(sheet, params.chartId);
      if (!current) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const affectedRanges = sheetRange(params.sheetId);
      const inverseParams: ChartInsertParams = {
        sheetId: params.sheetId,
        chartId: params.chartId,
        drawingId: current.drawing.id,
        bounds: { ...current.drawing.transform },
        payload: structuredClone(current.payload),
        zIndex: current.drawing.zIndex,
        rotation: current.drawing.transform.rotation,
      };
      applyTrackedMutation<ChartRemoveParams, ChartInsertParams>(context, {
        id: 'chart.drawing.remove',
        sheetId: params.sheetId,
        params,
        inverseId: 'chart.drawing.add',
        inverseParams,
        affectedRanges,
        apply: () => removeChartDrawing(context.workbook.getSheet(params.sheetId), params.chartId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('chart.remove');

  for (const chartType of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'] as const) {
    runtime.registry.registerCommand<{
      sheetId: string;
      chartId: string;
      drawingId: string;
      bounds: ChartInsertParams['bounds'];
      sourceRanges: RangeRef[];
      title?: string;
      stacked?: ChartDrawingPayload['stacked'];
    }>({
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
              stacked: params.stacked,
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
