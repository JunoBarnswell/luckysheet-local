import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type { ChartDrawingPayload, DrawingObject, RangeRef, WorksheetModel } from '@react-sheets/core-model';
import { applyTrackedMutation, sheetRange } from '../../command-helpers';

export type ChartType = ChartDrawingPayload['chartType'];
export type ChartAxisPosition = 'top' | 'bottom' | 'left' | 'right';
export type ChartAxisScale = 'linear' | 'logarithmic';

export interface ChartAxis {
  id: string;
  position: ChartAxisPosition;
  title?: string;
  scale?: ChartAxisScale;
  minimum?: number;
  maximum?: number;
  majorUnit?: number;
  numberFormat?: string;
  crossesAt?: number;
}

export interface ChartSeries {
  name: string;
  range: RangeRef;
  /** Scatter charts may provide explicit X/Y ranges while retaining range as the value range. */
  xRange?: RangeRef;
  yRange?: RangeRef;
  color?: string;
  chartType?: Exclude<ChartType, 'combo'>;
  axis?: 'primary' | 'secondary';
  smooth?: boolean;
}

/** Canonical chart payload. It remains a DrawingPayload member and is stored by payloadId. */
export type ChartPayload = Omit<ChartDrawingPayload, 'series'> & {
  series?: ChartSeries[];
  categoryAxis?: ChartAxis;
  valueAxis?: ChartAxis;
  secondaryCategoryAxis?: ChartAxis;
  secondaryValueAxis?: ChartAxis;
};

export interface ChartInsertParams {
  sheetId: string;
  drawing: DrawingObject;
  payload: ChartPayload;
}

export interface ChartUpdateParams {
  sheetId: string;
  chartId: string;
  payload: Partial<ChartPayload>;
}

export interface ChartSetTypeParams {
  sheetId: string;
  chartId: string;
  chartType: ChartType;
  stacked?: ChartPayload['stacked'];
}

export interface ChartSetLegendParams {
  sheetId: string;
  chartId: string;
  legendPosition: NonNullable<ChartPayload['legendPosition']>;
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
  series?: ChartPayload['series'];
  categoryRange?: RangeRef;
}

export interface ChartSetAxesParams {
  sheetId: string;
  chartId: string;
  categoryAxis?: ChartAxis;
  valueAxis?: ChartAxis;
  secondaryCategoryAxis?: ChartAxis;
  secondaryValueAxis?: ChartAxis;
}

export interface ChartSetSecondaryAxisParams {
  sheetId: string;
  chartId: string;
  seriesName: string;
  enabled: boolean;
}

interface ChartRemoveParams {
  sheetId: string;
  chartId: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && Number.isInteger(value.startRow) && Number.isInteger(value.endRow)
    && Number.isInteger(value.startColumn) && Number.isInteger(value.endColumn)
    && (value.startRow as number) >= 0 && (value.endRow as number) >= (value.startRow as number)
    && (value.startColumn as number) >= 0 && (value.endColumn as number) >= (value.startColumn as number);
}

function isAxis(value: unknown): value is ChartAxis {
  if (!isRecord(value)) return false;
  return typeof value.id === 'string'
    && ['top', 'bottom', 'left', 'right'].includes(String(value.position))
    && (value.scale === undefined || value.scale === 'linear' || value.scale === 'logarithmic')
    && (value.minimum === undefined || typeof value.minimum === 'number')
    && (value.maximum === undefined || typeof value.maximum === 'number')
    && (value.majorUnit === undefined || typeof value.majorUnit === 'number');
}

function isSeries(value: unknown): value is ChartSeries {
  if (!isRecord(value)) return false;
  const series = value as Record<string, unknown>;
  return typeof value.name === 'string'
    && isRange(series.range)
    && (series.xRange === undefined || isRange(series.xRange))
    && (series.yRange === undefined || isRange(series.yRange))
    && (series.chartType === undefined || ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter'].includes(String(series.chartType)))
    && (series.axis === undefined || series.axis === 'primary' || series.axis === 'secondary')
    && (series.smooth === undefined || typeof series.smooth === 'boolean');
}

function isChartPayload(value: unknown): value is ChartPayload {
  if (!isRecord(value) || value.kind !== 'chart') return false;
  const payload = value as Record<string, unknown>;
  const sourceRanges = payload.sourceRanges;
  const series = payload.series;
  return typeof payload.chartId === 'string'
    && ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'combo'].includes(String(value.chartType))
    && Array.isArray(sourceRanges)
    && sourceRanges.every(isRange)
    && (series === undefined || (Array.isArray(series) && series.every(isSeries)))
    && (payload.categoryRange === undefined || isRange(payload.categoryRange))
    && (payload.categoryAxis === undefined || isAxis(payload.categoryAxis))
    && (payload.valueAxis === undefined || isAxis(payload.valueAxis))
    && (payload.secondaryCategoryAxis === undefined || isAxis(payload.secondaryCategoryAxis))
    && (payload.secondaryValueAxis === undefined || isAxis(payload.secondaryValueAxis));
}

function validateChartPair(sheet: WorksheetModel, drawing: DrawingObject, payload: ChartPayload): void {
  if (drawing.kind !== 'chart' || payload.kind !== 'chart') throw new Error(`Chart pair kind mismatch: ${drawing.id}`);
  if (drawing.sheetId !== sheet.id || payload.chartId !== drawing.payloadId) throw new Error(`Chart pair identity mismatch: ${drawing.id}`);
  if (!isChartPayload(payload as unknown)) throw new Error('Invalid chart payload');
  if (sheet.drawings.some((entry) => entry.id === drawing.id)) throw new Error(`Drawing already exists: ${drawing.id}`);
  if (sheet.drawingPayloads.has(drawing.payloadId)) throw new Error(`Chart payload already exists: ${drawing.payloadId}`);
}

function findChartDrawing(sheet: WorksheetModel, chartId: string): { drawing: DrawingObject; payload: ChartPayload } | undefined {
  const drawing = sheet.drawings.find((entry) => entry.kind === 'chart' && entry.payloadId === chartId);
  if (!drawing) return undefined;
  const payload = sheet.drawingPayloads.get(drawing.payloadId);
  if (!payload || payload.kind !== 'chart') throw new Error(`Missing chart payload: ${chartId}`);
  if (payload.chartId !== chartId) throw new Error(`Chart payload identity mismatch: ${chartId}`);
  return { drawing, payload: payload as ChartPayload };
}

function addChartDrawing(sheet: WorksheetModel, params: ChartInsertParams): void {
  validateChartPair(sheet, params.drawing, params.payload);
  sheet.drawings.push(structuredClone(params.drawing));
  sheet.drawingPayloads.set(params.drawing.payloadId, structuredClone(params.payload));
}

function removeChartDrawing(sheet: WorksheetModel, chartId: string): { drawing: DrawingObject; payload: ChartPayload } {
  const current = findChartDrawing(sheet, chartId);
  if (!current) throw new Error(`Unknown chart: ${chartId}`);
  sheet.drawings.splice(sheet.drawings.findIndex((entry) => entry.id === current.drawing.id), 1);
  sheet.drawingPayloads.delete(current.drawing.payloadId);
  return { drawing: structuredClone(current.drawing), payload: structuredClone(current.payload) };
}

function updateChartPayload(sheet: WorksheetModel, params: { sheetId: string; chartId: string; payload: ChartPayload }): void {
  const current = findChartDrawing(sheet, params.chartId);
  if (!current) throw new Error(`Unknown chart: ${params.chartId}`);
  if (params.payload.chartId !== params.chartId) throw new Error(`Chart payload identity mismatch: ${params.chartId}`);
  sheet.drawingPayloads.set(current.drawing.payloadId, structuredClone(params.payload));
}

interface DrawingRemoveMutationParams {
  sheetId: string;
  drawingId: string;
}

function executeChartInsert(params: ChartInsertParams, context: CommandContext, expectedType?: ChartType): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  if (expectedType && params.payload.chartType !== expectedType) throw new Error(`Chart command type mismatch: expected ${expectedType}`);
  const sheet = context.workbook.getSheet(params.sheetId);
  validateChartPair(sheet, params.drawing, params.payload);
  const affectedRanges = sheetRange(params.sheetId);
  applyTrackedMutation<ChartInsertParams, DrawingRemoveMutationParams>(context, {
    id: 'drawing.add',
    sheetId: params.sheetId,
    params,
    inverseId: 'drawing.remove',
    inverseParams: { sheetId: params.sheetId, drawingId: params.drawing.id },
    affectedRanges,
    apply: () => addChartDrawing(context.workbook.getSheet(params.sheetId), params),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

function executeChartUpdate<P extends { sheetId: string; chartId: string }>(
  params: P,
  context: CommandContext,
  patch: (payload: ChartPayload, params: P) => ChartPayload,
): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  const sheet = context.workbook.getSheet(params.sheetId);
  const current = findChartDrawing(sheet, params.chartId);
  if (!current) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
  const nextPayload = patch(structuredClone(current.payload), params);
  if (!isChartPayload(nextPayload)) throw new Error(`Invalid chart payload: ${params.chartId}`);
  const affectedRanges = sheetRange(params.sheetId);
  const mutationParams = { sheetId: params.sheetId, payloadId: params.chartId, before: current.payload, after: nextPayload };
  applyTrackedMutation(context, {
    id: 'drawing.payload.update',
    sheetId: params.sheetId,
    params: mutationParams,
    inverseParams: { sheetId: params.sheetId, payloadId: params.chartId, before: nextPayload, after: current.payload },
    affectedRanges,
    apply: () => updateChartPayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, chartId: params.chartId, payload: nextPayload }),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

export function registerChartCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  runtime.registry.registerCommand<ChartInsertParams>({ id: 'chart.insert', execute: (params, context) => executeChartInsert(params, context) });
  commandIds.push('chart.insert');

  for (const chartType of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'] as const) {
    const id = `chart.insert.${chartType}`;
    runtime.registry.registerCommand<ChartInsertParams>({ id, execute: (params, context) => executeChartInsert(params, context, chartType) });
    commandIds.push(id);
  }

  runtime.registry.registerCommand<ChartUpdateParams>({ id: 'chart.update', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, ...input.payload, kind: 'chart', chartId: payload.chartId })) });
  commandIds.push('chart.update');
  runtime.registry.registerCommand<ChartSetTypeParams>({ id: 'chart.setType', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, chartType: input.chartType, stacked: input.stacked ?? payload.stacked })) });
  commandIds.push('chart.setType');
  runtime.registry.registerCommand<ChartSetLegendParams>({ id: 'chart.setLegend', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, legendPosition: input.legendPosition })) });
  commandIds.push('chart.setLegend');
  runtime.registry.registerCommand<ChartSetDataLabelsParams>({ id: 'chart.setDataLabels', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, showDataLabels: input.showDataLabels })) });
  commandIds.push('chart.setDataLabels');
  runtime.registry.registerCommand<ChartSetSeriesParams>({ id: 'chart.setSeries', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, sourceRanges: structuredClone(input.sourceRanges), series: input.series ? structuredClone(input.series) : payload.series, categoryRange: input.categoryRange ? structuredClone(input.categoryRange) : payload.categoryRange })) });
  commandIds.push('chart.setSeries');
  runtime.registry.registerCommand<ChartSetAxesParams>({ id: 'chart.setAxes', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, categoryAxis: input.categoryAxis ? structuredClone(input.categoryAxis) : payload.categoryAxis, valueAxis: input.valueAxis ? structuredClone(input.valueAxis) : payload.valueAxis, secondaryCategoryAxis: input.secondaryCategoryAxis ? structuredClone(input.secondaryCategoryAxis) : payload.secondaryCategoryAxis, secondaryValueAxis: input.secondaryValueAxis ? structuredClone(input.secondaryValueAxis) : payload.secondaryValueAxis })) });
  commandIds.push('chart.setAxes');
  runtime.registry.registerCommand<ChartSetSecondaryAxisParams>({ id: 'chart.setSecondaryAxis', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, series: (payload.series ?? []).map((series) => series.name === input.seriesName ? { ...series, axis: input.enabled ? 'secondary' : 'primary' } : series) })) });
  commandIds.push('chart.setSecondaryAxis');

  runtime.registry.registerCommand<ChartRemoveParams>({
    id: 'chart.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const current = findChartDrawing(sheet, params.chartId);
      if (!current) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const affectedRanges = sheetRange(params.sheetId);
      const inverseParams: ChartInsertParams = { sheetId: params.sheetId, drawing: structuredClone(current.drawing), payload: structuredClone(current.payload) };
      applyTrackedMutation<DrawingRemoveMutationParams, ChartInsertParams>(context, {
        id: 'drawing.remove',
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, drawingId: current.drawing.id },
        inverseId: 'drawing.add',
        inverseParams,
        affectedRanges,
        apply: () => removeChartDrawing(context.workbook.getSheet(params.sheetId), params.chartId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('chart.remove');
  return commandIds;
}

export const CHART_MUTATION_IDS = [] as const;

export const CHART_COMMAND_IDS = [
  'chart.insert', 'chart.update', 'chart.remove', 'chart.setType', 'chart.setLegend',
  'chart.setDataLabels', 'chart.setSeries', 'chart.setAxes', 'chart.setSecondaryAxis',
  'chart.insert.column', 'chart.insert.bar', 'chart.insert.line', 'chart.insert.area',
  'chart.insert.pie', 'chart.insert.doughnut', 'chart.insert.scatter', 'chart.insert.combo',
] as const;
