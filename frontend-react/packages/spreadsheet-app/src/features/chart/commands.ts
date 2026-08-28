import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import { chartStackingForSubtype, isChartSubtypeForType, type ChartAxisModel, type ChartDrawingPayload, type ChartSeriesModel, type ChartSource, type ChartSubtype, type DrawingObject, type RangeRef, type WorksheetModel } from '@react-sheets/core-model';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

export type ChartType = ChartDrawingPayload['chartType'];
export type ChartAxis = ChartAxisModel;
export type ChartSeries = ChartSeriesModel;

/** Canonical chart payload. It remains a DrawingPayload member and is stored by payloadId. */
export type ChartPayload = Omit<ChartDrawingPayload, 'series'> & {
  series?: ChartSeries[];
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
  subtype: ChartSubtype;
  stacked?: ChartPayload['stacked'];
}

export interface ChartSetLegendParams {
  sheetId: string;
  chartId: string;
  legendPosition: NonNullable<ChartPayload['elements']['legend']>['position'];
}

export interface ChartSetDataLabelsParams {
  sheetId: string;
  chartId: string;
  showDataLabels: boolean;
}

export interface ChartSetSeriesParams {
  sheetId: string;
  chartId: string;
  source: ChartSource;
  series?: ChartPayload['series'];
  categoryRange?: RangeRef;
}

export interface ChartSelectDataParams extends ChartSetSeriesParams {
  switchRowColumn?: boolean;
}

export interface ChartSeriesAddParams {
  sheetId: string;
  chartId: string;
  series: ChartSeries;
}

export interface ChartSeriesUpdateParams {
  sheetId: string;
  chartId: string;
  seriesId: string;
  series: Partial<ChartSeries>;
}

export interface ChartSeriesRemoveParams {
  sheetId: string;
  chartId: string;
  seriesId: string;
}

export interface ChartSeriesMoveParams {
  sheetId: string;
  chartId: string;
  seriesId: string;
  direction: 'up' | 'down';
}

export interface ChartSetAxesParams {
  sheetId: string;
  chartId: string;
  categoryAxis?: ChartAxis;
  valueAxis?: ChartAxis;
  secondaryCategoryAxis?: ChartAxis;
  secondaryValueAxis?: ChartAxis;
}

export interface ChartSetElementsParams {
  sheetId: string;
  chartId: string;
  elements: Partial<ChartPayload['elements']>;
}

export interface ChartSetSeriesStyleParams {
  sheetId: string;
  chartId: string;
  seriesName: string;
  style: Partial<Pick<ChartSeries, 'color' | 'chartType' | 'axis' | 'smooth' | 'marker' | 'dataLabels' | 'trendlines' | 'errorBars'>>;
}

export interface ChartSetDataTableParams {
  sheetId: string;
  chartId: string;
  dataTable: NonNullable<ChartPayload['elements']['dataTable']>;
}

export interface ChartSetTrendlinesParams {
  sheetId: string;
  chartId: string;
  seriesId: string;
  trendlines: NonNullable<ChartSeries['trendlines']>;
}

export interface ChartSetErrorBarsParams {
  sheetId: string;
  chartId: string;
  seriesId: string;
  errorBars?: ChartSeries['errorBars'];
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

const CHART_TYPES: readonly ChartType[] = [
  'column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'bubble',
  'treemap', 'sunburst', 'histogram', 'pareto', 'box-whisker', 'waterfall',
  'funnel', 'stock', 'surface', 'radar', 'map', 'combo',
];

function isChartSource(value: unknown): value is ChartSource {
  if (!isRecord(value)) return false;
  if (value.kind === 'worksheet-ranges') return Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange)
    && (value.identity === undefined || typeof value.identity === 'string')
    && (value.dynamic === undefined || typeof value.dynamic === 'boolean');
  if (value.kind === 'pivot') return typeof value.pivotId === 'string' && value.pivotId.length > 0;
  if (value.kind === 'table') return typeof value.tableId === 'string' && value.tableId.length > 0 && isRecord(value.bindings)
    && (value.structuredReference === undefined || typeof value.structuredReference === 'string');
  return value.kind === 'report-range' && isRange(value.range) && isRecord(value.bindings)
    && (value.identity === undefined || typeof value.identity === 'string');
}

function isAxis(value: unknown): value is ChartAxis {
  if (!isRecord(value)) return false;
  return typeof value.id === 'string'
    && ['top', 'bottom', 'left', 'right'].includes(String(value.position))
    && (value.scale === undefined || value.scale === 'linear' || value.scale === 'logarithmic')
    && (value.visible === undefined || typeof value.visible === 'boolean')
    && (value.minimum === undefined || typeof value.minimum === 'number')
    && (value.maximum === undefined || typeof value.maximum === 'number')
    && (value.majorUnit === undefined || typeof value.majorUnit === 'number')
    && (value.minorUnit === undefined || typeof value.minorUnit === 'number')
    && (value.axisType === undefined || ['category', 'value', 'date'].includes(String(value.axisType)))
    && (value.logBase === undefined || (typeof value.logBase === 'number' && value.logBase > 1))
    && (value.automaticMinimum === undefined || typeof value.automaticMinimum === 'boolean')
    && (value.automaticMaximum === undefined || typeof value.automaticMaximum === 'boolean')
    && (value.reverseOrder === undefined || typeof value.reverseOrder === 'boolean')
    && (value.crosses === undefined || ['automatic', 'value', 'maximum'].includes(String(value.crosses)))
    && (value.crossBetween === undefined || ['mid-category', 'between'].includes(String(value.crossBetween)))
    && (value.labelAngle === undefined || typeof value.labelAngle === 'number')
    && (value.labelInterval === undefined || Number.isSafeInteger(value.labelInterval))
    && (value.markInterval === undefined || Number.isSafeInteger(value.markInterval))
    && (value.displayUnits === undefined || ['none', 'hundreds', 'thousands', 'ten-thousands', 'millions', 'billions', 'trillions'].includes(String(value.displayUnits)))
    && (value.majorGridlines === undefined || isGridlines(value.majorGridlines))
    && (value.minorGridlines === undefined || isGridlines(value.minorGridlines));
}

function isGridlines(value: unknown): boolean {
  return isRecord(value)
    && typeof value.visible === 'boolean'
    && (value.width === undefined || typeof value.width === 'number')
    && (value.dash === undefined || ['solid', 'dash', 'dot'].includes(String(value.dash)));
}

function isElements(value: unknown): value is ChartPayload['elements'] {
  if (!isRecord(value)) return false;
  const legend = value.legend;
  return (value.title === undefined || typeof value.title === 'string')
    && (legend === undefined || (isRecord(legend) && typeof legend.visible === 'boolean' && ['top', 'bottom', 'left', 'right', 'top-right'].includes(String(legend.position))))
    && (value.dataLabels === undefined || isDataLabels(value.dataLabels))
    && (value.categoryAxis === undefined || isAxis(value.categoryAxis))
    && (value.valueAxis === undefined || isAxis(value.valueAxis))
    && (value.secondaryCategoryAxis === undefined || isAxis(value.secondaryCategoryAxis))
    && (value.secondaryValueAxis === undefined || isAxis(value.secondaryValueAxis))
    && (value.dataTable === undefined || isRecord(value.dataTable))
    && (value.hiddenData === 'show' || value.hiddenData === 'hideRows' || value.hiddenData === 'hideColumns');
}

function isDataLabels(value: unknown): boolean {
  return isRecord(value)
    && typeof value.visible === 'boolean'
    && (value.position === undefined || ['best-fit', 'center', 'inside-end', 'inside-base', 'outside-end', 'above', 'below', 'left', 'right'].includes(String(value.position)))
    && (value.target === undefined || ['chart', 'series', 'point'].includes(String(value.target)))
    && (value.valuesFromCells === undefined || isRange(value.valuesFromCells));
}

function isTrendline(value: unknown): boolean {
  return isRecord(value)
    && ['linear', 'exponential', 'logarithmic', 'polynomial', 'power', 'moving-average'].includes(String(value.type))
    && (value.order === undefined || (Number.isSafeInteger(value.order) && Number(value.order) >= 2))
    && (value.period === undefined || (Number.isSafeInteger(value.period) && Number(value.period) > 0))
    && ['displayEquation', 'displayRSquared'].every((key) => value[key] === undefined || typeof value[key] === 'boolean');
}

function isErrorBars(value: unknown): boolean {
  return isRecord(value)
    && ['fixed', 'percentage', 'standard-deviation', 'standard-error', 'custom'].includes(String(value.type))
    && (value.direction === undefined || ['vertical', 'horizontal', 'both'].includes(String(value.direction)))
    && (value.endStyle === undefined || ['cap', 'no-cap'].includes(String(value.endStyle)))
    && (value.value === undefined || (typeof value.value === 'number' && Number.isFinite(value.value)))
    && (value.plusValue === undefined || (typeof value.plusValue === 'number' && Number.isFinite(value.plusValue)))
    && (value.minusValue === undefined || (typeof value.minusValue === 'number' && Number.isFinite(value.minusValue)))
    && (value.plusRange === undefined || isRange(value.plusRange))
    && (value.minusRange === undefined || isRange(value.minusRange))
    && (value.type !== 'custom' || isRange(value.plusRange) && isRange(value.minusRange));
}

function isStockRoles(value: unknown): boolean {
  if (!isRecord(value) || !isRange(value.high) || !isRange(value.low) || !isRange(value.close)) return false;
  return (value.open === undefined || isRange(value.open)) && (value.volume === undefined || isRange(value.volume));
}

function isNativeIdentity(value: unknown): boolean {
  return isRecord(value)
    && typeof value.family === 'string'
    && typeof value.subtype === 'string'
    && (value.xlChartType === undefined || Number.isSafeInteger(value.xlChartType))
    && (value.part === undefined || typeof value.part === 'string')
    && (value.status === undefined || value.status === 'owned' || value.status === 'preserved-native');
}

function isSeries(value: unknown): value is ChartSeries {
  if (!isRecord(value)) return false;
  const series = value as Record<string, unknown>;
  return typeof value.name === 'string'
    && isRange(series.range)
    && (series.xRange === undefined || isRange(series.xRange))
    && (series.yRange === undefined || isRange(series.yRange))
    && (series.sizeRange === undefined || isRange(series.sizeRange))
    && (series.categoryRange === undefined || isRange(series.categoryRange))
    && (series.chartType === undefined || CHART_TYPES.includes(series.chartType as ChartType))
    && (series.subtype === undefined || typeof series.subtype === 'string')
    && (series.axis === undefined || series.axis === 'primary' || series.axis === 'secondary')
    && (series.smooth === undefined || typeof series.smooth === 'boolean')
    && (series.marker === undefined || isRecord(series.marker))
    && (series.dataLabels === undefined || isDataLabels(series.dataLabels))
    && (series.trendlines === undefined || (Array.isArray(series.trendlines) && series.trendlines.every(isTrendline)))
    && (series.errorBars === undefined || isErrorBars(series.errorBars))
    && (series.pointOverrides === undefined || isRecord(series.pointOverrides))
    && (series.visible === undefined || typeof series.visible === 'boolean')
    && (series.gapWidth === undefined || typeof series.gapWidth === 'number')
    && (series.overlap === undefined || typeof series.overlap === 'number')
    && (series.invertIfNegative === undefined || typeof series.invertIfNegative === 'boolean')
    && (series.stockRoles === undefined || isStockRoles(series.stockRoles));
}

function isChartPayload(value: unknown): value is ChartPayload {
  if (!isRecord(value) || value.kind !== 'chart') return false;
  const payload = value as Record<string, unknown>;
  const series = payload.series;
  return typeof payload.chartId === 'string'
    && CHART_TYPES.includes(value.chartType as ChartType)
    && typeof payload.subtype === 'string'
    && isChartSubtypeForType(value.chartType as ChartType, payload.subtype as ChartSubtype)
    && isChartSource(payload.source)
    && (series === undefined || (Array.isArray(series) && series.every(isSeries)))
    && isElements(payload.elements)
    && (payload.categoryRange === undefined || isRange(payload.categoryRange))
    && (payload.nativeIdentity === undefined || isNativeIdentity(payload.nativeIdentity))
    && (payload.histogramOptions === undefined || isRecord(payload.histogramOptions))
    && (payload.boxWhiskerOptions === undefined || isRecord(payload.boxWhiskerOptions))
    && (payload.waterfallOptions === undefined || isRecord(payload.waterfallOptions))
    && (payload.mapOptions === undefined || isRecord(payload.mapOptions))
    && (payload.dataOrientation === undefined || payload.dataOrientation === 'rows' || payload.dataOrientation === 'columns')
    && (payload.chartType !== 'combo' || (Array.isArray(series) && series.every((entry) => entry.chartType !== undefined)));
}

function validateChartSemantics(payload: ChartPayload): void {
  if (payload.nativeIdentity?.status === 'preserved-native') throw new Error(`UNSUPPORTED_FEATURE: Preserved-native chart ${payload.chartId} has no editable canonical owner`);
  if (payload.chartType === 'combo') {
    if (!payload.series?.length || payload.series.some((series) => !series.chartType)) throw new Error('INVALID_CHART_SOURCE: Combo charts require an explicit type for every series');
    for (const series of payload.series) {
      if (series.chartType && !['column', 'bar', 'line', 'area'].includes(series.chartType)) throw new Error(`INVALID_CHART_SOURCE: Combo series type ${series.chartType} is not supported by the canonical combo layout`);
      if (series.subtype && series.chartType && !isChartSubtypeForType(series.chartType, series.subtype)) throw new Error(`Chart series subtype ${series.subtype} does not belong to ${series.chartType}`);
    }
  }
  if (payload.chartType !== 'combo' && payload.series?.some((series) => series.chartType && series.chartType !== payload.chartType)) {
    throw new Error(`INVALID_CHART_SOURCE: ${payload.chartType} chart cannot contain a different series chart type`);
  }
  for (const series of payload.series ?? []) {
    const seriesType = series.chartType ?? payload.chartType;
    if (seriesType === 'stock' && !series.stockRoles) throw new Error('INVALID_CHART_SOURCE: Stock charts require explicit High/Low/Close role bindings');
    if ((seriesType === 'scatter' || seriesType === 'bubble') && (!series.xRange || !series.yRange)) throw new Error(`INVALID_CHART_SOURCE: ${seriesType} charts require explicit X/Y range bindings`);
    if (seriesType === 'bubble' && !series.sizeRange) throw new Error('INVALID_CHART_SOURCE: Bubble charts require an independent Size range binding');
    if (series.errorBars?.type === 'custom' && (!series.errorBars.plusRange || !series.errorBars.minusRange)) throw new Error('INVALID_CHART_SOURCE: Custom error bars require explicit plus and minus ranges');
  }
  if (payload.dataOrientation === 'rows' && payload.series?.some((series) => series.range.startRow === series.range.endRow)) throw new Error('INVALID_CHART_SOURCE: Row-oriented chart series must contain at least one data column');
}

function validateChartPair(sheet: WorksheetModel, drawing: DrawingObject, payload: ChartPayload): void {
  if (drawing.kind !== 'chart' || payload.kind !== 'chart') throw new Error(`Chart pair kind mismatch: ${drawing.id}`);
  if (drawing.sheetId !== sheet.id || payload.chartId !== drawing.payloadId) throw new Error(`Chart pair identity mismatch: ${drawing.id}`);
  if (!isChartPayload(payload as unknown)) throw new Error('Invalid chart payload');
  validateChartSemantics(payload);
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

function executeChartInsert(params: ChartInsertParams, context: CommandContext, expectedType?: ChartType): { operationId: string; mutationCount: number; affectedRanges: ReturnType<typeof sheetRange> } {
  if (expectedType && params.payload.chartType !== expectedType) throw new Error(`Chart command type mismatch: expected ${expectedType}`);
  const sheet = context.workbook.getSheet(params.sheetId);
  validateChartPair(sheet, params.drawing, params.payload);
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: 'drawing.add',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{ id: 'drawing.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, drawingId: params.drawing.id }, affectedRanges }],
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
  if (!current) throw new Error(`Unknown chart: ${params.chartId}`);
  const nextPayload = patch(structuredClone(current.payload), params);
  if (!isChartPayload(nextPayload)) throw new Error(`Invalid chart payload: ${params.chartId}`);
  validateChartSemantics(nextPayload);
  const affectedRanges = sheetRange(params.sheetId);
  const mutationParams = { sheetId: params.sheetId, payloadId: params.chartId, before: current.payload, after: nextPayload };
  context.applyMutation({
    id: 'drawing.payload.update',
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params: mutationParams,
    affectedRanges,
    inverse: [{ id: 'drawing.payload.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, payloadId: params.chartId, before: nextPayload, after: current.payload }, affectedRanges }],
    apply: () => updateChartPayload(context.workbook.getSheet(params.sheetId), { sheetId: params.sheetId, chartId: params.chartId, payload: nextPayload }),
  });
  return { operationId: context.operationId, mutationCount: 1, affectedRanges };
}

export function registerChartCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  runtime.registry.registerCommand<ChartInsertParams>({ id: 'chart.insert', execute: (params, context) => executeChartInsert(params, context) });
  commandIds.push('chart.insert');

  for (const chartType of CHART_TYPES) {
    const id = `chart.insert.${chartType}`;
    runtime.registry.registerCommand<ChartInsertParams>({ id, execute: (params, context) => executeChartInsert(params, context, chartType) });
    commandIds.push(id);
  }

  runtime.registry.registerCommand<ChartUpdateParams>({ id: 'chart.update', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, ...input.payload, kind: 'chart', chartId: payload.chartId })) });
  commandIds.push('chart.update');
  runtime.registry.registerCommand<ChartSetTypeParams>({ id: 'chart.setType', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    if (!isChartSubtypeForType(input.chartType, input.subtype)) throw new Error(`Chart subtype ${input.subtype} does not belong to ${input.chartType}`);
    const next = { ...payload, chartType: input.chartType, subtype: input.subtype, stacked: input.stacked ?? chartStackingForSubtype(input.subtype) };
    if (next.stacked === undefined) delete next.stacked;
    return next;
  }) });
  commandIds.push('chart.setType');
  runtime.registry.registerCommand<ChartSetLegendParams>({ id: 'chart.setLegend', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, elements: { ...payload.elements, legend: { visible: true, position: input.legendPosition } } })) });
  commandIds.push('chart.setLegend');
  runtime.registry.registerCommand<ChartSetDataLabelsParams>({ id: 'chart.setDataLabels', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, elements: { ...payload.elements, dataLabels: { ...(payload.elements.dataLabels ?? { visible: false }), visible: input.showDataLabels } } })) });
  commandIds.push('chart.setDataLabels');
  runtime.registry.registerCommand<ChartSetSeriesParams>({ id: 'chart.setSeries', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, source: structuredClone(input.source), series: input.series ? structuredClone(input.series) : payload.series, categoryRange: input.categoryRange ? structuredClone(input.categoryRange) : payload.categoryRange })) });
  commandIds.push('chart.setSeries');
  runtime.registry.registerCommand<ChartSelectDataParams>({ id: 'chart.selectData', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({
    ...payload,
    source: structuredClone(input.source),
    series: input.series ? structuredClone(input.series) : payload.series,
    categoryRange: input.categoryRange ? structuredClone(input.categoryRange) : payload.categoryRange,
    ...(input.switchRowColumn === undefined ? {} : { dataOrientation: input.switchRowColumn ? payload.dataOrientation === 'rows' ? 'columns' : 'rows' : payload.dataOrientation ?? 'columns' }),
  })) });
  commandIds.push('chart.selectData');
  runtime.registry.registerCommand<ChartSeriesAddParams>({ id: 'chart.series.add', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const series = [...(payload.series ?? [])];
    if (series.some((entry) => entry.id === input.series.id)) throw new Error(`Duplicate chart series: ${input.series.id}`);
    return { ...payload, series: [...series, structuredClone(input.series)] };
  }) });
  commandIds.push('chart.series.add');
  runtime.registry.registerCommand<ChartSeriesUpdateParams>({ id: 'chart.series.update', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const index = (payload.series ?? []).findIndex((entry) => entry.id === input.seriesId);
    if (index < 0) throw new Error(`Unknown chart series: ${input.seriesId}`);
    const series = [...(payload.series ?? [])];
    series[index] = { ...series[index]!, ...structuredClone(input.series), id: input.seriesId };
    return { ...payload, series };
  }) });
  commandIds.push('chart.series.update');
  runtime.registry.registerCommand<ChartSeriesRemoveParams>({ id: 'chart.series.remove', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const series = payload.series ?? [];
    if (!series.some((entry) => entry.id === input.seriesId)) throw new Error(`Unknown chart series: ${input.seriesId}`);
    if (series.length <= 1) throw new Error('INVALID_CHART_SOURCE: A chart must retain at least one series');
    return { ...payload, series: series.filter((entry) => entry.id !== input.seriesId) };
  }) });
  commandIds.push('chart.series.remove');
  runtime.registry.registerCommand<ChartSeriesMoveParams>({ id: 'chart.series.move', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const series = [...(payload.series ?? [])];
    const index = series.findIndex((entry) => entry.id === input.seriesId);
    if (index < 0) throw new Error(`Unknown chart series: ${input.seriesId}`);
    const target = input.direction === 'up' ? index - 1 : index + 1;
    if (target < 0 || target >= series.length) return payload;
    [series[index], series[target]] = [series[target]!, series[index]!];
    return { ...payload, series };
  }) });
  commandIds.push('chart.series.move');
  runtime.registry.registerCommand<ChartSetAxesParams>({ id: 'chart.setAxes', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, elements: { ...payload.elements, categoryAxis: input.categoryAxis ? structuredClone(input.categoryAxis) : payload.elements.categoryAxis, valueAxis: input.valueAxis ? structuredClone(input.valueAxis) : payload.elements.valueAxis, secondaryCategoryAxis: input.secondaryCategoryAxis ? structuredClone(input.secondaryCategoryAxis) : payload.elements.secondaryCategoryAxis, secondaryValueAxis: input.secondaryValueAxis ? structuredClone(input.secondaryValueAxis) : payload.elements.secondaryValueAxis } })) });
  commandIds.push('chart.setAxes');
  runtime.registry.registerCommand<ChartSetSecondaryAxisParams>({ id: 'chart.setSecondaryAxis', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    let found = false;
    const series: ChartSeries[] = (payload.series ?? []).map((entry) => entry.name === input.seriesName ? (found = true, { ...entry, axis: (input.enabled ? 'secondary' : 'primary') as 'primary' | 'secondary' }) : entry);
    if (!found) throw new Error(`Unknown chart series: ${input.seriesName}`);
    return { ...payload, series };
  }) });
  commandIds.push('chart.setSecondaryAxis');
  runtime.registry.registerCommand<ChartSetElementsParams>({ id: 'chart.setElements', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, elements: { ...payload.elements, ...structuredClone(input.elements), legend: input.elements.legend ? structuredClone(input.elements.legend) : payload.elements.legend, dataLabels: input.elements.dataLabels ? structuredClone(input.elements.dataLabels) : payload.elements.dataLabels, categoryAxis: input.elements.categoryAxis ? structuredClone(input.elements.categoryAxis) : payload.elements.categoryAxis, valueAxis: input.elements.valueAxis ? structuredClone(input.elements.valueAxis) : payload.elements.valueAxis, secondaryCategoryAxis: input.elements.secondaryCategoryAxis ? structuredClone(input.elements.secondaryCategoryAxis) : payload.elements.secondaryCategoryAxis, secondaryValueAxis: input.elements.secondaryValueAxis ? structuredClone(input.elements.secondaryValueAxis) : payload.elements.secondaryValueAxis, plotArea: input.elements.plotArea ? structuredClone(input.elements.plotArea) : payload.elements.plotArea, chartArea: input.elements.chartArea ? structuredClone(input.elements.chartArea) : payload.elements.chartArea } })) });
  commandIds.push('chart.setElements');
  runtime.registry.registerCommand<ChartSetSeriesStyleParams>({ id: 'chart.setSeriesStyle', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    let found = false;
    const series = (payload.series ?? []).map((entry) => {
      if (entry.name !== input.seriesName && entry.id !== input.seriesName) return entry;
      found = true;
      const style = input.style;
      return {
        ...entry,
        ...(style.color === undefined ? {} : { color: style.color }),
        ...(style.chartType === undefined ? {} : { chartType: style.chartType }),
        ...(style.axis === undefined ? {} : { axis: style.axis }),
        ...(style.smooth === undefined ? {} : { smooth: style.smooth }),
        ...(style.marker === undefined ? {} : { marker: structuredClone(style.marker) }),
        ...(style.dataLabels === undefined ? {} : { dataLabels: structuredClone(style.dataLabels) }),
        ...(style.trendlines === undefined ? {} : { trendlines: structuredClone(style.trendlines) }),
        ...(style.errorBars === undefined ? {} : { errorBars: structuredClone(style.errorBars) }),
      };
    });
    if (!found) throw new Error(`Unknown chart series: ${input.seriesName}`);
    return { ...payload, series };
  }) });
  commandIds.push('chart.setSeriesStyle');
  runtime.registry.registerCommand<ChartSetDataTableParams>({ id: 'chart.setDataTable', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => ({ ...payload, elements: { ...payload.elements, dataTable: structuredClone(input.dataTable) } })) });
  commandIds.push('chart.setDataTable');
  runtime.registry.registerCommand<ChartSetTrendlinesParams>({ id: 'chart.setTrendlines', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const series = (payload.series ?? []).map((entry) => entry.id === input.seriesId ? { ...entry, trendlines: structuredClone(input.trendlines) } : entry);
    if (!(payload.series ?? []).some((entry) => entry.id === input.seriesId)) throw new Error(`Unknown chart series: ${input.seriesId}`);
    return { ...payload, series };
  }) });
  commandIds.push('chart.setTrendlines');
  runtime.registry.registerCommand<ChartSetErrorBarsParams>({ id: 'chart.setErrorBars', execute: (params, context) => executeChartUpdate(params, context, (payload, input) => {
    const series = (payload.series ?? []).map((entry) => entry.id === input.seriesId ? { ...entry, ...(input.errorBars ? { errorBars: structuredClone(input.errorBars) } : { errorBars: undefined }) } : entry);
    if (!(payload.series ?? []).some((entry) => entry.id === input.seriesId)) throw new Error(`Unknown chart series: ${input.seriesId}`);
    return { ...payload, series };
  }) });
  commandIds.push('chart.setErrorBars');

  runtime.registry.registerCommand<ChartRemoveParams>({
    id: 'chart.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const current = findChartDrawing(sheet, params.chartId);
      if (!current) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const affectedRanges = sheetRange(params.sheetId);
      const inverseParams: ChartInsertParams = { sheetId: params.sheetId, drawing: structuredClone(current.drawing), payload: structuredClone(current.payload) };
      context.applyMutation({
        id: 'drawing.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, drawingId: current.drawing.id },
        affectedRanges,
        inverse: [{ id: 'drawing.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverseParams, affectedRanges }],
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
  'chart.setElements', 'chart.setSeriesStyle', 'chart.selectData', 'chart.setDataTable',
  'chart.setTrendlines', 'chart.setErrorBars', 'chart.series.add', 'chart.series.update',
  'chart.series.remove', 'chart.series.move',
  'chart.insert.column', 'chart.insert.bar', 'chart.insert.line', 'chart.insert.area',
  'chart.insert.pie', 'chart.insert.doughnut', 'chart.insert.scatter', 'chart.insert.bubble',
  'chart.insert.treemap', 'chart.insert.sunburst', 'chart.insert.histogram', 'chart.insert.pareto',
  'chart.insert.box-whisker', 'chart.insert.waterfall', 'chart.insert.funnel', 'chart.insert.stock',
  'chart.insert.surface', 'chart.insert.radar', 'chart.insert.map', 'chart.insert.combo',
] as const;
