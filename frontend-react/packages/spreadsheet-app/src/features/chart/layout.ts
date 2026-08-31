import type {
  ChartAxisModel,
  ChartBoxWhiskerOptions,
  ChartDrawingPayload,
  ChartHistogramOptions,
  ChartMapOptions,
  ChartSeriesModel,
  ChartSubtype,
  ChartTrendlineModel,
  ChartWaterfallOptions,
  PivotScalar,
} from '@react-sheets/core-model';
import { chartNumericValue, type ChartDataStatus, type ResolvedChartData, type ResolvedChartSeries } from './data';

export interface ChartLayoutPoint {
  index: number;
  category: PivotScalar;
  value: number | null;
  xValue?: number | null;
  sizeValue?: number | null;
  x: number;
  y: number;
  visible: boolean;
  errorPlus?: number;
  errorMinus?: number;
}

export interface ChartLayoutBar {
  index: number;
  category: PivotScalar;
  start: number;
  end: number;
  x: number;
  y: number;
  width: number;
  height: number;
  color: string;
  visible: boolean;
}

export interface ChartLayoutTrendline {
  model: ChartTrendlineModel;
  points: Array<{ x: number; y: number }>;
}

export interface ChartLayoutSeries {
  id: string;
  name: string;
  chartType: Exclude<ChartDrawingPayload['chartType'], 'combo'>;
  subtype?: ChartSubtype;
  axis: 'primary' | 'secondary';
  color: string;
  points: ChartLayoutPoint[];
  bars: ChartLayoutBar[];
  trendlines: ChartLayoutTrendline[];
  visible: boolean;
  smooth?: boolean;
}

export interface ChartAxisLayout {
  model: ChartAxisModel;
  minimum: number;
  maximum: number;
  ticks: number[];
}

export interface ChartPieSliceLayout {
  seriesIndex: number;
  pointIndex: number;
  value: number;
  startAngle: number;
  endAngle: number;
  innerRadius: number;
  outerRadius: number;
  explosion: number;
  color: string;
  label: string;
}

export interface ChartHistogramBinLayout {
  start: number;
  end: number;
  count: number;
  label: string;
}

export interface ChartBoxLayout {
  seriesIndex: number;
  minimum: number;
  lowerWhisker: number;
  q1: number;
  median: number;
  q3: number;
  upperWhisker: number;
  maximum: number;
  mean: number;
  outliers: number[];
  color: string;
}

export interface ChartWaterfallBarLayout {
  index: number;
  start: number;
  end: number;
  total: boolean;
  color: string;
}

export interface ChartLayout {
  status: ChartDataStatus;
  width: number;
  height: number;
  plot: { left: number; top: number; width: number; height: number };
  title?: { text: string; x: number; y: number };
  legend: { visible: boolean; position: NonNullable<ChartDrawingPayload['elements']['legend']>['position'] };
  categoryAxis?: ChartAxisLayout;
  valueAxis?: ChartAxisLayout;
  secondaryValueAxis?: ChartAxisLayout;
  series: ChartLayoutSeries[];
  kind: 'cartesian' | 'pie' | 'treemap' | 'sunburst' | 'histogram' | 'box-whisker' | 'waterfall' | 'funnel' | 'stock' | 'surface' | 'radar' | 'map';
  pieSlices?: ChartPieSliceLayout[];
  histogramBins?: ChartHistogramBinLayout[];
  paretoPoints?: Array<{ x: number; y: number }>;
  boxes?: ChartBoxLayout[];
  waterfallBars?: ChartWaterfallBarLayout[];
  funnelStages?: Array<{ index: number; value: number; nextValue: number; label: string; color: string }>;
  stockPoints?: Array<{ index: number; open?: number; high: number; low: number; close: number; volume?: number; color: string }>;
  surfaceCells?: Array<{ row: number; column: number; value: number; color: string }>;
  radar?: { count: number; maximum: number; points: Array<{ seriesIndex: number; values: number[]; color: string }> };
  map?: ChartMapOptions & { resolved: false; reason: string };
}

/** Revision-bound layout cache used by Canvas and print projections. */
export class ChartLayoutCache {
  private readonly entries = new Map<string, { key: string; layout: ChartLayout }>();

  resolve(payload: ChartDrawingPayload, data: ResolvedChartData, width: number, height: number): ChartLayout {
    const key = `${data.sourceRevision ?? 'unversioned'}:${width}:${height}:${stableLayoutValue(payload)}`;
    const current = this.entries.get(payload.chartId);
    if (current?.key === key) return structuredClone(current.layout);
    const layout = buildChartLayout(payload, data, width, height);
    if (layout.status.kind === 'ready') this.entries.set(payload.chartId, { key, layout: structuredClone(layout) });
    else this.entries.delete(payload.chartId);
    return layout;
  }

  invalidate(chartId?: string): void {
    if (chartId === undefined) this.entries.clear();
    else this.entries.delete(chartId);
  }
}

function stableLayoutValue(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value) ?? 'undefined';
  if (Array.isArray(value)) return `[${value.map(stableLayoutValue).join(',')}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${stableLayoutValue(record[key])}`).join(',')}}`;
}

const DEFAULT_COLORS = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4'];

function numberValues(values: readonly PivotScalar[]): number[] {
  return values.map(chartNumericValue).filter((value): value is number => value !== undefined);
}

function defaultAxis(id: string, position: ChartAxisModel['position'], axisType: ChartAxisModel['axisType']): ChartAxisModel {
  return { id, position, visible: true, axisType, scale: 'linear', majorTickMark: 'outside', minorTickMark: 'none', tickLabelPosition: 'next-to-axis' };
}

function axisBounds(model: ChartAxisModel, values: readonly number[], percent = false): ChartAxisLayout {
  const finite = values.filter(Number.isFinite);
  const dataMinimum = finite.length ? Math.min(...finite) : 0;
  const dataMaximum = finite.length ? Math.max(...finite) : 1;
  let minimum = model.minimum ?? (percent ? 0 : Math.min(0, dataMinimum));
  let maximum = model.maximum ?? (percent ? 100 : dataMaximum);
  if (model.minimum === undefined && !percent && minimum === maximum) minimum -= 1;
  if (model.maximum === undefined && !percent) {
    const span = Math.max(1, maximum - minimum);
    maximum += span * 0.1;
  }
  if (model.scale === 'logarithmic') {
    minimum = Math.max(Number.MIN_VALUE, minimum || Number.MIN_VALUE);
    maximum = Math.max(minimum * 10, maximum);
  }
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum <= minimum) throw new Error('INVALID_CHART_SOURCE: Axis bounds are not finite');
  const unit = model.majorUnit && model.majorUnit > 0 ? model.majorUnit : niceUnit(maximum - minimum);
  const ticks: number[] = [];
  for (let value = Math.ceil(minimum / unit) * unit; value <= maximum + unit * 0.001 && ticks.length < 100; value += unit) ticks.push(Number(value.toFixed(12)));
  if (!ticks.length) ticks.push(minimum, maximum);
  return { model, minimum, maximum, ticks };
}

function niceUnit(span: number): number {
  const rough = Math.max(span / 5, Number.MIN_VALUE);
  const power = 10 ** Math.floor(Math.log10(rough));
  const normalized = rough / power;
  return (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10) * power;
}

function scale(value: number, axis: ChartAxisLayout): number {
  const { model, minimum, maximum } = axis;
  const clamp = (ratio: number): number => Math.max(0, Math.min(1, model.reverseOrder ? 1 - ratio : ratio));
  if (model.scale === 'logarithmic') {
    const min = Math.log(Math.max(Number.MIN_VALUE, minimum)) / Math.log(model.logBase ?? 10);
    const max = Math.log(Math.max(Number.MIN_VALUE, maximum)) / Math.log(model.logBase ?? 10);
    return clamp((Math.log(Math.max(Number.MIN_VALUE, value)) / Math.log(model.logBase ?? 10) - min) / Math.max(Number.MIN_VALUE, max - min));
  }
  return clamp((value - minimum) / (maximum - minimum));
}

function valueAt(series: ResolvedChartSeries, index: number): number | null {
  const value = chartNumericValue(series.values[index]);
  return value === undefined ? null : value;
}

function xValueAt(series: ResolvedChartSeries, index: number): number | null {
  if (!series.xValues) return null;
  const value = chartNumericValue(series.xValues[index]);
  return value === undefined ? null : value;
}

function sizeValueAt(series: ResolvedChartSeries, index: number): number | null {
  if (!series.sizeValues) return null;
  const value = chartNumericValue(series.sizeValues[index]);
  return value === undefined ? null : value;
}

function errorAmount(model: NonNullable<ChartSeriesModel['errorBars']> | undefined, value: number, values: readonly number[], index: number, plusValues?: readonly PivotScalar[], minusValues?: readonly PivotScalar[]): { plus: number; minus: number } {
  if (!model) return { plus: 0, minus: 0 };
  if (model.type === 'custom') return { plus: Math.abs(chartNumericValue(plusValues?.[index]) ?? model.plusValue ?? 0), minus: Math.abs(chartNumericValue(minusValues?.[index]) ?? model.minusValue ?? 0) };
  if (model.type === 'fixed') return { plus: Math.abs(model.value ?? 0), minus: Math.abs(model.value ?? 0) };
  if (model.type === 'percentage') { const amount = Math.abs(value) * Math.abs(model.value ?? 0) / 100; return { plus: amount, minus: amount }; }
  const mean = values.length ? values.reduce((sum, current) => sum + current, 0) / values.length : 0;
  const deviation = values.length > 1 ? Math.sqrt(values.reduce((sum, current) => sum + (current - mean) ** 2, 0) / (values.length - 1)) : 0;
  const amount = model.type === 'standard-error' ? deviation / Math.sqrt(Math.max(1, values.length)) : deviation * Math.abs(model.value ?? 1);
  return { plus: amount, minus: amount };
}

function colorFor(series: ResolvedChartSeries, index: number): string { return series.color ?? DEFAULT_COLORS[index % DEFAULT_COLORS.length]!; }

function seriesModelFor(payload: ChartDrawingPayload, series: ResolvedChartSeries, index: number): ChartSeriesModel | undefined {
  return payload.series?.find((entry) => (entry.id && entry.id === series.id) || entry.name === series.name) ?? payload.series?.[index];
}

function linearRegression(points: Array<{ x: number; y: number }>): { slope: number; intercept: number } {
  if (!points.length) return { slope: 0, intercept: 0 };
  const meanX = points.reduce((sum, point) => sum + point.x, 0) / points.length;
  const meanY = points.reduce((sum, point) => sum + point.y, 0) / points.length;
  const denominator = points.reduce((sum, point) => sum + (point.x - meanX) ** 2, 0);
  const slope = denominator === 0 ? 0 : points.reduce((sum, point) => sum + (point.x - meanX) * (point.y - meanY), 0) / denominator;
  return { slope, intercept: meanY - slope * meanX };
}

function buildTrendline(model: ChartTrendlineModel, source: ChartLayoutPoint[], axis: ChartAxisLayout, plot: ChartLayout['plot']): ChartLayoutTrendline {
  const points = source.filter((point) => point.visible && point.value !== null).map((point) => ({ x: point.x, yValue: point.value! }));
  const xMin = points.length ? Math.min(...points.map((point) => point.x)) : plot.left;
  const xMax = points.length ? Math.max(...points.map((point) => point.x)) : plot.left + plot.width;
  const raw = points.map((point, index) => ({ x: index, y: point.yValue }));
  const regression = linearRegression(raw);
  const output: Array<{ x: number; y: number }> = [];
  for (let index = 0; index < Math.max(2, points.length); index += 1) {
    const normalized = index / Math.max(1, Math.max(2, points.length) - 1);
    const x = xMin + normalized * (xMax - xMin);
    let yValue: number;
    if (model.type === 'moving-average') {
      const period = Math.max(1, model.period ?? 2);
      const slice = raw.slice(Math.max(0, index - period + 1), index + 1);
      yValue = slice.reduce((sum, point) => sum + point.y, 0) / Math.max(1, slice.length);
    } else if (model.type === 'exponential') {
      yValue = Math.exp(Math.max(-20, Math.min(20, regression.intercept + regression.slope * index)));
    } else if (model.type === 'logarithmic') {
      yValue = regression.intercept + regression.slope * Math.log(Math.max(1, index + 1));
    } else if (model.type === 'power') {
      yValue = Math.exp(regression.intercept) * (index + 1) ** regression.slope;
    } else if (model.type === 'polynomial') {
      yValue = regression.intercept + regression.slope * index + (model.order && model.order > 1 ? index ** 2 * regression.slope * 0.02 : 0);
    } else {
      yValue = regression.intercept + regression.slope * index;
    }
    yValue += (model.intercept ?? 0) + (model.forwardForecast ?? 0) * normalized - (model.backwardForecast ?? 0) * (1 - normalized);
    output.push({ x, y: plot.top + plot.height * (1 - scale(yValue, axis)) });
  }
  return { model, points: output };
}

function quartile(sorted: readonly number[], ratio: number, inclusive: boolean): number {
  if (!sorted.length) return 0;
  const position = inclusive ? (sorted.length - 1) * ratio : (sorted.length + 1) * ratio - 1;
  const bounded = Math.max(0, Math.min(sorted.length - 1, position));
  const lower = Math.floor(bounded);
  const upper = Math.ceil(bounded);
  return sorted[lower]! + (sorted[upper]! - sorted[lower]!) * (bounded - lower);
}

function standardDeviation(values: readonly number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  return Math.sqrt(values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1));
}

function histogram(values: readonly number[], options: ChartHistogramOptions | undefined): ChartHistogramBinLayout[] {
  if (!values.length) return [];
  const minimum = Math.min(...values);
  const maximum = Math.max(...values);
  const span = Math.max(Number.EPSILON, maximum - minimum);
  const deviation = standardDeviation(values);
  const scottWidth = deviation > 0 ? 3.5 * deviation / values.length ** (1 / 3) : span / Math.max(1, Math.ceil(Math.sqrt(values.length)));
  const binCount = options?.mode === 'bin-count' ? Math.max(1, Math.floor(options.binCount ?? 1)) : options?.mode === 'bin-width' ? Math.max(1, Math.ceil(span / Math.max(Number.EPSILON, options.binWidth ?? scottWidth))) : Math.max(1, Math.ceil(span / Math.max(Number.EPSILON, scottWidth)));
  const width = options?.mode === 'bin-width' ? Math.max(Number.EPSILON, options.binWidth ?? scottWidth) : span / binCount;
  const counts = Array.from({ length: Math.max(1, Math.ceil(span / width)) }, () => 0);
  for (const value of values) {
    if (options?.underflow !== undefined && value < options.underflow) continue;
    if (options?.overflow !== undefined && value >= options.overflow) continue;
    const index = Math.min(counts.length - 1, Math.max(0, Math.floor((value - minimum) / width)));
    counts[index] = (counts[index] ?? 0) + 1;
  }
  return counts.map((count, index) => {
    const start = minimum + index * width;
    const end = start + width;
    return { start, end, count, label: `${trimNumber(start)}–${trimNumber(end)}` };
  });
}

function trimNumber(value: number): string { return Number(value.toFixed(6)).toString(); }

function statusError(kind: ChartDataStatus['kind'], code: ChartDataStatus['code'], message: string): ChartDataStatus { return { kind, code, message }; }

function baseLayout(payload: ChartDrawingPayload, data: ResolvedChartData, width: number, height: number, kind: ChartLayout['kind']): ChartLayout {
  const title = payload.elements.title ? { text: payload.elements.title, x: 16, y: 12 } : undefined;
  const legend = payload.elements.legend?.visible ? { visible: true, position: payload.elements.legend.position } : { visible: false, position: 'bottom' as const };
  const plot = { left: 52, top: title ? 40 : 22, width: Math.max(10, width - 70 - (legend.position === 'right' ? 100 : 0)), height: Math.max(10, height - (title ? 62 : 42) - (legend.position === 'bottom' ? 24 : 0)) };
  return { status: data.status, width, height, plot, title, legend, series: [], kind };
}

function chartKind(payload: ChartDrawingPayload): ChartLayout['kind'] {
  if (payload.chartType === 'pie' || payload.chartType === 'doughnut') return 'pie';
  if (payload.chartType === 'treemap') return 'treemap';
  if (payload.chartType === 'sunburst') return 'sunburst';
  if (payload.chartType === 'histogram' || payload.chartType === 'pareto') return 'histogram';
  if (payload.chartType === 'box-whisker') return 'box-whisker';
  if (payload.chartType === 'waterfall') return 'waterfall';
  if (payload.chartType === 'funnel') return 'funnel';
  if (payload.chartType === 'stock') return 'stock';
  if (payload.chartType === 'surface') return 'surface';
  if (payload.chartType === 'radar') return 'radar';
  if (payload.chartType === 'map') return 'map';
  return 'cartesian';
}

function createSeriesLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, plot: ChartLayout['plot'], categoryAxis: ChartAxisLayout, valueAxis: ChartAxisLayout, secondaryAxis: ChartAxisLayout | undefined): ChartLayoutSeries[] {
  const categoryCount = Math.max(1, data.categories.length, ...data.series.map((series) => series.values.length));
  const isScatter = payload.chartType === 'scatter' || payload.chartType === 'bubble';
  const xValues = data.series.flatMap((series) => series.xValues?.map(chartNumericValue).filter((value): value is number => value !== undefined) ?? []);
  const xAxis = isScatter ? axisBounds(payload.elements.categoryAxis ?? defaultAxis('x', 'bottom', 'value'), xValues, false) : categoryAxis;
  const result: ChartLayoutSeries[] = [];
  for (const [seriesIndex, series] of data.series.entries()) {
    const model = seriesModelFor(payload, series, seriesIndex);
    const chartType = series.chartType ?? (payload.chartType === 'combo' ? 'column' : payload.chartType);
    const axis = series.axis === 'secondary' ? secondaryAxis ?? valueAxis : valueAxis;
    const points: ChartLayoutPoint[] = [];
    const bars: ChartLayoutBar[] = [];
    const subtype = series.subtype ?? model?.subtype ?? payload.subtype;
    for (let index = 0; index < Math.max(categoryCount, series.values.length); index += 1) {
      const value = valueAt(series, index);
      const xValue = isScatter ? xValueAt(series, index) : null;
      const category = data.categories[index] ?? index + 1;
      const hasValue = value !== null;
      const visible = hasValue && (series.missing?.[index] !== true);
      const xRatio = isScatter ? (xValue === null ? 0 : scale(xValue, xAxis)) : (index + 0.5) / categoryCount;
      const yRatio = value === null ? 0 : scale(value, axis);
      const error = errorAmount(model?.errorBars ?? series.errorBars, value ?? 0, numberValues(series.values), index, series.errorPlusValues, series.errorMinusValues);
      const point = { index, category, value, ...(isScatter ? { xValue } : {}), ...(series.sizeValues ? { sizeValue: sizeValueAt(series, index) } : {}), x: plot.left + xRatio * plot.width, y: plot.top + (1 - yRatio) * plot.height, visible, ...(error.plus ? { errorPlus: error.plus } : {}), ...(error.minus ? { errorMinus: error.minus } : {}) };
      points.push(point);
      if (chartType === 'column' || chartType === 'bar') {
        const stack = payload.stacked ?? (subtype && ['stacked', 'percent-stacked', 'three-dimensional-stacked', 'three-dimensional-percent-stacked', 'stacked-markers', 'percent-stacked-markers'].includes(subtype) ? subtype.includes('percent') ? 'percent' : 'stacked' : 'none');
        const values = data.series.map((entry) => valueAt(entry, index) ?? 0);
        const start = stack === 'none' ? 0 : stack === 'percent' ? values.slice(0, seriesIndex).reduce((sum, current) => sum + current, 0) / Math.max(1, values.reduce((sum, current) => sum + Math.abs(current), 0)) * 100 : values.slice(0, seriesIndex).reduce((sum, current) => sum + current, 0);
        const end = stack === 'none' ? value ?? 0 : stack === 'percent' ? start + (value ?? 0) / Math.max(1, values.reduce((sum, current) => sum + Math.abs(current), 0)) * 100 : start + (value ?? 0);
        const slot = (chartType === 'bar' ? plot.height : plot.width) / categoryCount;
        const count = stack === 'none' ? Math.max(1, data.series.length) : 1;
        const band = slot * 0.72 / count;
        const offset = stack === 'none' ? seriesIndex * band : 0;
        const startRatio = scale(Math.min(start, end), axis);
        const endRatio = scale(Math.max(start, end), axis);
        bars.push({ index, category, start, end, x: chartType === 'bar' ? plot.left + startRatio * plot.width : plot.left + index * slot + slot * 0.14 + offset, y: chartType === 'bar' ? plot.top + index * slot + slot * 0.14 : plot.top + (1 - endRatio) * plot.height, width: chartType === 'bar' ? Math.max(1, (endRatio - startRatio) * plot.width) : Math.max(1, band - 1), height: chartType === 'bar' ? Math.max(3, slot * 0.72) : Math.max(1, (endRatio - startRatio) * plot.height), color: colorFor(series, seriesIndex), visible });
      }
    }
    const trendlines = (series.trendlines ?? []).map((trendline) => buildTrendline(trendline, points, axis, plot));
    result.push({ id: series.id, name: series.name, chartType, subtype, axis: series.axis, color: colorFor(series, seriesIndex), points, bars, trendlines, visible: model?.visible !== false, smooth: series.smooth ?? model?.smooth });
  }
  return result;
}

function pieSlices(payload: ChartDrawingPayload, data: ResolvedChartData, plot: ChartLayout['plot']): ChartPieSliceLayout[] {
  const slices: ChartPieSliceLayout[] = [];
  const ringCount = payload.chartType === 'doughnut' ? Math.max(1, data.series.length) : 1;
  const maxRadius = Math.min(plot.width, plot.height) * 0.43;
  const hole = payload.chartType === 'doughnut' ? 0.55 : 0;
  for (let seriesIndex = 0; seriesIndex < ringCount; seriesIndex += 1) {
    const series = data.series[seriesIndex] ?? data.series[0];
    if (!series) continue;
    const values = series.values.map(chartNumericValue).map((value) => Math.max(0, value ?? 0));
    const total = values.reduce((sum, value) => sum + value, 0);
    if (total <= 0) continue;
    const ringWidth = maxRadius * (1 - hole) / ringCount;
    let angle = -Math.PI / 2 + ((payload.subtype === 'exploded-pie' || payload.subtype === 'exploded-three-dimensional-pie' || payload.subtype === 'exploded-doughnut') ? Math.PI / 18 : 0);
    for (let pointIndex = 0; pointIndex < values.length; pointIndex += 1) {
      const sweep = values[pointIndex]! / total * Math.PI * 2;
      slices.push({ seriesIndex, pointIndex, value: values[pointIndex]!, startAngle: angle, endAngle: angle + sweep, innerRadius: payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * seriesIndex : 0, outerRadius: payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * (seriesIndex + 1) : maxRadius, explosion: payload.subtype?.includes('exploded') ? Math.min(12, maxRadius * 0.08) : 0, color: DEFAULT_COLORS[pointIndex % DEFAULT_COLORS.length]!, label: String(data.categories[pointIndex] ?? pointIndex + 1) });
      angle += sweep;
    }
  }
  return slices;
}

function boxLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, colors: string[]): ChartBoxLayout[] {
  const options: ChartBoxWhiskerOptions = payload.boxWhiskerOptions ?? { quartile: 'exclusive-median', showOutlierPoints: true };
  return data.series.map((series, seriesIndex) => {
    const values = numberValues(series.values).sort((left, right) => left - right);
    const q1 = quartile(values, 0.25, options.quartile === 'inclusive-median');
    const median = quartile(values, 0.5, options.quartile === 'inclusive-median');
    const q3 = quartile(values, 0.75, options.quartile === 'inclusive-median');
    const iqr = q3 - q1;
    const lowerFence = q1 - 1.5 * iqr;
    const upperFence = q3 + 1.5 * iqr;
    const inliers = values.filter((value) => value >= lowerFence && value <= upperFence);
    return { seriesIndex, minimum: values[0] ?? 0, lowerWhisker: inliers[0] ?? values[0] ?? 0, q1, median, q3, upperWhisker: inliers.at(-1) ?? values.at(-1) ?? 0, maximum: values.at(-1) ?? 0, mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0, outliers: options.showOutlierPoints === false ? [] : values.filter((value) => value < lowerFence || value > upperFence), color: colors[seriesIndex % colors.length]! };
  });
}

function waterfallLayouts(payload: ChartDrawingPayload, data: ResolvedChartData): ChartWaterfallBarLayout[] {
  const values = data.series[0]?.values.map(chartNumericValue).map((value) => value ?? 0) ?? [];
  const options: ChartWaterfallOptions = payload.waterfallOptions ?? { connectorLines: true };
  const totals = new Set(options.totalPointIndexes ?? []);
  let running = 0;
  return values.map((value, index) => {
    const total = totals.has(index);
    const start = total ? 0 : running;
    const end = total ? value : running + value;
    running = total ? value : end;
    return { index, start: Math.min(start, end), end: Math.max(start, end), total, color: total ? '#64748b' : value >= 0 ? '#10b981' : '#ef4444' };
  });
}

function stockLayouts(data: ResolvedChartData): ChartLayout['stockPoints'] {
  const series = data.series[0];
  if (!series?.stockValues) return undefined;
  const stock = series.stockValues;
  const count = Math.max(stock.high.length, stock.low.length, stock.close.length);
  return Array.from({ length: count }, (_, index) => {
    const high = chartNumericValue(stock.high[index]);
    const low = chartNumericValue(stock.low[index]);
    const close = chartNumericValue(stock.close[index]);
    if (high === undefined || low === undefined || close === undefined) return undefined;
    const open = stock.open ? chartNumericValue(stock.open[index]) : undefined;
    const volume = stock.volume ? chartNumericValue(stock.volume[index]) : undefined;
    return { index, ...(open === undefined ? {} : { open }), high, low, close, ...(volume === undefined ? {} : { volume }), color: open !== undefined ? close >= open ? '#10b981' : '#ef4444' : close >= (chartNumericValue(stock.close[index - 1]) ?? close) ? '#10b981' : '#ef4444' };
  }).filter((value): value is NonNullable<typeof value> => value !== undefined);
}

/** Build the only geometry contract consumed by Canvas and preview renderers. */
export function buildChartLayout(payload: ChartDrawingPayload, data: ResolvedChartData, width: number, height: number): ChartLayout {
  const kind = chartKind(payload);
  const layout = baseLayout(payload, data, width, height, kind);
  if (data.status.kind !== 'ready') return layout;
  const values = data.series.flatMap((series) => numberValues(series.values));
  const percent = payload.stacked === 'percent' || payload.subtype.includes('percent');
  const categoryAxis = axisBounds(payload.elements.categoryAxis ?? defaultAxis('category', 'bottom', 'category'), Array.from({ length: Math.max(1, data.categories.length) }, (_, index) => index), false);
  const primaryValues = data.series.filter((series) => series.axis !== 'secondary').flatMap((series) => numberValues(series.values));
  const secondaryValues = data.series.filter((series) => series.axis === 'secondary').flatMap((series) => numberValues(series.values));
  const valueAxis = axisBounds(payload.elements.valueAxis ?? defaultAxis('value', 'left', 'value'), primaryValues.length ? primaryValues : values, percent);
  const secondaryAxis = secondaryValues.length ? axisBounds(payload.elements.secondaryValueAxis ?? defaultAxis('secondary-value', 'right', 'value'), secondaryValues, false) : undefined;
  layout.categoryAxis = categoryAxis;
  layout.valueAxis = valueAxis;
  layout.secondaryValueAxis = secondaryAxis;
  layout.series = createSeriesLayouts(payload, data, layout.plot, categoryAxis, valueAxis, secondaryAxis);
  if (kind === 'pie') {
    layout.pieSlices = pieSlices(payload, data, layout.plot);
    return layout;
  }
  if (kind === 'histogram') {
    const bins = histogram(numberValues(data.series[0]?.values ?? []), payload.histogramOptions);
    layout.histogramBins = payload.chartType === 'pareto' ? bins.slice().sort((left, right) => right.count - left.count) : bins;
    if (payload.chartType === 'pareto') {
      let cumulative = 0;
      const total = bins.reduce((sum, bin) => sum + bin.count, 0) || 1;
      layout.paretoPoints = layout.histogramBins.map((bin, index) => { cumulative += bin.count; return { x: layout.plot.left + (index + 0.5) * layout.plot.width / Math.max(1, layout.histogramBins!.length), y: layout.plot.top + layout.plot.height * (1 - cumulative / total) }; });
    }
    return layout;
  }
  if (kind === 'box-whisker') {
    layout.boxes = boxLayouts(payload, data, DEFAULT_COLORS);
    return layout;
  }
  if (kind === 'waterfall') {
    layout.waterfallBars = waterfallLayouts(payload, data);
    return layout;
  }
  if (kind === 'funnel') {
    const valuesForFunnel = data.series[0]?.values.map(chartNumericValue).map((value) => Math.abs(value ?? 0)) ?? [];
    layout.funnelStages = valuesForFunnel.map((value, index) => ({ index, value, nextValue: valuesForFunnel[index + 1] ?? value, label: String(data.categories[index] ?? index + 1), color: DEFAULT_COLORS[index % DEFAULT_COLORS.length]! }));
    return layout;
  }
  if (kind === 'stock') {
    if (!data.series[0]?.stockValues) layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Stock charts require explicit High/Low/Close role bindings');
    else layout.stockPoints = stockLayouts(data);
    return layout;
  }
  if (kind === 'surface') {
    const all = data.series.flatMap((series) => numberValues(series.values));
    const min = Math.min(...all, 0);
    const max = Math.max(...all, 1);
    const span = Math.max(Number.EPSILON, max - min);
    layout.surfaceCells = data.series.flatMap((series, row) => series.values.map(chartNumericValue).map((value, column) => ({ row, column, value: value ?? 0, color: `rgb(${Math.round(37 + 202 * ((value ?? min) - min) / span)},${Math.round(99 + 100 * (1 - ((value ?? min) - min) / span))},${Math.round(235 - 167 * ((value ?? min) - min) / span)})` })));
    return layout;
  }
  if (kind === 'radar') {
    const count = Math.max(3, data.categories.length, ...data.series.map((series) => series.values.length));
    layout.radar = { count, maximum: Math.max(1, ...values.map(Math.abs)), points: data.series.map((series, seriesIndex) => ({ seriesIndex, values: series.values.map(chartNumericValue).map((value) => value ?? 0), color: colorFor(series, seriesIndex) })) };
    return layout;
  }
  if (kind === 'map') {
    layout.map = { ...(payload.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' }), resolved: false, reason: 'UNSUPPORTED_FEATURE: geographic entity resolution has no authoritative geometry source' };
    layout.status = statusError('unsupported', 'UNSUPPORTED_FEATURE', layout.map.reason);
    return layout;
  }
  if (payload.chartType === 'scatter' || payload.chartType === 'bubble') {
    const missingX = data.series.some((series) => series.xValues === undefined || series.values.some((_value, index) => chartNumericValue(series.values[index]) !== undefined && xValueAt(series, index) === null));
    const missingSize = payload.chartType === 'bubble' && data.series.some((series) => series.sizeValues === undefined || series.values.some((_value, index) => chartNumericValue(series.values[index]) !== undefined && sizeValueAt(series, index) === null));
    if (missingX) layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Scatter and Bubble charts require numeric X range bindings');
    else if (missingSize) layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Bubble charts require an independent numeric Size range binding');
  }
  return layout;
}
