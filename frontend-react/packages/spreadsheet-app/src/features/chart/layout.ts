import type {
  ChartAxisModel,
  ChartBoxWhiskerOptions,
  ChartDrawingPayload,
  ChartHistogramOptions,
  ChartMapResource,
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
  /** Exact canvas hit radius for scatter and bubble points. */
  markerRadius?: number;
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
  seriesIndex: number;
  index: number;
  start: number;
  end: number;
  total: boolean;
  color: string;
  visible: boolean;
}

export interface ChartMapFeatureLayout {
  id: string;
  label: string;
  categoryIndex: number;
  value: number | null;
  color: string;
  polygons: Array<Array<{ x: number; y: number }>>;
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
  specialSeriesIndex?: number;
  kind: 'cartesian' | 'pie' | 'treemap' | 'sunburst' | 'histogram' | 'box-whisker' | 'waterfall' | 'funnel' | 'stock' | 'surface' | 'radar' | 'map';
  pieSlices?: ChartPieSliceLayout[];
  histogramBins?: ChartHistogramBinLayout[];
  paretoPoints?: Array<{ x: number; y: number }>;
  boxes?: ChartBoxLayout[];
  waterfallBars?: ChartWaterfallBarLayout[];
  funnelStages?: Array<{ index: number; value: number; nextValue: number; label: string; color: string; visible: boolean }>;
  stockPoints?: Array<{ index: number; open?: number; high: number; low: number; close: number; volume?: number; color: string }>;
  surfaceCells?: Array<{ row: number; column: number; seriesIndex: number; value: number | null; color: string; visible: boolean }>;
  radar?: { count: number; maximum: number; points: Array<{ seriesIndex: number; values: number[]; visible: boolean[]; color: string }> };
  map?: ChartMapOptions & ({ resolved: false; reason: string } | { resolved: true; featureCount: number });
  mapFeatures?: ChartMapFeatureLayout[];
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
  const dataMinimum = finite.length ? finite.reduce((minimum, value) => Math.min(minimum, value), Infinity) : 0;
  const dataMaximum = finite.length ? finite.reduce((maximum, value) => Math.max(maximum, value), -Infinity) : 1;
  let minimum = model.minimum ?? (percent ? dataMinimum < 0 ? -100 : 0 : Math.min(0, dataMinimum));
  let maximum = model.maximum ?? (percent ? dataMaximum > 0 ? 100 : 0 : dataMaximum);
  if (model.minimum === undefined && !percent && minimum === maximum) minimum -= 1;
  if (model.maximum === undefined && !percent) {
    const span = Math.max(1, maximum - minimum);
    maximum += span * 0.1;
  }
  if (model.scale === 'logarithmic') {
    if (finite.some((value) => value <= 0) || minimum <= 0 || maximum <= 0) {
      throw new Error('INVALID_CHART_SOURCE: logarithmic axes require strictly positive finite values');
    }
    maximum = Math.max(minimum * 10, maximum);
  }
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum <= minimum) throw new Error('INVALID_CHART_SOURCE: Axis bounds are not finite');
  const ticks: number[] = [];
  if (model.scale === 'logarithmic') {
    const base = model.logBase ?? 10;
    const firstExponent = Math.ceil(Math.log(minimum) / Math.log(base));
    const lastExponent = Math.floor(Math.log(maximum) / Math.log(base));
    for (let exponent = firstExponent; exponent <= lastExponent && ticks.length < 100; exponent += 1) {
      const tick = base ** exponent;
      if (tick >= minimum && tick <= maximum) ticks.push(Number(tick.toPrecision(12)));
    }
  } else {
    const unit = model.majorUnit && model.majorUnit > 0 ? model.majorUnit : niceUnit(maximum - minimum);
    for (let value = Math.ceil(minimum / unit) * unit; value <= maximum + unit * 0.001 && ticks.length < 100; value += unit) ticks.push(Number(value.toFixed(12)));
  }
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

function errorAmount(model: NonNullable<ChartSeriesModel['errorBars']> | undefined, value: number, statistics: { count: number; deviation: number }, index: number, plusValues?: readonly PivotScalar[], minusValues?: readonly PivotScalar[]): { plus: number; minus: number } {
  if (!model) return { plus: 0, minus: 0 };
  if (model.type === 'custom') return { plus: Math.abs(chartNumericValue(plusValues?.[index]) ?? model.plusValue ?? 0), minus: Math.abs(chartNumericValue(minusValues?.[index]) ?? model.minusValue ?? 0) };
  if (model.type === 'fixed') return { plus: Math.abs(model.value ?? 0), minus: Math.abs(model.value ?? 0) };
  if (model.type === 'percentage') { const amount = Math.abs(value) * Math.abs(model.value ?? 0) / 100; return { plus: amount, minus: amount }; }
  const amount = model.type === 'standard-error' ? statistics.deviation / Math.sqrt(Math.max(1, statistics.count)) : statistics.deviation * Math.abs(model.value ?? 1);
  return { plus: amount, minus: amount };
}

function colorFor(series: ResolvedChartSeries, index: number): string { return series.color ?? DEFAULT_COLORS[index % DEFAULT_COLORS.length]!; }

function seriesModelFor(payload: ChartDrawingPayload, series: ResolvedChartSeries, index: number): ChartSeriesModel | undefined {
  if (!payload.series) return undefined;
  const byId = series.id ? payload.series.find((entry) => entry.id === series.id) : undefined;
  if (byId) return byId;
  const byName = payload.series.filter((entry) => entry.name === series.name);
  return byName.length === 1 ? byName[0] : payload.series[index];
}

function categorySlot(index: number, count: number, axis: ChartAxisLayout): number {
  return axis.model.reverseOrder ? count - index - 1 : index;
}

function effectiveChartType(payload: ChartDrawingPayload, series: ResolvedChartSeries): Exclude<ChartDrawingPayload['chartType'], 'combo'> {
  return series.chartType ?? (payload.chartType === 'combo' ? 'column' : payload.chartType);
}

function effectiveStack(payload: ChartDrawingPayload, series: ResolvedChartSeries, model: ChartSeriesModel | undefined): 'none' | 'stacked' | 'percent' {
  const subtype = series.subtype ?? model?.subtype ?? payload.subtype;
  return payload.stacked ?? (subtype && ['stacked', 'percent-stacked', 'three-dimensional-stacked', 'three-dimensional-percent-stacked', 'stacked-markers', 'percent-stacked-markers'].includes(subtype)
    ? subtype.includes('percent') ? 'percent' : 'stacked'
    : 'none');
}

interface BarPlacement {
  ordinal: number;
  count: number;
  starts: number[];
  ends: number[];
}

function buildBarPlacements(payload: ChartDrawingPayload, data: ResolvedChartData, categoryCount: number): Map<number, BarPlacement> {
  const groups = new Map<string, number[]>();
  for (const [seriesIndex, series] of data.series.entries()) {
    const model = seriesModelFor(payload, series, seriesIndex);
    const chartType = effectiveChartType(payload, series);
    if ((chartType !== 'column' && chartType !== 'bar') || model?.visible === false) continue;
    const stack = effectiveStack(payload, series, model);
    const key = `${series.axis}:${chartType}:${stack}`;
    const group = groups.get(key);
    if (group) group.push(seriesIndex);
    else groups.set(key, [seriesIndex]);
  }
  const placements = new Map<number, BarPlacement>();
  for (const [key, seriesIndexes] of groups) {
    const stack = key.endsWith(':percent') ? 'percent' : key.endsWith(':stacked') ? 'stacked' : 'none';
    const positiveTotals = Array.from({ length: categoryCount }, () => 0);
    const negativeTotals = Array.from({ length: categoryCount }, () => 0);
    if (stack === 'percent') {
      for (const seriesIndex of seriesIndexes) {
        for (let index = 0; index < categoryCount; index += 1) {
          const value = valueAt(data.series[seriesIndex]!, index) ?? 0;
          if (value >= 0) positiveTotals[index] = positiveTotals[index]! + value;
          else negativeTotals[index] = negativeTotals[index]! + Math.abs(value);
        }
      }
    }
    const positiveRunning = Array.from({ length: categoryCount }, () => 0);
    const negativeRunning = Array.from({ length: categoryCount }, () => 0);
    for (const [ordinal, seriesIndex] of seriesIndexes.entries()) {
      const starts: number[] = [];
      const ends: number[] = [];
      for (let index = 0; index < categoryCount; index += 1) {
        const raw = valueAt(data.series[seriesIndex]!, index) ?? 0;
        const value = stack === 'percent'
          ? raw / Math.max(Number.EPSILON, raw >= 0 ? positiveTotals[index]! : negativeTotals[index]!) * 100
          : raw;
        const running = raw >= 0 ? positiveRunning : negativeRunning;
        starts.push(stack === 'none' ? 0 : running[index]!);
        ends.push(stack === 'none' ? value : running[index]! + value);
        if (stack !== 'none') running[index] = running[index]! + value;
      }
      placements.set(seriesIndex, { ordinal, count: stack === 'none' ? seriesIndexes.length : 1, starts, ends });
    }
  }
  return placements;
}

function axisValuesForSeries(payload: ChartDrawingPayload, data: ResolvedChartData, axis: 'primary' | 'secondary', placements: ReadonlyMap<number, BarPlacement>): number[] {
  const values: number[] = [];
  for (const [seriesIndex, series] of data.series.entries()) {
    if (series.axis !== axis || seriesModelFor(payload, series, seriesIndex)?.visible === false) continue;
    const placement = placements.get(seriesIndex);
    if (placement) {
      values.push(...placement.starts, ...placement.ends);
    } else {
      values.push(...numberValues(series.values));
    }
  }
  return values;
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
  const points = source.filter((point) => point.visible && point.value !== null).map((point) => ({ canvasX: point.x, predictor: point.xValue ?? point.index, yValue: point.value! }));
  const xMin = points.length ? points.reduce((value, point) => Math.min(value, point.canvasX), Infinity) : plot.left;
  const xMax = points.length ? points.reduce((value, point) => Math.max(value, point.canvasX), -Infinity) : plot.left + plot.width;
  const predictorMin = points.length ? points.reduce((value, point) => Math.min(value, point.predictor), Infinity) : 0;
  const predictorMax = points.length ? points.reduce((value, point) => Math.max(value, point.predictor), -Infinity) : 1;
  const raw = points.map((point) => ({ x: point.predictor, y: point.yValue }));
  const regression = linearRegression(raw);
  const output: Array<{ x: number; y: number }> = [];
  for (let index = 0; index < Math.max(2, points.length); index += 1) {
    const normalized = index / Math.max(1, Math.max(2, points.length) - 1);
    const predictor = predictorMin + normalized * (predictorMax - predictorMin);
    const x = xMin + normalized * (xMax - xMin);
    let yValue: number;
    if (model.type === 'moving-average') {
      const period = Math.max(1, model.period ?? 2);
      const slice = raw.slice(Math.max(0, index - period + 1), index + 1);
      yValue = slice.reduce((sum, point) => sum + point.y, 0) / Math.max(1, slice.length);
    } else if (model.type === 'exponential') {
      yValue = Math.exp(Math.max(-20, Math.min(20, regression.intercept + regression.slope * predictor)));
    } else if (model.type === 'logarithmic') {
      yValue = regression.intercept + regression.slope * Math.log(Math.max(1, predictor));
    } else if (model.type === 'power') {
      yValue = Math.exp(regression.intercept) * Math.max(Number.MIN_VALUE, predictor) ** regression.slope;
    } else if (model.type === 'polynomial') {
      yValue = regression.intercept + regression.slope * predictor + (model.order && model.order > 1 ? predictor ** 2 * regression.slope * 0.02 : 0);
    } else {
      yValue = regression.intercept + regression.slope * predictor;
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
  const minimum = values.reduce((value, next) => Math.min(value, next), Infinity);
  const maximum = values.reduce((value, next) => Math.max(value, next), -Infinity);
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

function createSeriesLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, plot: ChartLayout['plot'], categoryAxis: ChartAxisLayout, valueAxis: ChartAxisLayout, secondaryAxis: ChartAxisLayout | undefined, barPlacements: ReadonlyMap<number, BarPlacement>): ChartLayoutSeries[] {
  const categoryCount = Math.max(1, data.categories.length, ...data.series.map((series) => series.values.length));
  const isScatter = payload.chartType === 'scatter' || payload.chartType === 'bubble';
  const xAxis = categoryAxis;
  const result: ChartLayoutSeries[] = [];
  for (const [seriesIndex, series] of data.series.entries()) {
    const model = seriesModelFor(payload, series, seriesIndex);
    const chartType = effectiveChartType(payload, series);
    const axis = series.axis === 'secondary' ? secondaryAxis ?? valueAxis : valueAxis;
    const points: ChartLayoutPoint[] = [];
    const bars: ChartLayoutBar[] = [];
    const subtype = series.subtype ?? model?.subtype ?? payload.subtype;
    const errorModel = model?.errorBars ?? series.errorBars;
    const errorValues = errorModel && ['standard-error', 'standard-deviation'].includes(errorModel.type) ? numberValues(series.values) : [];
    const errorStatistics = { count: errorValues.length, deviation: standardDeviation(errorValues) };
    const bubbleSizes = series.sizeValues?.map(chartNumericValue).filter((value): value is number => value !== undefined).map(Math.abs) ?? [];
    const bubbleMaximum = payload.chartType === 'bubble' ? Math.max(1, ...bubbleSizes) : 1;
    for (let index = 0; index < Math.max(categoryCount, series.values.length); index += 1) {
      const value = valueAt(series, index);
      const xValue = isScatter ? xValueAt(series, index) : null;
      const category = data.categories[index] ?? index + 1;
      const hasValue = value !== null;
      const visible = hasValue && (series.missing?.[index] !== true) && (!isScatter || xValue !== null)
        && (payload.chartType !== 'bubble' || sizeValueAt(series, index) !== null);
      const xRatio = isScatter ? (xValue === null ? 0 : scale(xValue, xAxis)) : (categorySlot(index, categoryCount, categoryAxis) + 0.5) / categoryCount;
      const yRatio = value === null ? 0 : scale(value, axis);
      const error = errorAmount(errorModel, value ?? 0, errorStatistics, index, series.errorPlusValues, series.errorMinusValues);
      const sizeValue = series.sizeValues ? sizeValueAt(series, index) : undefined;
      const markerRadius = isScatter
        ? payload.chartType === 'bubble'
          ? Math.max(3, Math.min(24, 3 + Math.sqrt(Math.abs(sizeValue ?? 0) / bubbleMaximum) * 18))
          : 4
        : undefined;
      const point = { index, category, value, ...(isScatter ? { xValue } : {}), ...(sizeValue === undefined ? {} : { sizeValue }), ...(markerRadius === undefined ? {} : { markerRadius }), x: plot.left + xRatio * plot.width, y: plot.top + (1 - yRatio) * plot.height, visible, ...(error.plus ? { errorPlus: error.plus } : {}), ...(error.minus ? { errorMinus: error.minus } : {}) };
      points.push(point);
      if (chartType === 'column' || chartType === 'bar') {
        const placement = barPlacements.get(seriesIndex);
        const start = placement?.starts[index] ?? 0;
        const end = placement?.ends[index] ?? (value ?? 0);
        const slot = (chartType === 'bar' ? plot.height : plot.width) / categoryCount;
        const count = placement?.count ?? 1;
        const band = slot * 0.72 / count;
        const offset = (placement?.ordinal ?? 0) * band;
        const lowerRatio = scale(Math.min(start, end), axis);
        const upperRatio = scale(Math.max(start, end), axis);
        const barStart = Math.min(lowerRatio, upperRatio);
        const barEnd = Math.max(lowerRatio, upperRatio);
        const slotIndex = categorySlot(index, categoryCount, categoryAxis);
        bars.push({ index, category, start, end, x: chartType === 'bar' ? plot.left + barStart * plot.width : plot.left + slotIndex * slot + slot * 0.14 + offset, y: chartType === 'bar' ? plot.top + slotIndex * slot + slot * 0.14 : plot.top + (1 - barEnd) * plot.height, width: chartType === 'bar' ? Math.max(1, (barEnd - barStart) * plot.width) : Math.max(1, band - 1), height: chartType === 'bar' ? Math.max(3, slot * 0.72) : Math.max(1, (barEnd - barStart) * plot.height), color: colorFor(series, seriesIndex), visible });
      }
    }
    const trendlines = (series.trendlines ?? []).map((trendline) => buildTrendline(trendline, points, axis, plot));
    result.push({ id: series.id, name: series.name, chartType, subtype, axis: series.axis, color: colorFor(series, seriesIndex), points, bars, trendlines, visible: model?.visible !== false, smooth: series.smooth ?? model?.smooth });
  }
  return result;
}

function pieSlices(payload: ChartDrawingPayload, data: ResolvedChartData, plot: ChartLayout['plot']): ChartPieSliceLayout[] {
  const slices: ChartPieSliceLayout[] = [];
  const visibleSeriesIndexes = data.series.flatMap((series, seriesIndex) => seriesModelFor(payload, series, seriesIndex)?.visible === false ? [] : [seriesIndex]);
  const ringCount = payload.chartType === 'doughnut' ? Math.max(1, visibleSeriesIndexes.length) : 1;
  const maxRadius = Math.min(plot.width, plot.height) * 0.43;
  const hole = payload.chartType === 'doughnut' ? 0.55 : 0;
  for (let ringIndex = 0; ringIndex < ringCount; ringIndex += 1) {
    const seriesIndex = visibleSeriesIndexes[ringIndex] ?? visibleSeriesIndexes[0];
    const series = seriesIndex === undefined ? undefined : data.series[seriesIndex];
    if (!series) continue;
    const values = series.values.map((raw) => {
      const value = chartNumericValue(raw);
      return value === undefined && payload.elements.emptyCells !== 'zero' ? null : Math.max(0, value ?? 0);
    });
    const total = values.reduce<number>((sum, value) => sum + (value ?? 0), 0);
    if (total <= 0) continue;
    const ringWidth = maxRadius * (1 - hole) / ringCount;
    let angle = -Math.PI / 2 + ((payload.subtype === 'exploded-pie' || payload.subtype === 'exploded-three-dimensional-pie' || payload.subtype === 'exploded-doughnut') ? Math.PI / 18 : 0);
    for (let pointIndex = 0; pointIndex < values.length; pointIndex += 1) {
      const value = values[pointIndex] ?? 0;
      if (value <= 0) continue;
      const sweep = value / total * Math.PI * 2;
      slices.push({ seriesIndex, pointIndex, value, startAngle: angle, endAngle: angle + sweep, innerRadius: payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * ringIndex : 0, outerRadius: payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * (ringIndex + 1) : maxRadius, explosion: payload.subtype?.includes('exploded') ? Math.min(12, maxRadius * 0.08) : 0, color: DEFAULT_COLORS[pointIndex % DEFAULT_COLORS.length]!, label: String(data.categories[pointIndex] ?? pointIndex + 1) });
      angle += sweep;
    }
  }
  return slices;
}

function boxLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, colors: string[]): ChartBoxLayout[] {
  const options: ChartBoxWhiskerOptions = payload.boxWhiskerOptions ?? { quartile: 'exclusive-median', showOutlierPoints: true };
  return data.series.flatMap((series, seriesIndex) => {
    if (seriesModelFor(payload, series, seriesIndex)?.visible === false) return [];
    const values = numberValues(series.values).sort((left, right) => left - right);
    const q1 = quartile(values, 0.25, options.quartile === 'inclusive-median');
    const median = quartile(values, 0.5, options.quartile === 'inclusive-median');
    const q3 = quartile(values, 0.75, options.quartile === 'inclusive-median');
    const iqr = q3 - q1;
    const lowerFence = q1 - 1.5 * iqr;
    const upperFence = q3 + 1.5 * iqr;
    const inliers = values.filter((value) => value >= lowerFence && value <= upperFence);
    return [{ seriesIndex, minimum: values[0] ?? 0, lowerWhisker: inliers[0] ?? values[0] ?? 0, q1, median, q3, upperWhisker: inliers.at(-1) ?? values.at(-1) ?? 0, maximum: values.at(-1) ?? 0, mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0, outliers: options.showOutlierPoints === false ? [] : values.filter((value) => value < lowerFence || value > upperFence), color: colors[seriesIndex % colors.length]! }];
  });
}

function waterfallLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, seriesIndex: number): ChartWaterfallBarLayout[] {
  const values = data.series[seriesIndex]?.values.map(chartNumericValue) ?? [];
  const options: ChartWaterfallOptions = payload.waterfallOptions ?? { connectorLines: true };
  const totals = new Set(options.totalPointIndexes ?? []);
  let running = 0;
  return values.map((value, index) => {
    if (value === undefined) return { seriesIndex, index, start: running, end: running, total: false, color: '#cbd5e1', visible: false };
    const total = totals.has(index);
    const start = total ? 0 : running;
    const end = total ? value : running + value;
    running = total ? value : end;
    return { seriesIndex, index, start: Math.min(start, end), end: Math.max(start, end), total, color: total ? '#64748b' : value >= 0 ? '#10b981' : '#ef4444', visible: true };
  });
}

function stockLayouts(data: ResolvedChartData, seriesIndex: number): ChartLayout['stockPoints'] {
  const series = data.series[seriesIndex];
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

function stockValidationError(data: ResolvedChartData, seriesIndex: number): string | undefined {
  const series = data.series[seriesIndex];
  const stock = series?.stockValues;
  if (!series || !stock) return 'Stock charts require explicit High/Low/Close role bindings';
  const count = data.categories.length;
  const vectors = [stock.high, stock.low, stock.close, ...(stock.open ? [stock.open] : []), ...(stock.volume ? [stock.volume] : [])];
  if (count < 1 || vectors.some((vector) => vector.length !== count)) return 'Stock role bindings must have the same point count as categories';
  for (let index = 0; index < count; index += 1) {
    const high = chartNumericValue(stock.high[index]);
    const low = chartNumericValue(stock.low[index]);
    const close = chartNumericValue(stock.close[index]);
    const open = stock.open ? chartNumericValue(stock.open[index]) : undefined;
    const volume = stock.volume ? chartNumericValue(stock.volume[index]) : undefined;
    if (high === undefined || low === undefined || close === undefined || high < low
      || close < low || close > high || (open !== undefined && (open < low || open > high))
      || (volume !== undefined && volume < 0)) {
      return `Stock roles are invalid at point ${String(index + 1)}`;
    }
  }
  return undefined;
}

function mapColor(scaleName: ChartMapOptions['colorScale'], value: number | null, minimum: number, maximum: number, index: number): string {
  if (value === null) return '#e2e8f0';
  if (scaleName === 'category') return DEFAULT_COLORS[index % DEFAULT_COLORS.length]!;
  const ratio = maximum === minimum ? 0.5 : Math.max(0, Math.min(1, (value - minimum) / (maximum - minimum)));
  const from = scaleName === 'diverging' ? (ratio < 0.5 ? [37, 99, 235] : [255, 255, 255]) : [219, 234, 254];
  const to = scaleName === 'diverging' ? (ratio < 0.5 ? [255, 255, 255] : [220, 38, 38]) : [29, 78, 216];
  const local = scaleName === 'diverging' ? (ratio < 0.5 ? ratio * 2 : (ratio - 0.5) * 2) : ratio;
  const channels = from.map((channel, channelIndex) => Math.round(channel + (to[channelIndex]! - channel) * local));
  return `rgb(${channels[0]},${channels[1]},${channels[2]})`;
}

function mapLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, plot: ChartLayout['plot'], seriesIndex: number): { resource: ChartMapResource; features: ChartMapFeatureLayout[] } | ChartDataStatus {
  const options: ChartMapOptions = payload.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' };
  const resource = options.resource;
  if (!resource || resource.features.length === 0) return statusError('unsupported', 'UNSUPPORTED_FEATURE', 'UNSUPPORTED_FEATURE: map requires a validated offline GeoJSON resource');
  let minimumLongitude = Number.POSITIVE_INFINITY;
  let maximumLongitude = Number.NEGATIVE_INFINITY;
  let minimumLatitude = Number.POSITIVE_INFINITY;
  let maximumLatitude = Number.NEGATIVE_INFINITY;
  for (const feature of resource.features) {
    for (const polygon of feature.polygons) {
      for (const [longitude, latitude] of polygon) {
        minimumLongitude = Math.min(minimumLongitude, longitude);
        maximumLongitude = Math.max(maximumLongitude, longitude);
        minimumLatitude = Math.min(minimumLatitude, latitude);
        maximumLatitude = Math.max(maximumLatitude, latitude);
      }
    }
  }
  if (!Number.isFinite(minimumLongitude) || !Number.isFinite(maximumLongitude) || !Number.isFinite(minimumLatitude) || !Number.isFinite(maximumLatitude)) {
    return statusError('invalid', 'INVALID_CHART_SOURCE', 'INVALID_CHART_SOURCE: map resource has no drawable coordinates');
  }
  const longitudeSpan = Math.max(Number.EPSILON, maximumLongitude - minimumLongitude);
  const latitudeSpan = Math.max(Number.EPSILON, maximumLatitude - minimumLatitude);
  const firstSeries = data.series[seriesIndex];
  const categoryIndexByKey = new Map<string, number>();
  data.categories.forEach((category, index) => {
    const key = String(category ?? '').trim().toLowerCase();
    if (!categoryIndexByKey.has(key)) categoryIndexByKey.set(key, index);
  });
  const categoryIndexes = resource.features.map((feature) => categoryIndexByKey.get(feature.id.trim().toLowerCase()) ?? categoryIndexByKey.get(feature.label.trim().toLowerCase()) ?? -1);
  const values = resource.features.map((_feature, featureIndex) => {
    const categoryIndex = categoryIndexes[featureIndex] ?? -1;
    return categoryIndex < 0 ? null : chartNumericValue(firstSeries?.values[categoryIndex]);
  });
  const finite = values.filter((value): value is number => value !== null && Number.isFinite(value));
  const minimum = finite.length ? Math.min(...finite) : 0;
  const maximum = finite.length ? Math.max(...finite) : 1;
  const features = resource.features.map((feature, index) => ({
    id: feature.id,
    label: feature.label,
    categoryIndex: categoryIndexes[index] ?? -1,
    value: values[index] ?? null,
    color: mapColor(options.colorScale, values[index] ?? null, minimum, maximum, index),
    polygons: feature.polygons.map((polygon) => polygon.map(([longitude, latitude]) => ({
      x: plot.left + ((longitude - minimumLongitude) / longitudeSpan) * plot.width,
      y: plot.top + ((maximumLatitude - latitude) / latitudeSpan) * plot.height,
    }))),
  }));
  return { resource, features };
}

/** Build the only geometry contract consumed by Canvas and preview renderers. */
export function buildChartLayout(payload: ChartDrawingPayload, data: ResolvedChartData, width: number, height: number): ChartLayout {
  const kind = chartKind(payload);
  const layout = baseLayout(payload, data, width, height, kind);
  if (data.status.kind !== 'ready') return layout;
  const values = data.series.flatMap((series) => numberValues(series.values));
  const percent = payload.stacked === 'percent' || payload.subtype.includes('percent');
  const isScatter = payload.chartType === 'scatter' || payload.chartType === 'bubble';
  const xValues = data.series.flatMap((series) => series.xValues?.map(chartNumericValue).filter((value): value is number => value !== undefined) ?? []);
  const categoryCount = Math.max(1, data.categories.length, ...data.series.map((series) => series.values.length));
  const barPlacements = buildBarPlacements(payload, data, categoryCount);
  const primaryValues = axisValuesForSeries(payload, data, 'primary', barPlacements);
  const secondaryValues = axisValuesForSeries(payload, data, 'secondary', barPlacements);
  let categoryAxis: ChartAxisLayout;
  let valueAxis: ChartAxisLayout;
  let secondaryAxis: ChartAxisLayout | undefined;
  try {
    categoryAxis = axisBounds(
      payload.elements.categoryAxis ?? defaultAxis(isScatter ? 'x' : 'category', 'bottom', isScatter ? 'value' : 'category'),
      isScatter ? xValues : Array.from({ length: Math.max(1, data.categories.length) }, (_, index) => index),
      false,
    );
    valueAxis = axisBounds(payload.elements.valueAxis ?? defaultAxis('value', 'left', 'value'), primaryValues.length ? primaryValues : values, percent);
    secondaryAxis = secondaryValues.length ? axisBounds(payload.elements.secondaryValueAxis ?? defaultAxis('secondary-value', 'right', 'value'), secondaryValues, false) : undefined;
  } catch (error) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', error instanceof Error ? error.message : String(error));
    return layout;
  }
  layout.categoryAxis = categoryAxis;
  layout.valueAxis = valueAxis;
  layout.secondaryValueAxis = secondaryAxis;
  layout.series = createSeriesLayouts(payload, data, layout.plot, categoryAxis, valueAxis, secondaryAxis, barPlacements);
  const specialSeriesIndex = layout.series.findIndex((series) => series.visible);
  layout.specialSeriesIndex = specialSeriesIndex;
  if (specialSeriesIndex < 0) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'INVALID_CHART_SOURCE: chart has no visible series');
    return layout;
  }
  if (kind === 'map') {
    const map = mapLayouts(payload, data, layout.plot, specialSeriesIndex);
    if ('kind' in map) {
      layout.status = map;
      layout.map = { ...(payload.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' }), resolved: false, reason: map.message ?? 'Map resource is unavailable' };
    } else {
      layout.mapFeatures = map.features;
      layout.map = { ...(payload.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' }), resolved: true, featureCount: map.features.length };
    }
    return layout;
  }
  if (kind === 'treemap' || kind === 'sunburst') {
    if (data.series.some((series) => series.values.some((value) => {
      const numeric = chartNumericValue(value);
      return numeric !== undefined && numeric < 0;
    }))) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Treemap and sunburst charts require non-negative values');
      return layout;
    }
    if (!layout.series.some((series) => series.visible && series.points.some((point) => point.visible && (point.value ?? 0) > 0))) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Treemap and sunburst charts require at least one positive value');
      return layout;
    }
  }
  if (kind === 'pie') {
    if (payload.chartType === 'pie' && layout.series.filter((series) => series.visible).length !== 1) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pie charts require exactly one visible series');
      return layout;
    }
    if (data.series.some((series) => series.values.some((value) => {
      const numeric = chartNumericValue(value);
      return numeric !== undefined && numeric < 0;
    }))) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pie and doughnut charts require non-negative values');
      return layout;
    }
    layout.pieSlices = pieSlices(payload, data, layout.plot);
    if (layout.pieSlices.length === 0) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pie and doughnut charts require at least one positive value');
    }
    return layout;
  }
  if (kind === 'histogram') {
    const rawValues = data.series[specialSeriesIndex]?.values.map((value, index) => ({ index, value: chartNumericValue(value) })).filter((entry): entry is { index: number; value: number } => entry.value !== undefined) ?? [];
    const bins = histogram(rawValues.map((entry) => entry.value), payload.histogramOptions);
    if (rawValues.length === 0 || bins.every((bin) => bin.count === 0)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram charts require at least one value in range');
      return layout;
    }
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
    if (layout.boxes.length === 0 || !data.series.some((series, seriesIndex) => layout.series[seriesIndex]?.visible && numberValues(series.values).length > 0)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Box and whisker charts require at least one numeric value');
    }
    return layout;
  }
  if (kind === 'waterfall') {
    layout.waterfallBars = waterfallLayouts(payload, data, specialSeriesIndex);
    if (!layout.waterfallBars.some((bar) => bar.visible)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Waterfall charts require at least one numeric value');
    }
    return layout;
  }
  if (kind === 'funnel') {
    const rawValues = data.series[specialSeriesIndex]?.values.map(chartNumericValue) ?? [];
    if (rawValues.some((value) => value !== undefined && value < 0)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Funnel charts require non-negative values');
      return layout;
    }
    const valuesForFunnel = rawValues.map((value) => value ?? null);
    layout.funnelStages = valuesForFunnel.map((value, index) => ({ index, value: value ?? 0, nextValue: valuesForFunnel[index + 1] ?? value ?? 0, label: String(data.categories[index] ?? index + 1), color: DEFAULT_COLORS[index % DEFAULT_COLORS.length]!, visible: value !== null }));
    if (!layout.funnelStages.some((stage) => stage.visible)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Funnel charts require at least one numeric value');
    }
    return layout;
  }
  if (kind === 'stock') {
    const error = stockValidationError(data, specialSeriesIndex);
    if (error) layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', error);
    else layout.stockPoints = stockLayouts(data, specialSeriesIndex);
    return layout;
  }
  if (kind === 'surface') {
    const all = data.series.flatMap((series, seriesIndex) => layout.series[seriesIndex]?.visible === false ? [] : numberValues(series.values));
    const min = all.reduce((value, next) => Math.min(value, next), 0);
    const max = all.reduce((value, next) => Math.max(value, next), 1);
    const span = Math.max(Number.EPSILON, max - min);
    let visibleRow = 0;
    layout.surfaceCells = data.series.flatMap((series, seriesIndex) => {
      if (layout.series[seriesIndex]?.visible === false) return [];
      const row = visibleRow++;
      return series.values.map(chartNumericValue).map((value, column) => {
        if (value === undefined) return { row, column, seriesIndex, value: null, color: '#e2e8f0', visible: false };
        const ratio = (value - min) / span;
        return { row, column, seriesIndex, value, color: `rgb(${Math.round(37 + 202 * ratio)},${Math.round(99 + 100 * (1 - ratio))},${Math.round(235 - 167 * ratio)})`, visible: true };
      });
    });
    if (!layout.surfaceCells.some((cell) => cell.visible)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Surface charts require at least one numeric value');
    }
    return layout;
  }
  if (kind === 'radar') {
    const visibleSeries = data.series.flatMap((series, seriesIndex) => layout.series[seriesIndex]?.visible === false ? [] : [{ series, seriesIndex }]);
    const count = Math.max(3, data.categories.length, ...visibleSeries.map(({ series }) => series.values.length));
    const radarValues = visibleSeries.flatMap(({ series }) => numberValues(series.values));
    layout.radar = { count, maximum: radarValues.reduce((maximum, value) => Math.max(maximum, Math.abs(value)), 1), points: visibleSeries.map(({ series, seriesIndex }) => ({ seriesIndex, values: series.values.map(chartNumericValue).map((value) => value ?? 0), visible: series.values.map((value) => chartNumericValue(value) !== undefined), color: colorFor(series, seriesIndex) })) };
    if (radarValues.length === 0) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Radar charts require at least one numeric value');
    }
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
