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
import { isChartHistogramOptions } from '@react-sheets/core-model';
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
  dataLabelText?: string;
  dataLabelX?: number;
  dataLabelY?: number;
  dataLabelBounds?: { left: number; top: number; right: number; bottom: number };
}

export type ChartHistogramBinLayout = {
  count: number;
  label: string;
  geometry: { x: number; y: number; width: number; height: number };
} & (
  | { kind: 'numeric'; start: number; end: number; boundary?: 'underflow' | 'overflow'; category?: never; value?: never }
  | { kind: 'category'; category: PivotScalar; value: number; start?: never; end?: never }
);

type ChartHistogramBinValue =
  | { kind: 'numeric'; start: number; end: number; count: number; label: string; boundary?: 'underflow' | 'overflow' }
  | { kind: 'category'; category: PivotScalar; value: number; count: number; label: string };

type HistogramBuildResult = { ok: true; bins: ChartHistogramBinValue[] } | { ok: false; status: ChartDataStatus };

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
  innerPoints: number[];
  showMeanMarker: boolean;
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
  geometry: { x: number; y: number; width: number; height: number };
  connector?: { startX: number; endX: number; y: number };
}

export interface ChartMapFeatureLayout {
  id: string;
  label: string;
  categoryIndex: number;
  value: number | null;
  color: string;
  polygons: Array<Array<{ x: number; y: number }>>;
}

export interface ChartDataTableLayout {
  bounds: { left: number; top: number; width: number; height: number };
  rowHeight: number;
  categoryCount: number;
  legendColumnWidth: number;
  showLegendKeys: boolean;
  categories: readonly PivotScalar[];
  series: readonly { name: string; color: string; values: readonly PivotScalar[] }[];
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
  stockVolume?: { maximum: number; top: number; height: number; priceHeight: number };
  surfaceCells?: Array<{ row: number; column: number; seriesIndex: number; value: number | null; color: string; visible: boolean }>;
  radar?: {
    count: number;
    centerX: number;
    centerY: number;
    radius: number;
    points: Array<{
      seriesIndex: number;
      color: string;
      vertices: Array<{ index: number; x: number; y: number; visible: boolean }>;
    }>;
  };
  map?: ChartMapOptions & ({ resolved: false; reason: string } | { resolved: true; featureCount: number });
  mapFeatures?: ChartMapFeatureLayout[];
  dataTable?: ChartDataTableLayout;
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
  const logarithmic = model.scale === 'logarithmic';
  const positiveMinimum = finite.length ? finite.reduce((minimum, value) => value > 0 ? Math.min(minimum, value) : minimum, Infinity) : 1;
  let minimum = model.minimum ?? (logarithmic ? positiveMinimum : percent ? dataMinimum < 0 ? -100 : 0 : Math.min(0, dataMinimum));
  let maximum = model.maximum ?? (percent ? dataMaximum > 0 ? 100 : 0 : dataMaximum);
  if (model.minimum === undefined && !percent && !logarithmic && minimum === maximum) minimum -= 1;
  if (model.maximum === undefined && !percent && !logarithmic) {
    const span = Math.max(1, maximum - minimum);
    maximum += span * 0.1;
  }
  if (logarithmic) {
    const base = model.logBase ?? 10;
    if (finite.some((value) => value <= 0) || minimum <= 0 || maximum <= 0) {
      throw new Error('INVALID_CHART_SOURCE: logarithmic axes require strictly positive finite values');
    }
    if (!Number.isFinite(base) || base <= 1) throw new Error('INVALID_CHART_SOURCE: logarithmic axis base must be greater than one');
    if (model.maximum === undefined) maximum = Math.max(minimum * base, maximum);
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
      for (const value of placement.starts) values.push(value);
      for (const value of placement.ends) values.push(value);
    } else {
      for (const value of series.values) {
        const numeric = chartNumericValue(value);
        if (numeric !== undefined) values.push(numeric);
      }
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

function histogramFailure(kind: 'invalid' | 'unsupported', code: NonNullable<ChartDataStatus['code']>, message: string): HistogramBuildResult {
  return { ok: false, status: { kind, code, message } };
}

function histogram(values: readonly PivotScalar[], options: ChartHistogramOptions | undefined, maximumBinCount: number): HistogramBuildResult {
  if (options !== undefined && !isChartHistogramOptions(options)) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin options are invalid');
  }
  if (options?.mode === 'by-category') {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'By-category histogram options require categorical source grouping');
  }
  if (!Number.isSafeInteger(maximumBinCount) || maximumBinCount < 1) {
    return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'Histogram plot width cannot display a bin without overlap');
  }

  const calculateVariance = options === undefined || options.mode === 'automatic';
  let numericCount = 0;
  let minimum = Infinity;
  let maximum = -Infinity;
  let mean = 0;
  let squaredDeviation = 0;
  let regularValueCount = 0;
  for (const rawValue of values) {
    const numeric = chartNumericValue(rawValue);
    if (numeric === undefined) continue;
    numericCount += 1;
    minimum = Math.min(minimum, numeric);
    maximum = Math.max(maximum, numeric);
    const regular = (options?.underflow === undefined || numeric > options.underflow)
      && (options?.overflow === undefined || numeric <= options.overflow);
    if (regular) {
      regularValueCount += 1;
      if (calculateVariance) {
        const delta = numeric - mean;
        mean += delta / regularValueCount;
        squaredDeviation += delta * (numeric - mean);
      }
    }
  }
  if (numericCount === 0) return { ok: true, bins: [] };

  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || (calculateVariance && !Number.isFinite(squaredDeviation))) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram numeric range exceeds finite calculation bounds');
  }
  const underflow = options?.underflow;
  const overflow = options?.overflow;
  const specialBinCount = Number(underflow !== undefined) + Number(overflow !== undefined);
  const regularMinimum = underflow ?? minimum;
  const regularMaximum = overflow === undefined ? maximum : Math.min(maximum, overflow);
  const rawRegularSpan = regularMaximum - regularMinimum;
  if (!Number.isFinite(rawRegularSpan)) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram thresholds exceed finite calculation bounds');
  }
  if (rawRegularSpan < 0 && regularValueCount > 0) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram thresholds do not enclose the values assigned to regular bins');
  }
  const regularSpan = Math.max(0, rawRegularSpan);
  const sampleDeviation = calculateVariance && regularValueCount > 1
    ? Math.sqrt(Math.max(0, squaredDeviation / (regularValueCount - 1)))
    : 0;
  const fallbackWidth = Math.max(1, Math.abs(regularMinimum) * Number.EPSILON * 4);
  const automaticSampleCount = regularValueCount || numericCount;
  const automaticWidthCandidate = sampleDeviation > 0
    ? 3.5 * sampleDeviation / regularValueCount ** (1 / 3)
    : regularSpan > 0 ? regularSpan / Math.max(1, Math.ceil(Math.sqrt(automaticSampleCount))) : fallbackWidth;
  const automaticWidth = automaticWidthCandidate > 0 ? automaticWidthCandidate : regularSpan > 0 ? regularSpan : automaticWidthCandidate;
  if (!Number.isFinite(automaticWidth) || automaticWidth <= 0) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram automatic bin width is not finite and positive');
  }

  let regularBinCount: number;
  let regularWidth = automaticWidth;
  if (options?.mode === 'bin-count') {
    if (options.binCount! > maximumBinCount) {
      return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'Requested histogram bin count exceeds the drawable plot resolution');
    }
    regularBinCount = options.binCount! - specialBinCount;
    if (regularBinCount === 0 && regularValueCount > 0) {
      return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin count leaves values between the underflow and overflow bins unassigned');
    }
    if (regularSpan > 0 && regularBinCount > 0) regularWidth = regularSpan / regularBinCount;
  } else if (options?.mode === 'bin-width' || options?.mode === 'automatic' || options === undefined) {
    if (options?.mode === 'bin-width') regularWidth = options.binWidth!;
    if (regularSpan === 0) regularBinCount = regularValueCount > 0 ? 1 : 0;
    else {
      const widthCount = regularSpan / regularWidth;
      const availableRegularBins = maximumBinCount - specialBinCount;
      if (!Number.isFinite(widthCount) || widthCount > availableRegularBins) {
        return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'Histogram bin width exceeds the drawable plot resolution');
      }
      regularBinCount = Math.max(1, Math.ceil(widthCount));
    }
  } else {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram mode is invalid');
  }

  const totalBinCount = specialBinCount + regularBinCount;
  if (!Number.isSafeInteger(totalBinCount) || totalBinCount > maximumBinCount) {
    return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'Histogram bins exceed the drawable plot resolution');
  }
  if (regularSpan === 0 && regularBinCount > 1) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin count cannot partition a zero-width numeric range');
  }
  if (regularBinCount > 0 && (!Number.isFinite(regularWidth) || regularWidth <= 0)) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin boundaries are not representable at numeric precision');
  }
  if (regularBinCount > 0 && regularSpan === 0 && regularValueCount === 0) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram thresholds leave no range for the requested regular bins');
  }

  const regularCounts = Array<number>(regularBinCount).fill(0);
  let underflowCount = 0;
  let overflowCount = 0;
  for (const rawValue of values) {
    const value = chartNumericValue(rawValue);
    if (value === undefined) continue;
    if (underflow !== undefined && value <= underflow) {
      underflowCount += 1;
      continue;
    }
    if (overflow !== undefined && value > overflow) {
      overflowCount += 1;
      continue;
    }
    if (regularBinCount === 0) {
      return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram regular values have no bin');
    }
    const index = regularSpan === 0
      ? 0
      : Math.max(0, Math.min(regularBinCount - 1, Math.ceil((value - regularMinimum) / regularWidth) - 1));
    regularCounts[index] = regularCounts[index]! + 1;
  }

  const bins: ChartHistogramBinValue[] = [];
  if (underflow !== undefined) bins.push({ kind: 'numeric', start: underflow, end: underflow, count: underflowCount, label: `≤ ${trimNumber(underflow)}`, boundary: 'underflow' });
  let previousEnd = regularMinimum;
  for (let index = 0; index < regularBinCount; index += 1) {
    const end = regularSpan === 0
      ? regularMaximum
      : index === regularBinCount - 1
        ? regularMaximum
        : options?.mode === 'bin-width'
          ? regularMinimum + regularWidth * (index + 1)
          : regularMinimum + regularSpan * (index + 1) / regularBinCount;
    if (!Number.isFinite(end) || end < previousEnd || (regularSpan > 0 && end === previousEnd)) {
      return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin boundaries are not finite and increasing');
    }
    const label = index === 0 && underflow === undefined
      ? `≤ ${trimNumber(end)}`
      : `(${trimNumber(previousEnd)}, ${trimNumber(end)}]`;
    bins.push({ kind: 'numeric', start: previousEnd, end, count: regularCounts[index]!, label });
    previousEnd = end;
  }
  if (overflow !== undefined) bins.push({ kind: 'numeric', start: overflow, end: overflow, count: overflowCount, label: `> ${trimNumber(overflow)}`, boundary: 'overflow' });
  const assignedCount = bins.reduce((sum, bin) => sum + bin.count, 0);
  if (assignedCount !== numericCount) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin assignment did not preserve every numeric source value');
  }
  return { ok: true, bins };
}

function categoricalHistogram(categories: readonly PivotScalar[], values: readonly PivotScalar[], maximumBinCount: number): HistogramBuildResult {
  if (categories.length !== values.length) {
    return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'By-category histogram requires one category for every value');
  }
  if (!Number.isSafeInteger(maximumBinCount) || maximumBinCount < 1) {
    return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'Histogram plot width cannot display a category without overlap');
  }
  const grouped = new Map<string, { category: PivotScalar; count: number; value: number; label: string }>();
  for (let index = 0; index < categories.length; index += 1) {
    const category = categories[index];
    const value = chartNumericValue(values[index]);
    if (category === undefined || value === undefined) continue;
    const key = category === null
      ? 'blank:'
      : typeof category === 'object'
        ? `error:${category.code}`
        : `${typeof category}:${String(category)}`;
    const label = category === null ? '(blank)' : typeof category === 'object' ? `#${category.code}` : String(category);
    const current = grouped.get(key);
    if (current) {
      const aggregate = current.value + value;
      if (!Number.isFinite(aggregate)) {
        return histogramFailure('invalid', 'INVALID_CHART_SOURCE', 'By-category histogram aggregate exceeds finite calculation bounds');
      }
      current.count += 1;
      current.value = aggregate;
    } else {
      if (grouped.size >= maximumBinCount) {
        return histogramFailure('unsupported', 'UNSUPPORTED_FEATURE', 'By-category histogram exceeds the drawable plot resolution');
      }
      grouped.set(key, { category, count: 1, value, label });
    }
  }
  return { ok: true, bins: [...grouped.values()].map((bin) => ({ kind: 'category', ...bin })) };
}

function createHistogramSeriesLayouts(payload: ChartDrawingPayload, data: ResolvedChartData): ChartLayoutSeries[] {
  return data.series.map((series, index) => {
    const model = seriesModelFor(payload, series, index);
    return {
      id: series.id,
      name: series.name,
      chartType: effectiveChartType(payload, series),
      subtype: series.subtype ?? model?.subtype ?? payload.subtype,
      axis: series.axis,
      color: colorFor(series, index),
      points: [],
      bars: [],
      trendlines: [],
      visible: model?.visible !== false,
      smooth: series.smooth ?? model?.smooth,
    };
  });
}

function buildHistogramLayout(payload: ChartDrawingPayload, data: ResolvedChartData, layout: ChartLayout): ChartLayout {
  const hasHistogramOptions = Object.prototype.hasOwnProperty.call(payload, 'histogramOptions');
  if (hasHistogramOptions && !isChartHistogramOptions(payload.histogramOptions)) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin options are invalid');
    return layout;
  }
  layout.series = createHistogramSeriesLayouts(payload, data);
  const visibleSeries = layout.series.filter((series) => series.visible);
  if (visibleSeries.length !== 1) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram and Pareto charts require exactly one visible series');
    return layout;
  }
  const specialSeriesIndex = layout.series.indexOf(visibleSeries[0]!);
  layout.specialSeriesIndex = specialSeriesIndex;
  const values = data.series[specialSeriesIndex]?.values ?? [];
  const maximumBinCount = Math.floor(layout.plot.width);
  const result = payload.histogramOptions?.mode === 'by-category'
    ? categoricalHistogram(data.categories, values, maximumBinCount)
    : histogram(values, payload.histogramOptions, maximumBinCount);
  if (!result.ok) {
    layout.status = result.status;
    return layout;
  }
  const bins = result.bins;
  if (bins.length === 0 || bins.every((bin) => bin.count === 0)) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram charts require at least one value in range');
    return layout;
  }
  let minimumValue = 0;
  let maximumValue = 1;
  let total = 0;
  for (const bin of bins) {
    const value = histogramValue(bin);
    if (!Number.isFinite(value)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram bin value exceeds finite calculation bounds');
      return layout;
    }
    if (payload.chartType === 'pareto') {
      if (value < 0) {
        layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pareto charts require non-negative category totals');
        return layout;
      }
      total += value;
      if (!Number.isFinite(total)) {
        layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pareto total exceeds finite calculation bounds');
        return layout;
      }
    }
    minimumValue = Math.min(minimumValue, value);
    maximumValue = Math.max(maximumValue, value);
  }
  if (!Number.isFinite(maximumValue - minimumValue)) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Histogram value range exceeds finite geometry bounds');
    return layout;
  }
  const orderedBins = payload.chartType === 'pareto' ? bins.slice().sort((left, right) => histogramValue(right) - histogramValue(left)) : bins;
  layout.histogramBins = histogramGeometry(orderedBins, layout.plot);
  if (payload.chartType === 'pareto') {
    let cumulative = 0;
    if (total <= 0) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Pareto charts require a positive total');
      return layout;
    }
    layout.paretoPoints = layout.histogramBins.map((bin) => { cumulative += histogramValue(bin); return { x: bin.geometry.x + bin.geometry.width / 2, y: layout.plot.top + layout.plot.height * (1 - cumulative / total) }; });
  }
  return layout;
}

function histogramValue(bin: ChartHistogramBinValue): number {
  return bin.kind === 'category' ? bin.value : bin.count;
}

function histogramGeometry(bins: readonly ChartHistogramBinValue[], plot: ChartLayout['plot']): ChartHistogramBinLayout[] {
  const minimum = bins.reduce((value, bin) => Math.min(value, histogramValue(bin)), 0);
  const maximum = bins.reduce((value, bin) => Math.max(value, histogramValue(bin)), 1);
  const span = Math.max(1, maximum - minimum);
  const slot = plot.width / Math.max(1, bins.length);
  const baseline = plot.top + plot.height * (1 - (0 - minimum) / span);
  return bins.map((bin, index) => {
    const valueY = plot.top + plot.height * (1 - (histogramValue(bin) - minimum) / span);
    const x = plot.left + index * slot;
    const geometry = {
      x,
      y: Math.min(baseline, valueY),
      width: Math.min(slot, Math.max(1, slot - 1)),
      height: histogramValue(bin) === 0 ? 0 : Math.max(1, Math.abs(baseline - valueY)),
    };
    return bin.kind === 'category'
      ? { kind: bin.kind, category: bin.category, value: bin.value, count: bin.count, label: bin.label, geometry }
      : { kind: bin.kind, start: bin.start, end: bin.end, count: bin.count, label: bin.label, ...(bin.boundary ? { boundary: bin.boundary } : {}), geometry };
  });
}

function trimNumber(value: number): string { return Number(value.toFixed(6)).toString(); }

function statusError(kind: ChartDataStatus['kind'], code: ChartDataStatus['code'], message: string): ChartDataStatus { return { kind, code, message }; }

function baseLayout(payload: ChartDrawingPayload, data: ResolvedChartData, width: number, height: number, kind: ChartLayout['kind']): ChartLayout {
  const titleText = payload.elements.titleText?.linkedFormula !== undefined
    ? payload.elements.title
    : payload.elements.titleText?.text ?? payload.elements.title;
  const title = titleText ? { text: titleText, x: 16, y: 12 } : undefined;
  const legend = payload.elements.legend?.visible ? { visible: true, position: payload.elements.legend.position } : { visible: false, position: 'bottom' as const };
  const dataTable = payload.elements.dataTable;
  const tableRequested = dataTable?.visible === true && kind === 'cartesian';
  const visibleSeriesCount = data.series.reduce((count, series, index) => count + (seriesModelFor(payload, series, index)?.visible === false ? 0 : 1), 0);
  const requestedFontSize = dataTable?.font?.fontSize ?? 9;
  const invalidFontSize = tableRequested && (!Number.isFinite(requestedFontSize) || requestedFontSize <= 0);
  const tableFont = dataTable?.font;
  const tableBorder = dataTable?.border;
  const tableFill = tableFont?.fill;
  const tableFillTransparency = tableFill && typeof tableFill !== 'string' ? tableFill.transparency ?? 0 : 0;
  const invalidTableStyle = tableRequested && (
    (tableFont?.rotation !== undefined && !Number.isFinite(tableFont.rotation))
    || (tableBorder?.width !== undefined && (!Number.isFinite(tableBorder.width) || tableBorder.width < 0))
    || (tableBorder?.transparency !== undefined && (!Number.isFinite(tableBorder.transparency) || tableBorder.transparency < 0 || tableBorder.transparency > 100))
    || (tableFill && typeof tableFill !== 'string' && (!Number.isFinite(tableFillTransparency) || tableFillTransparency < 0 || tableFillTransparency > 100))
    || (tableFill && typeof tableFill !== 'string' && tableFill.kind === 'solid' && !tableFont?.color && !tableFill.color)
  );
  const unsupportedTableStyle = tableRequested && (
    (tableFont?.rotation !== undefined && tableFont.rotation !== 0)
    || (tableBorder?.transparency !== undefined && tableBorder.transparency !== 0)
    || (tableFill && typeof tableFill !== 'string' && (tableFill.kind !== 'solid' || tableFillTransparency !== 0))
  );
  const tableRowHeight = invalidFontSize ? 12 : Math.max(12, requestedFontSize + 4);
  const tableHeight = tableRequested ? (visibleSeriesCount + 1) * tableRowHeight : 0;
  const plotWidth = Math.max(10, width - 70 - (legend.position === 'right' ? 100 : 0));
  const legendColumnWidth = Math.min(112, plotWidth * 0.32);
  let categoryCount = Math.max(1, data.categories.length);
  if (tableRequested) {
    for (const [index, series] of data.series.entries()) {
      if (seriesModelFor(payload, series, index)?.visible !== false) categoryCount = Math.max(categoryCount, series.values.length);
    }
  }
  const tableTooNarrow = tableRequested && visibleSeriesCount > 0
    && (plotWidth - legendColumnWidth) / categoryCount < requestedFontSize + 8;
  const topReserve = title ? 62 : 42;
  const bottomLegendReserve = legend.visible && legend.position === 'bottom' ? 24 : 0;
  const tableReserve = tableRequested ? 32 + tableHeight : 0;
  const requestedPlotHeight = height - topReserve - bottomLegendReserve - tableReserve;
  const status = data.status.kind !== 'ready' ? data.status
    : invalidFontSize || invalidTableStyle ? statusError('invalid', 'INVALID_CHART_SOURCE', 'INVALID_CHART_SOURCE: chart data table contains invalid font or border values')
      : unsupportedTableStyle ? statusError('unsupported', 'UNSUPPORTED_FEATURE', 'Chart data table text rotation, fill, or border transparency is not supported')
        : tableRequested && requestedPlotHeight < 10
          ? statusError('unsupported', 'UNSUPPORTED_FEATURE', 'Chart data table does not fit within the drawing bounds')
          : tableTooNarrow ? statusError('unsupported', 'UNSUPPORTED_FEATURE', 'Chart data table categories do not fit within the drawing bounds')
          : data.status;
  const plot = { left: 52, top: title ? 40 : 22, width: plotWidth, height: Math.max(10, requestedPlotHeight) };
  return { status, width, height, plot, title, legend, series: [], kind };
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
        bars.push({ index, category, start, end, x: chartType === 'bar' ? plot.left + barStart * plot.width : plot.left + slotIndex * slot + slot * 0.14 + offset, y: chartType === 'bar' ? plot.top + slotIndex * slot + slot * 0.14 + offset : plot.top + (1 - barEnd) * plot.height, width: chartType === 'bar' ? Math.max(1, (barEnd - barStart) * plot.width) : Math.max(1, band - 1), height: chartType === 'bar' ? Math.max(1, band - 1) : Math.max(1, (barEnd - barStart) * plot.height), color: colorFor(series, seriesIndex), visible });
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
    const seriesIndex = visibleSeriesIndexes[ringIndex];
    if (seriesIndex === undefined) continue;
    const series = data.series[seriesIndex];
    if (!series) continue;
    const values = series.values.map((raw) => {
      const value = chartNumericValue(raw);
      return value === undefined && payload.elements.emptyCells !== 'zero' ? null : Math.max(0, value ?? 0);
    });
    const total = values.reduce<number>((sum, value) => sum + (value ?? 0), 0);
    if (total <= 0) continue;
    const ringWidth = maxRadius * (1 - hole) / ringCount;
    let angle = -Math.PI / 2 + ((payload.subtype === 'exploded-pie' || payload.subtype === 'exploded-three-dimensional-pie' || payload.subtype === 'exploded-doughnut') ? Math.PI / 18 : 0);
    const labels = seriesModelFor(payload, series, seriesIndex)?.dataLabels ?? payload.elements.dataLabels;
    for (let pointIndex = 0; pointIndex < values.length; pointIndex += 1) {
      const value = values[pointIndex] ?? 0;
      if (value <= 0) continue;
      const sweep = value / total * Math.PI * 2;
      const innerRadius = payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * ringIndex : 0;
      const outerRadius = payload.chartType === 'doughnut' ? maxRadius * hole + ringWidth * (ringIndex + 1) : maxRadius;
      const explosion = payload.subtype?.includes('exploded') ? Math.min(12, maxRadius * 0.08) : 0;
      const label = String(data.categories[pointIndex] ?? pointIndex + 1);
      const slice: ChartPieSliceLayout = { seriesIndex, pointIndex, value, startAngle: angle, endAngle: angle + sweep, innerRadius, outerRadius, explosion, color: DEFAULT_COLORS[pointIndex % DEFAULT_COLORS.length]!, label };
      if (labels?.visible) {
        const parts: string[] = [];
        if (labels.showSeriesName) parts.push(series.name);
        if (labels.showCategoryName) parts.push(label);
        if (labels.showValue !== false) parts.push(String(value));
        if (labels.showPercentage) parts.push(`${Math.round(value / total * 10000) / 100}%`);
        if (parts.length) {
          const text = parts.join(labels.separator ?? ', ');
          const midpoint = angle + sweep / 2;
          const offsetRadius = labels.position === 'outside-end' ? outerRadius + 12
            : labels.position === 'inside-base' ? innerRadius + (outerRadius - innerRadius) * 0.35
              : labels.position === 'center' ? (innerRadius + outerRadius) / 2
                : innerRadius + (outerRadius - innerRadius) * 0.72;
          const centerX = plot.left + plot.width / 2 + Math.cos(midpoint) * explosion;
          const centerY = plot.top + plot.height / 2 + Math.sin(midpoint) * explosion;
          const dataLabelX = centerX + Math.cos(midpoint) * offsetRadius;
          const dataLabelY = centerY + Math.sin(midpoint) * offsetRadius;
          const halfWidth = Math.max(8, text.length * 3.2);
          slice.dataLabelText = text;
          slice.dataLabelX = dataLabelX;
          slice.dataLabelY = dataLabelY;
          slice.dataLabelBounds = { left: dataLabelX - halfWidth, top: dataLabelY - 7, right: dataLabelX + halfWidth, bottom: dataLabelY + 7 };
        }
      }
      slices.push(slice);
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
    return [{ seriesIndex, minimum: values[0] ?? 0, lowerWhisker: inliers[0] ?? values[0] ?? 0, q1, median, q3, upperWhisker: inliers.at(-1) ?? values.at(-1) ?? 0, maximum: values.at(-1) ?? 0, mean: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0, innerPoints: options.showInnerPoints === true ? inliers : [], showMeanMarker: options.showMeanMarkers === true, outliers: options.showOutlierPoints === false ? [] : values.filter((value) => value < lowerFence || value > upperFence), color: colors[seriesIndex % colors.length]! }];
  });
}

function waterfallLayouts(payload: ChartDrawingPayload, data: ResolvedChartData, seriesIndex: number, plot: ChartLayout['plot']): ChartWaterfallBarLayout[] {
  const values = data.series[seriesIndex]?.values.map(chartNumericValue) ?? [];
  const options: ChartWaterfallOptions = payload.waterfallOptions ?? { connectorLines: true };
  const totals = new Set(options.totalPointIndexes ?? []);
  let running = 0;
  const bars = values.map((value, index) => {
    if (value === undefined) return { seriesIndex, index, start: running, end: running, connectorValue: running, total: false, color: '#cbd5e1', visible: false };
    const total = totals.has(index);
    const connectorValue = running;
    const from = total ? 0 : running;
    const to = total ? value : running + value;
    running = to;
    return { seriesIndex, index, start: Math.min(from, to), end: Math.max(from, to), connectorValue, total, color: total ? '#64748b' : value >= 0 ? '#10b981' : '#ef4444', visible: true };
  });
  let minimum = 0;
  let maximum = 1;
  for (const bar of bars) {
    minimum = Math.min(minimum, bar.start, bar.end);
    maximum = Math.max(maximum, bar.start, bar.end);
  }
  const span = Math.max(1, maximum - minimum);
  const slot = plot.width / Math.max(1, bars.length);
  const y = (value: number): number => plot.top + plot.height * (1 - (value - minimum) / span);
  return bars.map(({ connectorValue, ...bar }) => {
    const top = y(bar.end);
    const bottom = y(bar.start);
    const x = plot.left + bar.index * slot + slot * 0.16;
    const previousX = plot.left + (bar.index - 1) * slot + slot * 0.16;
    return {
      ...bar,
      geometry: { x, y: Math.min(top, bottom), width: slot * 0.68, height: Math.max(1, Math.abs(bottom - top)) },
      ...(bar.index > 0 ? { connector: { startX: previousX + slot * 0.68, endX: x, y: y(connectorValue) } } : {}),
    };
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
  let minimum = Infinity;
  let maximum = -Infinity;
  for (const value of values) {
    if (value === null || value === undefined || !Number.isFinite(value)) continue;
    minimum = Math.min(minimum, value);
    maximum = Math.max(maximum, value);
  }
  if (minimum === Infinity) {
    minimum = 0;
    maximum = 1;
  }
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
  if (layout.status.kind !== 'ready') return layout;
  if (payload.elements.dataTable?.visible && kind !== 'cartesian') {
    layout.status = statusError('unsupported', 'UNSUPPORTED_FEATURE', `Chart data tables are not supported for ${payload.chartType} charts`);
    return layout;
  }
  if (kind === 'histogram') return buildHistogramLayout(payload, data, layout);
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
  if (['waterfall', 'funnel', 'stock', 'map'].includes(payload.chartType)
    && layout.series.filter((series) => series.visible).length > 1) {
    layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', `${payload.chartType} charts require exactly one visible series`);
    return layout;
  }
  if (payload.elements.dataTable?.visible) {
    let categoryCount = Math.max(1, data.categories.length);
    const tableSeries: Array<{ name: string; color: string; values: readonly PivotScalar[] }> = [];
    for (let index = 0; index < layout.series.length; index += 1) {
      const series = layout.series[index]!;
      if (!series.visible) continue;
      const source = data.series[index]!;
      categoryCount = Math.max(categoryCount, source.values.length);
      tableSeries.push({ name: series.name, color: series.color, values: source.values });
    }
    const showLegendKeys = payload.elements.dataTable.showLegendKeys !== false;
    const legendColumnWidth = Math.min(112, layout.plot.width * 0.32);
    const fontSize = payload.elements.dataTable.font?.fontSize ?? 9;
    const rowHeight = Math.max(12, fontSize + 4);
    layout.dataTable = {
      bounds: {
        left: layout.plot.left,
        top: layout.plot.top + layout.plot.height + 24,
        width: layout.plot.width,
        height: (tableSeries.length + 1) * rowHeight,
      },
      rowHeight,
      categoryCount,
      legendColumnWidth,
      showLegendKeys,
      categories: data.categories,
      series: tableSeries,
    };
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
  if (kind === 'box-whisker') {
    layout.boxes = boxLayouts(payload, data, DEFAULT_COLORS);
    if (layout.boxes.length === 0 || !data.series.some((series, seriesIndex) => layout.series[seriesIndex]?.visible && numberValues(series.values).length > 0)) {
      layout.status = statusError('invalid', 'INVALID_CHART_SOURCE', 'Box and whisker charts require at least one numeric value');
    }
    return layout;
  }
  if (kind === 'waterfall') {
    layout.waterfallBars = waterfallLayouts(payload, data, specialSeriesIndex, layout.plot);
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
    else {
      layout.stockPoints = stockLayouts(data, specialSeriesIndex);
      const stockSubtype = layout.series[specialSeriesIndex]?.subtype ?? payload.subtype;
      if (stockSubtype.includes('volume')) {
        let maximum = 1;
        for (const point of layout.stockPoints ?? []) maximum = Math.max(maximum, point.volume ?? 0);
        layout.stockVolume = {
          maximum,
          top: layout.plot.top + layout.plot.height * 0.78,
          height: layout.plot.height * 0.18,
          priceHeight: layout.plot.height * 0.72,
        };
      }
    }
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
    const centerX = layout.plot.left + layout.plot.width / 2;
    const centerY = layout.plot.top + layout.plot.height / 2;
    const radius = Math.min(layout.plot.width, layout.plot.height) * 0.42;
    const valueAxis = layout.valueAxis!;
    layout.radar = {
      count,
      centerX,
      centerY,
      radius,
      points: visibleSeries.map(({ series, seriesIndex }) => ({
        seriesIndex,
        color: colorFor(series, seriesIndex),
        vertices: Array.from({ length: count }, (_, index) => {
          const value = chartNumericValue(series.values[index]);
          const angle = -Math.PI / 2 + Math.PI * 2 * index / count;
          const pointRadius = value === undefined ? 0 : radius * scale(value, valueAxis);
          return {
            index,
            x: centerX + Math.cos(angle) * pointRadius,
            y: centerY + Math.sin(angle) * pointRadius,
            visible: value !== undefined,
          };
        }),
      })),
    };
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
