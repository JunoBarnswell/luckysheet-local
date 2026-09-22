import { chartSeriesSupportsErrorBars, chartSeriesSupportsTrendlines, chartStackingForSubtype, type ChartDrawingPayload, type ChartSeriesModel, type FormulaValue } from './domain';
import type { RangeRef } from './index';

/** Build role-complete explicit series for chart families whose data roles cannot be inferred by renderers. */
export function buildExplicitChartSeries(type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype'], range: RangeRef, orientation: 'rows' | 'columns' = 'columns'): ChartSeriesModel[] | undefined {
  const byRows = orientation === 'rows';
  const dataStart = (byRows ? range.startRow : range.startColumn) + 1;
  const width = (byRows ? range.endRow - range.startRow : range.endColumn - range.startColumn);
  const vector = (index: number): RangeRef => byRows
    ? { ...range, startRow: index, endRow: index, startColumn: range.startColumn + 1 }
    : { ...range, startColumn: index, endColumn: index, startRow: range.startRow + 1 };
  if (width < 1 || (byRows ? range.endColumn <= range.startColumn : range.endRow <= range.startRow)) {
    throw new Error('INVALID_CHART_SOURCE: 图表数据区域必须包含表头及至少一个数据点');
  }
  if (type === 'scatter' || type === 'bubble') {
    const requiredColumns = type === 'bubble' ? 3 : 2;
    if (width < requiredColumns || width % requiredColumns !== 0) {
      throw new Error(`INVALID_CHART_SOURCE: ${type} chart data columns must form complete groups of ${requiredColumns}`);
    }
    return Array.from({ length: width / requiredColumns }, (_, index) => {
      const offset = dataStart + index * requiredColumns;
      const xRange = vector(offset);
      const yRange = vector(offset + 1);
      const sizeRange = type === 'bubble' ? vector(offset + 2) : undefined;
      return { id: `series:${index + 1}`, name: `Series ${index + 1}`, range: yRange, xRange, yRange, ...(sizeRange ? { sizeRange } : {}), chartType: type, subtype };
    });
  }
  if (type === 'stock') {
    const needsOpen = subtype.includes('open');
    const needsVolume = subtype.includes('volume');
    const requiredColumns = 3 + (needsOpen ? 1 : 0) + (needsVolume ? 1 : 0);
    if (width !== requiredColumns) {
      throw new Error(`INVALID_CHART_SOURCE: ${subtype} stock chart requires exactly a category column plus ${requiredColumns} role columns`);
    }
    let column = dataStart;
    const takeVector = (): RangeRef => vector(column++);
    const volume = needsVolume ? takeVector() : undefined;
    const open = needsOpen ? takeVector() : undefined;
    const high = takeVector();
    const low = takeVector();
    const close = takeVector();
    return [{
      id: 'series:1',
      name: 'Stock',
      range: close,
      stockRoles: { ...(open ? { open } : {}), high, low, close, ...(volume ? { volume } : {}) },
      chartType: type,
      subtype,
    }];
  }
  if (type === 'combo') {
    if (width < 1) throw new Error('INVALID_CHART_SOURCE: combo chart requires a category column and at least one value column');
    const comboTypes: Array<Exclude<ChartDrawingPayload['chartType'], 'combo'>> = subtype === 'stacked-area-clustered-column'
      ? ['area', 'column']
      : subtype === 'clustered-column-line' || subtype === 'clustered-column-line-secondary' ? ['column', 'line'] : ['column'];
    return Array.from({ length: width }, (_, index) => ({
      id: `series:${index + 1}`,
      name: `Series ${index + 1}`,
      range: vector(dataStart + index),
      chartType: comboTypes[index % comboTypes.length]!,
      axis: subtype === 'clustered-column-line-secondary' && index % comboTypes.length === 1 ? 'secondary' : 'primary',
    }));
  }
  return undefined;
}

/** Retarget a canonical payload and remove semantics owned only by the previous chart family. */
export function retargetChartPayload(payload: ChartDrawingPayload, type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype'], stacked?: ChartDrawingPayload['stacked']): ChartDrawingPayload {
  const next = structuredClone(payload);
  const specialized = ['scatter', 'bubble', 'stock', 'combo'].includes(type);
  if (specialized && (payload.chartType !== type || payload.subtype !== subtype)) {
    if (next.source.kind !== 'worksheet-ranges') throw new Error(`UNSUPPORTED_FEATURE: ${type} chart conversion requires worksheet ranges with explicit roles`);
    next.series = next.source.ranges.flatMap((range, rangeIndex) => (buildExplicitChartSeries(type, subtype, range, next.dataOrientation ?? 'columns') ?? []).map((series, seriesIndex) => ({
      ...series,
      id: next.source.kind === 'worksheet-ranges' && next.source.ranges.length === 1 ? series.id : `${series.id ?? 'series'}:${rangeIndex + 1}:${seriesIndex + 1}`,
    })));
  }
  next.chartType = type;
  next.subtype = subtype;
  const effectiveStack = stacked ?? chartStackingForSubtype(subtype);
  if (effectiveStack) next.stacked = effectiveStack;
  else delete next.stacked;
  if (type !== 'map') delete next.mapOptions;
  if (type !== 'histogram' && type !== 'pareto') delete next.histogramOptions;
  if (type !== 'box-whisker') delete next.boxWhiskerOptions;
  if (type !== 'waterfall') delete next.waterfallOptions;
  next.series = next.series?.map((series) => {
    const result = structuredClone(series);
    result.chartType = type === 'combo'
      ? result.chartType && ['column', 'bar', 'line', 'area'].includes(result.chartType) ? result.chartType : 'column'
      : type;
    result.subtype = type === 'combo' ? undefined : subtype;
    if (type !== 'scatter' && type !== 'bubble') {
      delete result.xRange;
      delete result.yRange;
      delete result.sizeRange;
    } else if (type !== 'bubble') delete result.sizeRange;
    if (type !== 'stock') delete result.stockRoles;
    const seriesType = result.chartType ?? type;
    if (!chartSeriesSupportsTrendlines(seriesType)) delete result.trendlines;
    if (!chartSeriesSupportsErrorBars(seriesType)) delete result.errorBars;
    return result;
  });
  return next;
}

/** A declared series is a data vector, never a table whose first cell may be discarded. */
export function validateChartVector(range: RangeRef): void {
  if (!range.sheetId || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(value => Number.isSafeInteger(value) && value >= 0)
    || range.endRow < range.startRow || range.endColumn < range.startColumn
    || (range.startRow !== range.endRow && range.startColumn !== range.endColumn)) {
    throw new Error('INVALID_CHART_SOURCE: 图表系列和分类范围必须是有效的单行或单列');
  }
}

/**
 * The only worksheet-range header/series resolver, shared by display and native export.
 * Automatic tables use a header row and label column (transposed for row orientation).
 * Explicit ranges already identify data: no header inference or value-dependent trimming.
 * This is a projection; it never stores a second chart model or mutates the payload.
 */
export function resolveWorksheetChartRanges(
  payload: ChartDrawingPayload,
  readHeader: (range: RangeRef) => FormulaValue,
): { categoryRange: RangeRef; series: ChartSeriesModel[] } {
  if (payload.source.kind !== 'worksheet-ranges' || !payload.source.ranges.length) {
    throw new Error('INVALID_CHART_SOURCE: 图表缺少工作表数据区域');
  }
  const source = payload.source.ranges[0]!;
  const rows = payload.dataOrientation === 'rows';
  const categoryRange = payload.categoryRange ?? (rows
    ? { ...source, endRow: source.startRow, startColumn: Math.min(source.startColumn + 1, source.endColumn) }
    : { ...source, endColumn: source.startColumn, startRow: Math.min(source.startRow + 1, source.endRow) });
  validateChartVector(categoryRange);
  if (payload.series?.length) {
    for (const entry of payload.series) {
      const ranges = [entry.range, entry.xRange, entry.yRange, entry.sizeRange, entry.categoryRange,
        entry.errorBars?.plusRange, entry.errorBars?.minusRange, ...Object.values(entry.stockRoles ?? {})];
      for (const range of ranges) if (range) validateChartVector(range);
    }
    return { categoryRange, series: payload.series };
  }
  const series: ChartSeriesModel[] = [];
  for (const range of payload.source.ranges) {
    if (!range.sheetId || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(value => Number.isSafeInteger(value) && value >= 0)
      || range.endRow < range.startRow || range.endColumn < range.startColumn
      || (rows ? range.endColumn === range.startColumn : range.endRow === range.startRow)) {
      throw new Error('INVALID_CHART_SOURCE: 自动系列需要表头及至少一个数据点');
    }
    const first = rows ? range.startRow : range.startColumn;
    const last = rows ? range.endRow : range.endColumn;
    for (let index = first === last ? first : first + 1; index <= last; index += 1) {
      const header: RangeRef = rows
        ? { ...range, startRow: index, endRow: index, endColumn: range.startColumn }
        : { ...range, startColumn: index, endColumn: index, endRow: range.startRow };
      const value = readHeader(header);
      const name = value === null || value === '' ? `Series ${series.length + 1}` : typeof value === 'object' ? value.code : String(value);
      const values: RangeRef = rows
        ? { ...range, startRow: index, endRow: index, startColumn: range.startColumn + 1 }
        : { ...range, startColumn: index, endColumn: index, startRow: range.startRow + 1 };
      series.push({ id: `series:${range.sheetId}:${rows ? 'row' : 'column'}:${header.startRow}:${header.startColumn}`, name, range: values });
    }
  }
  return { categoryRange, series };
}
