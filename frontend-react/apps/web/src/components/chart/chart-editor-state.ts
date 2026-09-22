import { chartSeriesSupportsErrorBars, chartSeriesSupportsTrendlines, chartStackingForSubtype, resolveWorksheetChartRanges, type ChartDrawingPayload, type ChartSeriesModel, type RangeRef } from '@react-sheets/core-model';
import { parseRangeInput } from '../../domain/range-input';

export interface ChartSeriesDraft {
  value: ChartSeriesModel;
  range: string;
  xRange: string;
  yRange: string;
  sizeRange: string;
  errorPlusRange: string;
  errorMinusRange: string;
  stockOpenRange: string;
  stockHighRange: string;
  stockLowRange: string;
  stockCloseRange: string;
  stockVolumeRange: string;
}

export interface ChartEditorDraft {
  base: ChartDrawingPayload;
  value: ChartDrawingPayload;
  sourceRanges: string;
  categoryRange: string;
  series: ChartSeriesDraft[];
}

export function formatChartRange(range: RangeRef | undefined): string {
  if (!range) return '';
  const columnLabel = (column: number): string => {
    let label = '';
    for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) label = String.fromCharCode(65 + ((value - 1) % 26)) + label;
    return label;
  };
  return `${columnLabel(range.startColumn)}${range.startRow + 1}:${columnLabel(range.endColumn)}${range.endRow + 1}`;
}

export function chartSeriesDraft(value: ChartSeriesModel): ChartSeriesDraft {
  return {
    value: structuredClone(value),
    range: formatChartRange(value.range),
    xRange: formatChartRange(value.xRange),
    yRange: formatChartRange(value.yRange),
    sizeRange: formatChartRange(value.sizeRange),
    errorPlusRange: formatChartRange(value.errorBars?.plusRange),
    errorMinusRange: formatChartRange(value.errorBars?.minusRange),
    stockOpenRange: formatChartRange(value.stockRoles?.open),
    stockHighRange: formatChartRange(value.stockRoles?.high),
    stockLowRange: formatChartRange(value.stockRoles?.low),
    stockCloseRange: formatChartRange(value.stockRoles?.close),
    stockVolumeRange: formatChartRange(value.stockRoles?.volume),
  };
}

export function chartEditorDraft(payload: ChartDrawingPayload): ChartEditorDraft {
  return {
    base: structuredClone(payload), value: structuredClone(payload),
    sourceRanges: payload.source.kind === 'worksheet-ranges' ? payload.source.ranges.map(formatChartRange).join('; ')
      : payload.source.kind === 'report-range' ? formatChartRange(payload.source.range) : '',
    categoryRange: formatChartRange(payload.categoryRange),
    series: (payload.series ?? []).map(chartSeriesDraft),
  };
}

export function parseChartRange(text: string, sheetId: string, label: string): RangeRef {
  const parsed = parseRangeInput(text.trim(), sheetId);
  if (!parsed) throw new Error(`${label}不是有效范围，请使用 A1:B20 格式`);
  if (parsed.endRow >= 1048576 || parsed.endColumn >= 16384) throw new Error(`${label}超出工作表最大范围`);
  return { sheetId, ...parsed };
}

/** Parse source/category text without requiring specialized series roles; used before materializing explicit series. */
export function chartSourcePayloadFromDraft(draft: ChartEditorDraft, sheetId: string): ChartDrawingPayload {
  const payload = structuredClone(draft.value);
  if (payload.source.kind === 'worksheet-ranges') {
    payload.source.ranges = draft.sourceRanges.split(';').map((text, index) => parseChartRange(text, payload.source.kind === 'worksheet-ranges' ? payload.source.ranges[index]?.sheetId ?? sheetId : sheetId, `数据区域 ${index + 1}`));
  } else if (payload.source.kind === 'report-range') {
    payload.source.range = parseChartRange(draft.sourceRanges, payload.source.range.sheetId, '报表区域');
  }
  if (payload.source.kind === 'worksheet-ranges' && draft.categoryRange.trim()) payload.categoryRange = parseChartRange(draft.categoryRange, draft.base.categoryRange?.sheetId ?? sheetId, '分类标签');
  else delete payload.categoryRange;
  return payload;
}

/** Apply a chart-family transition without retaining bindings or options owned by the previous family. */
export function retargetChartDraft(draft: ChartEditorDraft, chartType: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype']): ChartEditorDraft {
  const next = structuredClone(draft);
  const previousType = next.value.chartType;
  next.value.chartType = chartType;
  next.value.subtype = subtype;
  const stacking = chartStackingForSubtype(subtype);
  if (stacking) next.value.stacked = stacking;
  else delete next.value.stacked;
  if (chartType !== 'map') delete next.value.mapOptions;
  if (chartType !== 'histogram' && chartType !== 'pareto') delete next.value.histogramOptions;
  if (chartType !== 'box-whisker') delete next.value.boxWhiskerOptions;
  if (chartType !== 'waterfall') delete next.value.waterfallOptions;
  if (previousType !== chartType && ['scatter', 'bubble', 'stock', 'combo'].includes(chartType)) {
    next.series = [];
    delete next.value.series;
    return next;
  }
  next.series = next.series.map((entry) => {
    const value = structuredClone(entry.value);
    value.chartType = chartType === 'combo'
      ? value.chartType && ['column', 'bar', 'line', 'area'].includes(value.chartType) ? value.chartType : 'column'
      : chartType;
    value.subtype = chartType === 'combo' ? undefined : subtype;
    if (chartType !== 'scatter' && chartType !== 'bubble') {
      entry.xRange = '';
      entry.yRange = '';
      entry.sizeRange = '';
      delete value.xRange;
      delete value.yRange;
      delete value.sizeRange;
    }
    if (chartType !== 'stock') {
      entry.stockOpenRange = '';
      entry.stockHighRange = '';
      entry.stockLowRange = '';
      entry.stockCloseRange = '';
      entry.stockVolumeRange = '';
      delete value.stockRoles;
    }
    const seriesType = value.chartType ?? chartType;
    if (!chartSeriesSupportsTrendlines(seriesType)) delete value.trendlines;
    if (!chartSeriesSupportsErrorBars(seriesType)) {
      delete value.errorBars;
      entry.errorPlusRange = '';
      entry.errorMinusRange = '';
    }
    return { ...entry, value };
  });
  return next;
}

/** Convert editor text only at the command boundary; invalid input never reuses an old range. */
export function chartPayloadFromDraft(draft: ChartEditorDraft, sheetId: string): ChartDrawingPayload {
  const payload = chartSourcePayloadFromDraft(draft, sheetId);
  if (draft.series.length || payload.series) payload.series = draft.series.map((entry, index) => {
    const series = structuredClone(entry.value);
    if (!series.name.trim()) throw new Error(`系列 ${index + 1} 的名称不能为空`);
    series.range = parseChartRange(entry.range, series.range.sheetId, `系列 ${index + 1}`);
    const seriesType = series.chartType ?? payload.chartType;
    if (seriesType === 'scatter' || seriesType === 'bubble') {
      series.xRange = parseChartRange(entry.xRange, series.xRange?.sheetId ?? series.range.sheetId, `系列 ${index + 1} X 范围`);
      series.yRange = parseChartRange(entry.yRange, series.yRange?.sheetId ?? series.range.sheetId, `系列 ${index + 1} Y 范围`);
      series.range = series.yRange;
      if (seriesType === 'bubble') series.sizeRange = parseChartRange(entry.sizeRange, series.sizeRange?.sheetId ?? series.range.sheetId, `系列 ${index + 1} 气泡大小范围`);
      else delete series.sizeRange;
    } else {
      delete series.xRange;
      delete series.yRange;
      delete series.sizeRange;
    }
    if (seriesType === 'stock') {
      const stockSubtype = series.subtype ?? payload.subtype;
      const needsOpen = stockSubtype.includes('open');
      const needsVolume = stockSubtype.includes('volume');
      const high = parseChartRange(entry.stockHighRange, series.stockRoles?.high.sheetId ?? series.range.sheetId, `系列 ${index + 1} High 范围`);
      const low = parseChartRange(entry.stockLowRange, series.stockRoles?.low.sheetId ?? series.range.sheetId, `系列 ${index + 1} Low 范围`);
      const close = parseChartRange(entry.stockCloseRange, series.stockRoles?.close.sheetId ?? series.range.sheetId, `系列 ${index + 1} Close 范围`);
      const open = needsOpen ? parseChartRange(entry.stockOpenRange, series.stockRoles?.open?.sheetId ?? series.range.sheetId, `系列 ${index + 1} Open 范围`) : undefined;
      const volume = needsVolume ? parseChartRange(entry.stockVolumeRange, series.stockRoles?.volume?.sheetId ?? series.range.sheetId, `系列 ${index + 1} Volume 范围`) : undefined;
      series.stockRoles = { ...(open ? { open } : {}), high, low, close, ...(volume ? { volume } : {}) };
      series.range = close;
    } else {
      delete series.stockRoles;
    }
    if (series.errorBars?.type === 'custom') {
      series.errorBars = {
        ...series.errorBars,
        plusRange: parseChartRange(entry.errorPlusRange, series.errorBars.plusRange?.sheetId ?? series.range.sheetId, `系列 ${index + 1} 自定义正误差范围`),
        minusRange: parseChartRange(entry.errorMinusRange, series.errorBars.minusRange?.sheetId ?? series.range.sheetId, `系列 ${index + 1} 自定义负误差范围`),
      };
    } else if (series.errorBars) {
      delete series.errorBars.plusRange;
      delete series.errorBars.minusRange;
    }
    if ((series.errorBars?.type === 'fixed' || series.errorBars?.type === 'percentage')
      && (!Number.isFinite(series.errorBars.value) || series.errorBars.value! < 0)) {
      throw new Error(`系列 ${index + 1} 的误差线值必须是非负有限数字`);
    }
    return series;
  });
  if (['scatter', 'bubble', 'stock'].includes(payload.chartType) && !payload.series?.length) {
    throw new Error(`${payload.chartType} 图表必须先生成可编辑系列并设置完整的数据角色`);
  }
  for (const [index, series] of (payload.series ?? []).entries()) {
    const seriesType = series.chartType ?? payload.chartType;
    if ((seriesType === 'scatter' || seriesType === 'bubble') && (!series.xRange || !series.yRange)) throw new Error(`系列 ${index + 1} 必须设置 X 与 Y 范围`);
    if (seriesType === 'bubble' && !series.sizeRange) throw new Error(`系列 ${index + 1} 必须设置气泡大小范围`);
    if (seriesType === 'stock' && (!series.stockRoles
      || ((series.subtype ?? payload.subtype).includes('open') && !series.stockRoles.open)
      || ((series.subtype ?? payload.subtype).includes('volume') && !series.stockRoles.volume))) throw new Error(`系列 ${index + 1} 必须设置与股票图子类型匹配的角色范围`);
  }
  for (const axis of [payload.elements.valueAxis, payload.elements.secondaryValueAxis, payload.elements.categoryAxis, payload.elements.secondaryCategoryAxis]) {
    if (!axis) continue;
    if ([axis.minimum, axis.maximum].some(value => value !== undefined && !Number.isFinite(value))) throw new Error('坐标轴边界必须是有限数字');
    if (axis.minimum !== undefined && axis.maximum !== undefined && axis.minimum >= axis.maximum) throw new Error('坐标轴最小值必须小于最大值');
    if (axis.scale === 'logarithmic' && axis.minimum !== undefined && axis.minimum <= 0) throw new Error('对数坐标轴的最小值必须大于 0');
    if (axis.scale === 'logarithmic' && axis.maximum !== undefined && axis.maximum <= 0) throw new Error('对数坐标轴的最大值必须大于 0');
  }
  if (payload.source.kind === 'worksheet-ranges') resolveWorksheetChartRanges(payload, () => null);
  return payload;
}
