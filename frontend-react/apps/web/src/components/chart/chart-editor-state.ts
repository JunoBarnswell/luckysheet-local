import { resolveWorksheetChartRanges, type ChartDrawingPayload, type ChartSeriesModel, type RangeRef } from '@react-sheets/core-model';
import { parseRangeInput } from '../../domain/range-input';

export interface ChartSeriesDraft {
  value: ChartSeriesModel;
  range: string;
  xRange: string;
  yRange: string;
  sizeRange: string;
  errorPlusRange: string;
  errorMinusRange: string;
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

/** Convert editor text only at the command boundary; invalid input never reuses an old range. */
export function chartPayloadFromDraft(draft: ChartEditorDraft, sheetId: string): ChartDrawingPayload {
  const payload = structuredClone(draft.value);
  if (payload.source.kind === 'worksheet-ranges') {
    payload.source.ranges = draft.sourceRanges.split(';').map((text, index) => parseChartRange(text, payload.source.kind === 'worksheet-ranges' ? payload.source.ranges[index]?.sheetId ?? sheetId : sheetId, `数据区域 ${index + 1}`));
  } else if (payload.source.kind === 'report-range') {
    payload.source.range = parseChartRange(draft.sourceRanges, payload.source.range.sheetId, '报表区域');
  }
  if (draft.categoryRange.trim()) payload.categoryRange = parseChartRange(draft.categoryRange, draft.base.categoryRange?.sheetId ?? sheetId, '分类标签');
  else delete payload.categoryRange;
  if (draft.series.length || payload.series) payload.series = draft.series.map((entry, index) => {
    const series = structuredClone(entry.value);
    if (!series.name.trim()) throw new Error(`系列 ${index + 1} 的名称不能为空`);
    series.range = parseChartRange(entry.range, series.range.sheetId, `系列 ${index + 1}`);
    for (const key of ['xRange', 'yRange', 'sizeRange'] as const) {
      if (entry[key].trim()) series[key] = parseChartRange(entry[key], series[key]?.sheetId ?? series.range.sheetId, `系列 ${index + 1} ${key}`);
      else delete series[key];
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
    return series;
  });
  if (['scatter', 'bubble', 'stock'].includes(payload.chartType) && !payload.series?.length) {
    throw new Error(`${payload.chartType} 图表必须先生成可编辑系列并设置完整的数据角色`);
  }
  for (const [index, series] of (payload.series ?? []).entries()) {
    const seriesType = series.chartType ?? payload.chartType;
    if ((seriesType === 'scatter' || seriesType === 'bubble') && (!series.xRange || !series.yRange)) throw new Error(`系列 ${index + 1} 必须设置 X 与 Y 范围`);
    if (seriesType === 'bubble' && !series.sizeRange) throw new Error(`系列 ${index + 1} 必须设置气泡大小范围`);
    if (seriesType === 'stock' && !series.stockRoles) throw new Error(`系列 ${index + 1} 必须设置股票图角色范围`);
  }
  for (const axis of [payload.elements.valueAxis, payload.elements.secondaryValueAxis, payload.elements.categoryAxis]) {
    if (!axis) continue;
    if ([axis.minimum, axis.maximum].some(value => value !== undefined && !Number.isFinite(value))) throw new Error('坐标轴边界必须是有限数字');
    if (axis.minimum !== undefined && axis.maximum !== undefined && axis.minimum >= axis.maximum) throw new Error('坐标轴最小值必须小于最大值');
    if (axis.scale === 'logarithmic' && axis.minimum !== undefined && axis.minimum <= 0) throw new Error('对数坐标轴的最小值必须大于 0');
  }
  if (payload.source.kind === 'worksheet-ranges') resolveWorksheetChartRanges(payload, () => null);
  return payload;
}
