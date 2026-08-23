import type { ChartDrawingPayload, ChartModel } from '@react-sheets/core-model';
import type { ChartInsertParams } from '@react-sheets/pro-features';

const CHART_INSERT_TYPES = new Set<ChartModel['type']>(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']);

export function resolveChartInsertCommandId(type: ChartModel['type']): string {
  return CHART_INSERT_TYPES.has(type) ? `chart.insert.${type}` : 'chart.insert';
}

export function buildChartInsertParams(chart: ChartModel, drawingId: string): ChartInsertParams {
  return {
    sheetId: chart.sheetId,
    chartId: chart.id,
    drawingId,
    bounds: { ...chart.bounds },
    payload: buildChartDrawingPayload(chart),
  };
}

export function buildChartDrawingPayload(chart: ChartModel): ChartDrawingPayload {
  return {
    kind: 'chart',
    chartId: chart.id,
    chartType: chart.type,
    title: chart.title,
    pivotId: chart.pivotId,
    sourceRanges: structuredClone(chart.sourceRanges),
    series: chart.series?.map((entry) => ({
      name: entry.name,
      range: structuredClone(entry.range),
      color: entry.color,
    })),
    categoryRange: chart.categoryRange ? structuredClone(chart.categoryRange) : undefined,
    legendPosition: chart.legendPosition ?? 'bottom',
    showDataLabels: chart.showDataLabels ?? false,
  };
}

export function buildChartMetadataPatch(chart: ChartModel): Partial<ChartDrawingPayload> | undefined {
  const patch: Partial<ChartDrawingPayload> = {};
  if (chart.pivotId) patch.pivotId = chart.pivotId;
  if (chart.series) {
    patch.series = chart.series.map((entry) => ({
      name: entry.name,
      range: structuredClone(entry.range),
      color: entry.color,
    }));
  }
  if (chart.categoryRange) patch.categoryRange = structuredClone(chart.categoryRange);
  if (chart.legendPosition) patch.legendPosition = chart.legendPosition;
  if (chart.showDataLabels != null) patch.showDataLabels = chart.showDataLabels;
  return Object.keys(patch).length > 0 ? patch : undefined;
}
