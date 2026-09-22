import type { ChartDrawingPayload, ChartSeriesModel, FormulaValue } from './domain';
import type { RangeRef } from './index';

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
