import type { PivotResultNode, PivotResultTree, PivotScalar, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import type { ChartPayload, ChartSeries } from './commands';

export interface ResolvedChartSeries {
  name: string;
  values: PivotScalar[];
  color?: string;
  axis: 'primary' | 'secondary';
  chartType?: Exclude<ChartPayload['chartType'], 'combo'>;
}

export interface ResolvedChartData {
  categories: PivotScalar[];
  series: ResolvedChartSeries[];
  source: 'range' | 'pivot';
}

function cellValue(sheet: WorksheetModel, row: number, column: number): PivotScalar {
  return sheet.cells.get(row, column)?.value ?? null;
}

function rangeValues(sheet: WorksheetModel, range: { startRow: number; endRow: number; startColumn: number; endColumn: number }, hiddenData: ChartPayload['elements']['hiddenData'] = 'show'): PivotScalar[][] {
  const rows: PivotScalar[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (hiddenData === 'hideRows' && sheet.hiddenRows.has(row)) continue;
    const values: PivotScalar[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (hiddenData === 'hideColumns' && sheet.hiddenColumns.has(column)) continue;
      values.push(cellValue(sheet, row, column));
    }
    rows.push(values);
  }
  return rows;
}

function seriesName(sheet: WorksheetModel, range: { startRow: number; startColumn: number }, fallback: string): string {
  const value = cellValue(sheet, range.startRow, range.startColumn);
  return value == null || value === '' ? fallback : String(value);
}

function resolveRangeData(workbook: WorkbookModel, payload: ChartPayload): ResolvedChartData {
  const firstRange = payload.sourceRanges[0];
  if (!firstRange) return { categories: [], series: [], source: 'range' };
  const firstSheet = workbook.getSheet(firstRange.sheetId);
  const firstRows = rangeValues(firstSheet, firstRange, payload.elements.hiddenData);
  const declaredSeries = payload.series ?? [];
  const firstXRange = declaredSeries.find((entry) => entry.xRange)?.xRange;
  const categories = payload.categoryRange
    ? rangeValues(workbook.getSheet(payload.categoryRange.sheetId), payload.categoryRange, payload.elements.hiddenData).map((row) => row[0] ?? null)
    : firstXRange
      ? rangeValues(workbook.getSheet(firstXRange.sheetId), firstXRange, payload.elements.hiddenData).slice(1).map((row) => row[0] ?? null)
    : firstRows.slice(1).map((row) => row[0] ?? null);

  const series: ResolvedChartSeries[] = [];
  if (declaredSeries.length > 0) {
    for (const declared of declaredSeries) {
      const valueRange = declared.yRange ?? declared.range;
      const rows = rangeValues(workbook.getSheet(valueRange.sheetId), valueRange, payload.elements.hiddenData);
      series.push({
        name: declared.name || seriesName(workbook.getSheet(valueRange.sheetId), valueRange, 'Series'),
        values: rows.slice(1).map((row) => row[0] ?? null),
        color: declared.color,
        axis: declared.axis ?? 'primary',
        chartType: declared.chartType,
      });
    }
    return { categories, series, source: 'range' };
  }

  for (const range of payload.sourceRanges) {
    const sheet = workbook.getSheet(range.sheetId);
    const rows = rangeValues(sheet, range, payload.elements.hiddenData);
    const width = range.endColumn - range.startColumn + 1;
    if (width <= 1) {
      series.push({ name: seriesName(sheet, range, `Series ${series.length + 1}`), values: rows.slice(1).map((row) => row[0] ?? null), axis: 'primary' });
      continue;
    }
    for (let columnIndex = 1; columnIndex < width; columnIndex += 1) {
      series.push({
        name: String(rows[0]?.[columnIndex] ?? `Series ${series.length + 1}`),
        values: rows.slice(1).map((row) => row[columnIndex] ?? null),
        axis: 'primary',
      });
    }
  }
  return { categories, series, source: 'range' };
}

function leafNodes(nodes: PivotResultNode[], prefix: string[] = []): Array<{ label: string; node: PivotResultNode }> {
  const result: Array<{ label: string; node: PivotResultNode }> = [];
  for (const node of nodes) {
    const path = [...prefix, node.label];
    if (node.children.length === 0) result.push({ label: path.join(' / '), node });
    else result.push(...leafNodes(node.children, path));
  }
  return result;
}

function resolvePivotData(payload: ChartPayload, tree: PivotResultTree): ResolvedChartData {
  const leaves = leafNodes(tree.rows);
  const categories = leaves.map((entry) => entry.label);
  const width = Math.max(tree.columnPaths.length, leaves.reduce((max, entry) => Math.max(max, entry.node.values[0]?.values.length ?? 0), 0));
  const declared = payload.series ?? [];
  const series: ResolvedChartSeries[] = [];
  for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
    const declaredSeries = declared[columnIndex];
    series.push({
      name: declaredSeries?.name ?? (tree.columnPaths[columnIndex]?.map((value) => String(value ?? '(blank)')).join(' / ') || `Series ${columnIndex + 1}`),
      values: leaves.map((entry) => entry.node.values[0]?.values[columnIndex] ?? null),
      color: declaredSeries?.color,
      axis: declaredSeries?.axis ?? 'primary',
      chartType: declaredSeries?.chartType,
    });
  }
  return { categories, series, source: 'pivot' };
}

/** Resolve chart data from live ranges or the canonical, read-only PivotResultTree cache. */
export function resolveChartData(
  workbook: WorkbookModel,
  payload: ChartPayload,
  pivotResults: Readonly<Record<string, PivotResultTree>> = {},
): ResolvedChartData {
  if (payload.pivotId) {
    const tree = pivotResults[payload.pivotId];
    if (tree) return resolvePivotData(payload, tree);
  }
  return resolveRangeData(workbook, payload);
}

export function normalizeChartSeries(series: ChartSeries[] | undefined): ChartSeries[] | undefined {
  if (!series) return undefined;
  return series.map((entry) => ({
    ...structuredClone(entry),
    axis: entry.axis ?? 'primary',
    chartType: entry.chartType ?? undefined,
  }));
}
