import { createPivotMemberKey, formatPivotMember, pivotMemberKey, type PivotModel, type PivotResultNode, type PivotResultTree, type PivotScalar, type WorkbookModel, type WorksheetModel } from '@react-sheets/core-model';
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

export interface PivotChartCategory {
  id: string;
  path: string[];
  label: string;
}

export interface PivotChartSeries {
  id: string;
  name: string;
  columnPath: PivotScalar[];
  valueFieldId?: string;
  valueIndex: number;
  values: PivotScalar[];
}

export interface PivotChartData {
  categories: PivotChartCategory[];
  series: PivotChartSeries[];
}

/**
 * Project the live Pivot result matrix into the chart's stable category and
 * series identities.  The renderer must not infer Pivot semantics from a
 * single result cell; every visible Column path and Values placement is a
 * Cartesian series, while the complete typed Row path is the category key.
 */
export function buildPivotChartData(tree: PivotResultTree, pivot?: PivotModel): PivotChartData {
  const leaves: Array<{ node: PivotResultNode; path: string[] }> = [];
  const collectLeaves = (nodes: readonly PivotResultNode[], parentPath: string[] = []): void => {
    for (const node of nodes) {
      const path = node.path?.length ? pivotRowPathLabels(node, tree) : [...parentPath, node.label];
      if (node.children.length > 0) collectLeaves(node.children, path);
      else leaves.push({ node, path });
    }
  };
  collectLeaves(tree.rows);

  const categories = leaves.map(({ node, path }, index) => {
    return { id: node.path?.join('|') ?? node.nodeId ?? `row:${index}`, path, label: path.join(' / ') || node.label || `Row ${index + 1}` };
  });
  const columnPaths: PivotScalar[][] = [];
  const seenColumns = new Set<string>();
  const addColumnPath = (path: PivotScalar[]): void => {
    const key = pivotPathKey(path);
    if (seenColumns.has(key)) return;
    seenColumns.add(key);
    columnPaths.push([...path]);
  };
  for (const path of tree.columnPaths) addColumnPath(path);
  for (const { node } of leaves) for (const cell of node.values) addColumnPath(cell.columnPath);
  if (columnPaths.length === 0) columnPaths.push([]);

  const valueFields = tree.valueFields ?? pivot?.layout.values.map((field) => ({ ...field, sourceFieldId: field.fieldId })) ?? [];
  const valueCount = Math.max(valueFields.length, ...leaves.flatMap(({ node }) => node.values.map((cell) => cell.values.length)), 0);
  const series: PivotChartSeries[] = [];
  for (const [columnIndex, columnPath] of columnPaths.entries()) {
    const columnCaption = columnPath.map(pivotScalarLabel).join(' / ');
    for (let valueIndex = 0; valueIndex < valueCount; valueIndex += 1) {
      const field = valueFields[valueIndex];
      const valueFieldId = field?.sourceFieldId ?? field?.fieldId ?? pivot?.layout.values[valueIndex]?.fieldId;
      const valueCaption = field?.displayName ?? fieldName(valueFieldId, tree);
      const name = columnCaption ? `${columnCaption} ${valueCaption}` : valueCaption;
      series.push({
        id: `${pivotPathKey(columnPath)}|value:${valueIndex}:${valueFieldId ?? 'unknown'}`,
        name,
        columnPath: [...columnPath],
        ...(valueFieldId ? { valueFieldId } : {}),
        valueIndex,
        values: leaves.map(({ node }) => {
          const cell = node.values.find((candidate) => pivotPathKey(candidate.columnPath) === pivotPathKey(columnPath));
          return cell?.values[valueIndex] ?? null;
        }),
      });
    }
  }
  return { categories, series };
}

function pivotRowPathLabels(node: PivotResultNode, tree: PivotResultTree): string[] {
  if (!node.path?.length || (node.path.length === 1 && node.path[0] === '__root__')) return [node.label];
  return node.path.map((segment) => {
    const separator = segment.indexOf('=');
    if (separator <= 0) return segment;
    const fieldId = segment.slice(0, separator);
    const memberToken = segment.slice(separator + 1);
    const field = tree.fields.fields.find((candidate) => candidate.fieldId === fieldId);
    const matchingValue = field?.values?.find((value) => pivotMemberKey(createPivotMemberKey(value)) === memberToken);
    if (matchingValue !== undefined) return pivotScalarLabel(matchingValue);
    return pivotMemberTokenLabel(memberToken);
  });
}

function pivotMemberTokenLabel(token: string): string {
  if (token === 'blank:null') return '(blank)';
  const separator = token.indexOf(':');
  if (separator <= 0) return token;
  const value = token.slice(separator + 1);
  try { return pivotScalarLabel(JSON.parse(value) as PivotScalar); } catch { return value; }
}

function pivotScalarLabel(value: PivotScalar): string { return formatPivotMember(value); }
function pivotPathKey(path: readonly PivotScalar[]): string { return path.map((value) => pivotMemberKey(createPivotMemberKey(value))).join('|'); }
function fieldName(fieldId: string | undefined, tree: PivotResultTree): string { return tree.fields.fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId ?? 'Value'; }

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

function resolvePivotData(payload: ChartPayload, tree: PivotResultTree): ResolvedChartData {
  const projected = buildPivotChartData(tree);
  const declared = payload.series ?? [];
  const series: ResolvedChartSeries[] = projected.series.map((entry, index) => {
    const declaredSeries = declared[index];
    return {
      name: declaredSeries?.name ?? entry.name,
      values: [...entry.values],
      color: declaredSeries?.color,
      axis: declaredSeries?.axis ?? 'primary',
      chartType: declaredSeries?.chartType,
    };
  });
  return { categories: projected.categories.map((category) => category.label), series, source: 'pivot' };
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
