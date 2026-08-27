import { createPivotMemberKey, formatPivotMember, pivotMemberKey, type ChartDrawingPayload, type PivotModel, type PivotResultNode, type PivotResultTree, type PivotScalar, type WorkbookModel, type WorkbookTableModel, type WorksheetModel } from '@react-sheets/core-model';
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

export interface StructuredChartSheet {
  getCell(row: number, column: number): { value?: PivotScalar } | undefined;
  hiddenRows: ReadonlySet<number> | readonly number[];
  hiddenColumns: ReadonlySet<number> | readonly number[];
}

export interface StructuredChartSeries {
  name: string;
  values: number[];
}

export interface StructuredChartData {
  categories: string[];
  series: StructuredChartSeries[];
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
  valueId?: string;
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
      const valueId = field?.valueId ?? pivot?.layout.values[valueIndex]?.valueId;
      const sourceFieldId = field?.sourceFieldId ?? field?.fieldId ?? pivot?.layout.values[valueIndex]?.fieldId;
      const valueCaption = field?.displayName ?? fieldName(sourceFieldId, tree);
      const name = columnCaption ? `${columnCaption} ${valueCaption}` : valueCaption;
      series.push({
        id: `${pivotPathKey(columnPath)}|value:${valueIndex}:${valueId ?? 'unknown'}`,
        name,
        columnPath: [...columnPath],
        ...(valueId ? { valueId } : {}),
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
  if (payload.source.kind !== 'worksheet-ranges') throw new Error(`Chart source mismatch: expected worksheet-ranges, received ${payload.source.kind}`);
  const sourceRanges = payload.source.ranges;
  const firstRange = sourceRanges[0];
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

  for (const range of sourceRanges) {
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

function containsHidden(collection: ReadonlySet<number> | readonly number[], value: number): boolean {
  return 'has' in collection ? collection.has(value) : collection.indexOf(value) >= 0;
}

/** Canonical resolver for table/report-backed chart bindings used by both the model and canvas projection. */
export function resolveStructuredChartBindings(
  payload: ChartDrawingPayload,
  tables: readonly WorkbookTableModel[],
  getSheet: (sheetId: string) => StructuredChartSheet | undefined,
): StructuredChartData {
  const source = payload.source;
  if (source.kind !== 'table' && source.kind !== 'report-range') throw new Error(`Chart binding resolver received ${source.kind}`);
  const table = source.kind === 'table' ? tables.find((entry) => entry.id === source.tableId) : undefined;
  if (source.kind === 'table' && !table) throw new Error(`Chart table binding not found: ${source.tableId}`);
  const sourceRange = table?.sourceRange ?? (source.kind === 'report-range' ? source.range : undefined);
  if (!sourceRange) throw new Error(`Chart source ${source.kind} has no worksheet-backed range`);
  const sheet = getSheet(sourceRange.sheetId);
  if (!sheet) throw new Error(`Chart source sheet not found: ${sourceRange.sheetId}`);
  const fields = source.kind === 'table'
    ? table!.fields.map((field) => ({ id: field.id, name: field.name, ordinal: field.ordinal }))
    : Array.from({ length: sourceRange.endColumn - sourceRange.startColumn + 1 }, (_, offset) => ({ id: `report-column-${offset}`, name: String(sheet.getCell(sourceRange.startRow, sourceRange.startColumn + offset)?.value ?? `Column ${offset + 1}`), ordinal: offset }));
  const fieldById = new Map(fields.map((field) => [field.id, field]));
  const showHiddenData = payload.elements.hiddenData === 'show';
  const visible = (field: { ordinal: number } | undefined): boolean => Boolean(field && (showHiddenData || !containsHidden(sheet.hiddenColumns, sourceRange.startColumn + field.ordinal)));
  const categoryBinding = source.bindings.category[0];
  const categoryField = categoryBinding && visible(fieldById.get(categoryBinding.fieldId)) ? fieldById.get(categoryBinding.fieldId) : undefined;
  const valueBindings = source.bindings.values.filter((binding) => visible(fieldById.get(binding.fieldId)));
  if (!valueBindings.length) throw new Error(`Chart source ${source.kind} has no visible numeric value bindings`);
  const buckets = new Map<string, Map<string, number[]>>();
  for (let row = sourceRange.startRow + 1; row <= sourceRange.endRow; row += 1) {
    if (!showHiddenData && containsHidden(sheet.hiddenRows, row)) continue;
    const category = String(categoryField ? sheet.getCell(row, sourceRange.startColumn + categoryField.ordinal)?.value ?? '' : row - sourceRange.startRow);
    const byField = buckets.get(category) ?? new Map<string, number[]>();
    for (const binding of valueBindings) {
      const field = fieldById.get(binding.fieldId);
      if (!field) throw new Error(`Chart binding field not found: ${binding.fieldId}`);
      const raw = sheet.getCell(row, sourceRange.startColumn + field.ordinal)?.value;
      if (typeof raw === 'number' && Number.isFinite(raw)) byField.set(binding.fieldId, [...(byField.get(binding.fieldId) ?? []), raw]);
    }
    buckets.set(category, byField);
  }
  const aggregate = (values: number[], mode: typeof source.bindings.values[number]['aggregate']): number => {
    if (!values.length || mode === 'none') return values.length ? values[values.length - 1]! : 0;
    if (mode === 'count') return values.length;
    if (mode === 'min') return Math.min(...values);
    if (mode === 'max') return Math.max(...values);
    if (mode === 'average') return values.reduce((sum, value) => sum + value, 0) / values.length;
    return values.reduce((sum, value) => sum + value, 0);
  };
  const categories = [...buckets.keys()];
  const series = valueBindings.map((binding) => ({ name: fieldById.get(binding.fieldId)!.name, values: categories.map((category) => aggregate(buckets.get(category)?.get(binding.fieldId) ?? [], binding.aggregate)) }));
  const sortBinding = valueBindings.find((binding) => binding.sort);
  if (!sortBinding) return { categories, series };
  const sortIndex = series.findIndex((entry) => entry.name === fieldById.get(sortBinding.fieldId)?.name);
  if (sortIndex < 0) return { categories, series };
  const order = sortBinding.sort === 'desc' ? -1 : 1;
  const orderIndexes = categories.map((_category, index) => index).sort((left, right) => (series[sortIndex]!.values[left]! - series[sortIndex]!.values[right]!) * order);
  return { categories: orderIndexes.map((index) => categories[index]!), series: series.map((entry) => ({ ...entry, values: orderIndexes.map((index) => entry.values[index]!) })) };
}

/** Resolve chart data from live ranges or the canonical, read-only PivotResultTree cache. */
export function resolveChartData(
  workbook: WorkbookModel,
  payload: ChartPayload,
  pivotResults: Readonly<Record<string, PivotResultTree>> = {},
): ResolvedChartData {
  if (payload.source.kind === 'pivot') {
    const tree = pivotResults[payload.source.pivotId];
    if (!tree) throw new Error(`Pivot reference unavailable: ${payload.source.pivotId}`);
    return resolvePivotData(payload, tree);
  }
  if (payload.source.kind === 'table' || payload.source.kind === 'report-range') {
    const structured = resolveStructuredChartBindings(payload, [...workbook.dataModel.tables.values()], (sheetId) => {
      const sheet = workbook.getSheet(sheetId);
      return { getCell: (row: number, column: number) => sheet.cells.get(row, column), hiddenRows: sheet.hiddenRows, hiddenColumns: sheet.hiddenColumns };
    });
    return { categories: structured.categories, series: structured.series.map((entry) => ({ ...entry, axis: 'primary' as const })), source: 'range' };
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
