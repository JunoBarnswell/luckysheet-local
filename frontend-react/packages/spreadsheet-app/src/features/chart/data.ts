import {
  createPivotMemberKey,
  formatPivotMember,
  pivotMemberKey,
  type ChartDrawingPayload,
  type ChartSource,
  type ChartSubtype,
  type PivotModel,
  type PivotResultNode,
  type PivotResultTree,
  type PivotScalar,
  type RangeRef,
  type WorkbookModel,
  type WorkbookTableModel,
} from '@react-sheets/core-model';
import type { ChartPayload, ChartSeries } from './commands';

export type ChartDataSourceKind = 'range' | 'pivot' | 'table' | 'report-range';

export interface ResolvedChartSeries {
  id: string;
  name: string;
  /** Y values are kept positionally; null/error means a missing data point. */
  values: PivotScalar[];
  xValues?: PivotScalar[];
  sizeValues?: PivotScalar[];
  missing?: boolean[];
  color?: string;
  axis: 'primary' | 'secondary';
  chartType?: Exclude<ChartPayload['chartType'], 'combo'>;
  subtype?: ChartSubtype;
  marker?: ChartSeries['marker'];
  smooth?: boolean;
  trendlines?: ChartSeries['trendlines'];
  errorBars?: ChartSeries['errorBars'];
  errorPlusValues?: PivotScalar[];
  errorMinusValues?: PivotScalar[];
  stockRoles?: ChartSeries['stockRoles'];
  stockValues?: {
    open?: PivotScalar[];
    high: PivotScalar[];
    low: PivotScalar[];
    close: PivotScalar[];
    volume?: PivotScalar[];
  };
}

export interface ChartBindingModel {
  source: ChartDataSourceKind;
  orientation: 'rows' | 'columns';
  categories: PivotScalar[];
  series: ResolvedChartSeries[];
  hierarchyLevels: PivotScalar[][];
  nonContiguous: boolean;
  dynamicRangeIdentity?: string;
  tableStructuredReference?: string;
}

export interface ChartDataStatus {
  kind: 'ready' | 'invalid' | 'unsupported';
  code?: 'INVALID_CHART_SOURCE' | 'UNSUPPORTED_FEATURE' | 'PIVOT_REFERENCE_UNAVAILABLE';
  message?: string;
}

export interface ResolvedChartData {
  categories: PivotScalar[];
  series: ResolvedChartSeries[];
  source: ChartDataSourceKind;
  binding: ChartBindingModel;
  status: ChartDataStatus;
}

export interface StructuredChartSheet {
  getCell(row: number, column: number): { value?: PivotScalar } | undefined;
  hiddenRows: ReadonlySet<number> | readonly number[];
  hiddenColumns: ReadonlySet<number> | readonly number[];
}

export interface StructuredChartSeries {
  id: string;
  name: string;
  values: Array<number | null>;
  categories: string[];
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
 * Project the live Pivot result matrix into stable category and series
 * identities. Every visible column path and value field remains a distinct
 * series; no renderer is allowed to infer Pivot semantics from a cell.
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

  const categories = leaves.map(({ node, path }, index) => ({
    id: node.path?.join('|') ?? node.nodeId ?? `row:${index}`,
    path,
    label: path.join(' / ') || node.label || `Row ${index + 1}`,
  }));
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
  for (const columnPath of columnPaths) {
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

function containsHidden(collection: ReadonlySet<number> | readonly number[], value: number): boolean {
  return 'has' in collection ? collection.has(value) : collection.indexOf(value) >= 0;
}

function isMissing(value: PivotScalar | undefined): boolean {
  return value == null
    || (typeof value === 'object' && value.kind === 'error')
    || (typeof value === 'number' && !Number.isFinite(value))
    || (typeof value === 'string' && value.trim() !== '' && chartNumericValue(value) === undefined);
}

function normalizeEmptyValues(values: PivotScalar[], mode: ChartPayload['elements']['emptyCells']): { values: PivotScalar[]; missing: boolean[] } {
  const normalized = values.map((value) => mode === 'zero' && (value === null || value === undefined || value === '') ? 0 : value);
  return { values: normalized, missing: normalized.map(isMissing) };
}

export function chartNumericValue(value: PivotScalar | undefined): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    const numeric = Number(value.replace(/[$,%]/g, ''));
    return Number.isFinite(numeric) ? numeric : undefined;
  }
  return undefined;
}

function scalarValue(sheet: StructuredChartSheet, row: number, column: number): PivotScalar {
  return sheet.getCell(row, column)?.value ?? null;
}

function readRange(
  sheet: StructuredChartSheet,
  range: RangeRef,
  hiddenData: ChartPayload['elements']['hiddenData'],
): PivotScalar[][] {
  const rows: PivotScalar[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (hiddenData === 'hideRows' && containsHidden(sheet.hiddenRows, row)) continue;
    const values: PivotScalar[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (hiddenData === 'hideColumns' && containsHidden(sheet.hiddenColumns, column)) continue;
      values.push(scalarValue(sheet, row, column));
    }
    rows.push(values);
  }
  return rows;
}

function scalarColumn(sheet: StructuredChartSheet, range: RangeRef, hiddenData: ChartPayload['elements']['hiddenData'], headerRow?: number): PivotScalar[] {
  const rows = readRange(sheet, range, hiddenData);
  const values = rows.map((row) => row[0] ?? null);
  return headerRow === range.startRow && values.length > 0 ? values.slice(1) : values;
}

function seriesName(sheet: StructuredChartSheet, range: RangeRef, fallback: string): string {
  const value = scalarValue(sheet, range.startRow, range.startColumn);
  return value == null || value === '' ? fallback : String(value);
}

function sheetFor(getSheet: (sheetId: string) => StructuredChartSheet | undefined, range: RangeRef): StructuredChartSheet {
  const sheet = getSheet(range.sheetId);
  if (!sheet) throw new Error(`Chart source sheet not found: ${range.sheetId}`);
  return sheet;
}

function chartSeriesFromDeclaration(payload: ChartPayload, declared: ChartSeries, getSheet: (sheetId: string) => StructuredChartSheet | undefined, sourceRange?: RangeRef): ResolvedChartSeries {
  const valueRange = declared.yRange ?? declared.range;
  const sheet = sheetFor(getSheet, valueRange);
  const headerRow = sourceRange?.startRow;
  const normalizedValues = normalizeEmptyValues(scalarColumn(sheet, valueRange, payload.elements.hiddenData, headerRow), payload.elements.emptyCells);
  const values = normalizedValues.values;
  const xValues = declared.xRange ? scalarColumn(sheetFor(getSheet, declared.xRange), declared.xRange, payload.elements.hiddenData, headerRow) : undefined;
  const sizeValues = declared.sizeRange ? scalarColumn(sheetFor(getSheet, declared.sizeRange), declared.sizeRange, payload.elements.hiddenData, headerRow) : undefined;
  const stockValues = declared.stockRoles ? {
    ...(declared.stockRoles.open ? { open: scalarColumn(sheetFor(getSheet, declared.stockRoles.open), declared.stockRoles.open, payload.elements.hiddenData, headerRow) } : {}),
    high: scalarColumn(sheetFor(getSheet, declared.stockRoles.high), declared.stockRoles.high, payload.elements.hiddenData, headerRow),
    low: scalarColumn(sheetFor(getSheet, declared.stockRoles.low), declared.stockRoles.low, payload.elements.hiddenData, headerRow),
    close: scalarColumn(sheetFor(getSheet, declared.stockRoles.close), declared.stockRoles.close, payload.elements.hiddenData, headerRow),
    ...(declared.stockRoles.volume ? { volume: scalarColumn(sheetFor(getSheet, declared.stockRoles.volume), declared.stockRoles.volume, payload.elements.hiddenData, headerRow) } : {}),
  } : undefined;
  return {
    id: declared.id ?? `series:${declared.name}:${valueRange.sheetId}:${valueRange.startRow}:${valueRange.startColumn}`,
    name: declared.name || seriesName(sheet, valueRange, 'Series'),
    values,
    ...(xValues ? { xValues } : {}),
    ...(sizeValues ? { sizeValues } : {}),
    missing: normalizedValues.missing,
    color: declared.color,
    axis: declared.axis ?? 'primary',
    chartType: declared.chartType,
    subtype: declared.subtype,
    marker: declared.marker,
    smooth: declared.smooth,
    trendlines: declared.trendlines,
    errorBars: declared.errorBars,
    ...(declared.errorBars?.plusRange ? { errorPlusValues: scalarColumn(sheetFor(getSheet, declared.errorBars.plusRange), declared.errorBars.plusRange, payload.elements.hiddenData, headerRow) } : {}),
    ...(declared.errorBars?.minusRange ? { errorMinusValues: scalarColumn(sheetFor(getSheet, declared.errorBars.minusRange), declared.errorBars.minusRange, payload.elements.hiddenData, headerRow) } : {}),
    stockRoles: declared.stockRoles,
    ...(stockValues ? { stockValues } : {}),
  };
}

function rangeSourceData(payload: ChartPayload, getSheet: (sheetId: string) => StructuredChartSheet | undefined): { categories: PivotScalar[]; series: ResolvedChartSeries[] } {
  if (payload.source.kind !== 'worksheet-ranges') throw new Error(`Chart source mismatch: expected worksheet-ranges, received ${payload.source.kind}`);
  const sourceRange = payload.source.ranges[0];
  if (!sourceRange) return { categories: [], series: [] };
  const sheet = sheetFor(getSheet, sourceRange);
  const matrix = readRange(sheet, sourceRange, payload.elements.hiddenData);
  const categoryRange = payload.categoryRange;
  const categories = categoryRange
    ? scalarColumn(sheetFor(getSheet, categoryRange), categoryRange, payload.elements.hiddenData, categoryRange.startRow === sourceRange.startRow ? sourceRange.startRow : undefined)
    : payload.dataOrientation === 'rows'
      ? (matrix[0]?.slice(1) ?? [])
      : matrix.slice(1).map((row) => row[0] ?? null);
  const declared = payload.series ?? [];
  if (declared.length > 0) return { categories, series: declared.map((entry) => chartSeriesFromDeclaration(payload, entry, getSheet, sourceRange)) };
  const series: ResolvedChartSeries[] = [];
  const width = sourceRange.endColumn - sourceRange.startColumn + 1;
  if (payload.dataOrientation === 'rows') {
    for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex] ?? [];
      const normalized = normalizeEmptyValues(row.slice(1), payload.elements.emptyCells);
      series.push({ id: `series:${sourceRange.sheetId}:${sourceRange.startRow + rowIndex}`, name: String(row[0] ?? `Series ${series.length + 1}`), values: normalized.values, missing: normalized.missing, axis: 'primary' });
    }
    return { categories, series };
  }
  const appendMatrixSeries = (range: RangeRef, rangeMatrix: PivotScalar[][]): void => {
    const rangeWidth = range.endColumn - range.startColumn + 1;
    if (rangeWidth <= 1) {
      const normalized = normalizeEmptyValues(rangeMatrix.slice(1).map((row) => row[0] ?? null), payload.elements.emptyCells);
      series.push({ id: `series:${range.sheetId}:${range.startColumn}`, name: seriesName(sheetFor(getSheet, range), range, `Series ${series.length + 1}`), values: normalized.values, missing: normalized.missing, axis: 'primary' });
      return;
    }
    for (let columnIndex = 1; columnIndex < rangeWidth; columnIndex += 1) {
      const normalized = normalizeEmptyValues(rangeMatrix.slice(1).map((row) => row[columnIndex] ?? null), payload.elements.emptyCells);
      series.push({ id: `series:${range.sheetId}:${range.startColumn + columnIndex}`, name: String(rangeMatrix[0]?.[columnIndex] ?? `Series ${series.length + 1}`), values: normalized.values, missing: normalized.missing, axis: 'primary' });
    }
  };
  if (width <= 1) {
    const normalized = normalizeEmptyValues(matrix.slice(1).map((row) => row[0] ?? null), payload.elements.emptyCells);
    series.push({ id: 'series:1', name: seriesName(sheet, sourceRange, 'Series 1'), values: normalized.values, missing: normalized.missing, axis: 'primary' });
  } else {
    for (let columnIndex = 1; columnIndex < width; columnIndex += 1) {
      const normalized = normalizeEmptyValues(matrix.slice(1).map((row) => row[columnIndex] ?? null), payload.elements.emptyCells);
      series.push({
        id: `series:${sourceRange.sheetId}:${sourceRange.startColumn + columnIndex}`,
        name: String(matrix[0]?.[columnIndex] ?? `Series ${series.length + 1}`),
        values: normalized.values,
        missing: normalized.missing,
        axis: 'primary',
      });
    }
  }
  for (const range of payload.source.ranges.slice(1)) appendMatrixSeries(range, readRange(sheetFor(getSheet, range), range, payload.elements.hiddenData));
  return { categories, series };
}

function resolvePivotData(payload: ChartPayload, tree: PivotResultTree): { categories: PivotScalar[]; series: ResolvedChartSeries[] } {
  const projected = buildPivotChartData(tree);
  const declared = payload.series ?? [];
  const series: ResolvedChartSeries[] = projected.series.map((entry, index) => {
    const declaredSeries = declared[index];
    const normalized = normalizeEmptyValues([...entry.values], payload.elements.emptyCells);
    const values = normalized.values;
    return {
      id: declaredSeries?.id ?? entry.id,
      name: declaredSeries?.name ?? entry.name,
      values,
      missing: normalized.missing,
      color: declaredSeries?.color,
      axis: declaredSeries?.axis ?? 'primary',
      chartType: declaredSeries?.chartType,
      subtype: declaredSeries?.subtype,
      marker: declaredSeries?.marker,
      smooth: declaredSeries?.smooth,
      trendlines: declaredSeries?.trendlines,
      errorBars: declaredSeries?.errorBars,
      stockRoles: declaredSeries?.stockRoles,
    };
  });
  return { categories: projected.categories.map((category) => category.label), series };
}

function bindingFor(source: ChartSource, categories: PivotScalar[], series: ResolvedChartSeries[], options: Partial<Pick<ChartBindingModel, 'orientation' | 'hierarchyLevels' | 'nonContiguous' | 'dynamicRangeIdentity' | 'tableStructuredReference'>> = {}): ChartBindingModel {
  return {
    source: source.kind === 'worksheet-ranges' ? 'range' : source.kind,
    orientation: options.orientation ?? 'columns',
    categories: structuredClone(categories),
    series: structuredClone(series),
    hierarchyLevels: structuredClone(options.hierarchyLevels ?? []),
    nonContiguous: options.nonContiguous ?? (source.kind === 'worksheet-ranges' && source.ranges.length > 1),
    ...(options.dynamicRangeIdentity ? { dynamicRangeIdentity: options.dynamicRangeIdentity } : {}),
    ...(options.tableStructuredReference ? { tableStructuredReference: options.tableStructuredReference } : {}),
  };
}

function readyData(source: ChartSource, categories: PivotScalar[], series: ResolvedChartSeries[], options?: Parameters<typeof bindingFor>[3]): ResolvedChartData {
  return { categories, series, source: source.kind === 'worksheet-ranges' ? 'range' : source.kind, binding: bindingFor(source, categories, series, options), status: { kind: 'ready' } };
}

function invalidData(source: ChartSource, code: ChartDataStatus['code'], message: string): ResolvedChartData {
  return { categories: [], series: [], source: source.kind === 'worksheet-ranges' ? 'range' : source.kind, binding: bindingFor(source, [], []), status: { kind: code === 'UNSUPPORTED_FEATURE' ? 'unsupported' : 'invalid', code, message } };
}

/** Resolve a canonical chart against a worksheet reader without constructing a second model. */
export function resolveChartDataFromSources(payload: ChartPayload, getSheet: (sheetId: string) => StructuredChartSheet | undefined, pivotResults: Readonly<Record<string, PivotResultTree>> = {}, tables: readonly WorkbookTableModel[] = []): ResolvedChartData {
  try {
    if (payload.source.kind === 'pivot') {
      const tree = pivotResults[payload.source.pivotId];
      if (!tree) return invalidData(payload.source, 'PIVOT_REFERENCE_UNAVAILABLE', `Pivot reference unavailable: ${payload.source.pivotId}`);
      const resolved = resolvePivotData(payload, tree);
      return readyData(payload.source, resolved.categories, resolved.series, { orientation: payload.dataOrientation ?? 'columns' });
    }
    if (payload.source.kind === 'worksheet-ranges') {
      const resolved = rangeSourceData(payload, getSheet);
      return readyData(payload.source, resolved.categories, resolved.series, { orientation: payload.dataOrientation ?? 'columns', dynamicRangeIdentity: payload.source.identity });
    }
    const structured = resolveStructuredChartBindings(payload, tables, getSheet);
    const series = structured.series.map((entry) => ({
      id: entry.id,
      name: entry.name,
      values: entry.values,
      missing: entry.values.map((value) => value === null),
      axis: 'primary' as const,
    }));
    return readyData(payload.source, structured.categories, series, {
      tableStructuredReference: payload.source.kind === 'table' ? payload.source.structuredReference ?? `${payload.source.tableId}[${payload.source.bindings.values.map((entry) => entry.fieldId).join(',')}]` : undefined,
      dynamicRangeIdentity: payload.source.kind === 'report-range' ? payload.source.identity : undefined,
    });
  } catch (error) {
    return invalidData(payload.source, 'INVALID_CHART_SOURCE', error instanceof Error ? error.message : String(error));
  }
}

/** Resolve chart data from the live WorkbookModel for command/unit-test consumers. */
export function resolveChartData(workbook: WorkbookModel, payload: ChartPayload, pivotResults: Readonly<Record<string, PivotResultTree>> = {}): ResolvedChartData {
  const result = resolveChartDataFromSources(
    payload,
    (sheetId) => {
      const sheet = workbook.getSheet(sheetId);
      return { getCell: (row: number, column: number) => sheet.cells.get(row, column), hiddenRows: sheet.hiddenRows, hiddenColumns: sheet.hiddenColumns };
    },
    pivotResults,
    [...workbook.dataModel.tables.values()],
  );
  if (result.status.kind !== 'ready') throw new Error(`${result.status.code ?? 'INVALID_CHART_SOURCE'}: ${result.status.message ?? 'Chart data is unavailable'}`);
  return result;
}

/** Canonical resolver for table/report-backed chart bindings used by model and Canvas. */
export function resolveStructuredChartBindings(payload: ChartDrawingPayload, tables: readonly WorkbookTableModel[], getSheet: (sheetId: string) => StructuredChartSheet | undefined): StructuredChartData {
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
      const numeric = chartNumericValue(sheet.getCell(row, sourceRange.startColumn + field.ordinal)?.value);
      if (numeric !== undefined) byField.set(binding.fieldId, [...(byField.get(binding.fieldId) ?? []), numeric]);
    }
    buckets.set(category, byField);
  }
  const aggregate = (values: number[], mode: typeof source.bindings.values[number]['aggregate']): number | null => {
    if (!values.length || mode === 'none') return values.length ? values[values.length - 1]! : payload.elements.emptyCells === 'zero' ? 0 : null;
    if (mode === 'count') return values.length;
    if (mode === 'min') return Math.min(...values);
    if (mode === 'max') return Math.max(...values);
    if (mode === 'average') return values.reduce((sum, value) => sum + value, 0) / values.length;
    return values.reduce((sum, value) => sum + value, 0);
  };
  let categories = [...buckets.keys()];
  let series = valueBindings.map((binding) => ({
    id: binding.fieldId,
    name: fieldById.get(binding.fieldId)!.name,
    values: categories.map((category) => aggregate(buckets.get(category)?.get(binding.fieldId) ?? [], binding.aggregate)),
    categories: [...categories],
  }));
  const sortBinding = valueBindings.find((binding) => binding.sort);
  if (sortBinding) {
    const sortIndex = series.findIndex((entry) => entry.id === sortBinding.fieldId);
    if (sortIndex >= 0) {
      const order = sortBinding.sort === 'desc' ? -1 : 1;
      const orderIndexes = categories.map((_category, index) => index).sort((left, right) => {
        const leftValue = series[sortIndex]!.values[left];
        const rightValue = series[sortIndex]!.values[right];
        if (leftValue == null && rightValue == null) return 0;
        if (leftValue == null) return -1 * order;
        if (rightValue == null) return 1 * order;
        return (leftValue - rightValue) * order;
      });
      categories = orderIndexes.map((index) => categories[index]!);
      series = series.map((entry) => ({ ...entry, categories: [...categories], values: orderIndexes.map((index) => entry.values[index]!) }));
    }
  }
  return { categories, series };
}

export function normalizeChartSeries(series: ChartSeries[] | undefined): ChartSeries[] | undefined {
  if (!series) return undefined;
  return series.map((entry, index) => ({
    ...structuredClone(entry),
    id: entry.id ?? `series:${index + 1}`,
    axis: entry.axis ?? 'primary',
    chartType: entry.chartType ?? undefined,
    trendlines: entry.trendlines ? structuredClone(entry.trendlines) : undefined,
  }));
}

export interface ResolvedSparklineSeries {
  values: Array<number | null>;
  min: number;
  max: number;
}

/** Sparkline data uses the same positional hidden/empty rules as ChartDomain. */
export function resolveSparklineSeries(
  sparkline: import('@react-sheets/core-model').SparklineModel,
  getSheet: (sheetId: string) => StructuredChartSheet | undefined,
  group?: import('@react-sheets/core-model').SparklineGroup,
): ResolvedSparklineSeries {
  const source = sparkline.sourceRange;
  const sheet = getSheet(source.sheetId);
  if (!sheet) throw new Error(`Unknown sparkline source sheet: ${source.sheetId}`);
  const orientation = group?.dataOrientation ?? sparkline.dataOrientation ?? 'rows';
  const rows: Array<Array<PivotScalar>> = [];
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    if ((group?.hiddenCells ?? sparkline.hiddenCells ?? 'show') === 'hide' && containsHidden(sheet.hiddenRows, row)) continue;
    const values: PivotScalar[] = [];
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      if ((group?.hiddenCells ?? sparkline.hiddenCells ?? 'show') === 'hide' && containsHidden(sheet.hiddenColumns, column)) continue;
      values.push(scalarValue(sheet, row, column));
    }
    rows.push(values);
  }
  const values = orientation === 'columns'
    ? Array.from({ length: Math.max(0, source.endColumn - source.startColumn + 1) }, (_, column) => rows.map((row) => row[column] ?? null)).flat()
    : rows.flat();
  const emptyMode = group?.emptyCells ?? sparkline.emptyCells ?? 'gap';
  const resolved = values.map((value) => {
    const numeric = chartNumericValue(value);
    const empty = value === null || value === undefined || value === '';
    return numeric === undefined ? empty && emptyMode === 'zero' ? 0 : null : numeric;
  });
  const numbers = resolved.filter((value): value is number => value !== null);
  const output = sparkline.rightToLeft || group?.rightToLeft ? resolved.reverse() : resolved;
  return { values: output, min: Math.min(0, ...numbers), max: Math.max(0, ...numbers) };
}
