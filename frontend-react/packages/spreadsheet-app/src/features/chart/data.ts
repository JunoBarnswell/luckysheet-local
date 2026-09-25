import {
  createPivotMemberKey,
  resolveWorksheetChartRanges,
  validateChartVector,
  formatPivotMember,
  pivotMemberKey,
  chartTextFormulaEntries,
  chartTextFormulaRange,
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
import type { FormulaSheetIdentity } from '@react-sheets/formula-engine';

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
  categories: readonly PivotScalar[];
  series: readonly ResolvedChartSeries[];
  hierarchyLevels: readonly (readonly PivotScalar[])[];
  nonContiguous: boolean;
  dynamicRangeIdentity?: string;
  tableStructuredReference?: string;
}

export interface ChartDataStatus {
  kind: 'ready' | 'loading' | 'invalid' | 'unsupported';
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

export interface ChartFormulaOwnerContext {
  readonly ownerSheetId: string;
  readonly sheetOrder: readonly FormulaSheetIdentity[];
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
  const fieldById = new Map(tree.fields.fields.map((field) => [field.fieldId, field]));
  const memberLabelsByField = new Map<string, Map<string, string>>();
  for (const field of tree.fields.fields) {
    memberLabelsByField.set(field.fieldId, new Map((field.values ?? []).map((value) => [pivotMemberKey(createPivotMemberKey(value)), pivotScalarLabel(value)])));
  }
  const leaves: Array<{ node: PivotResultNode; path: string[] }> = [];
  const collectLeaves = (nodes: readonly PivotResultNode[], parentPath: string[] = []): void => {
    for (const node of nodes) {
      const path = node.path?.length ? pivotRowPathLabels(node, memberLabelsByField) : [...parentPath, node.label];
      if (node.children.length > 0) collectLeaves(node.children, path);
      else leaves.push({ node, path });
    }
  };
  collectLeaves(tree.rows);

  const categories = leaves.map(({ node, path }, index) => ({
    id: node.path?.length ? JSON.stringify(node.path) : node.nodeId ?? `row:${index}`,
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
  let valueCount = valueFields.length;
  const cellsByLeafAndPath = leaves.map(({ node }) => {
    const cells = new Map<string, PivotResultNode['values'][number]>();
    for (const cell of node.values) {
      valueCount = Math.max(valueCount, cell.values.length);
      cells.set(pivotPathKey(cell.columnPath), cell);
    }
    return cells;
  });
  const series: PivotChartSeries[] = [];
  for (const columnPath of columnPaths) {
    const columnPathKey = pivotPathKey(columnPath);
    const columnCaption = columnPath.map(pivotScalarLabel).join(' / ');
    for (let valueIndex = 0; valueIndex < valueCount; valueIndex += 1) {
      const field = valueFields[valueIndex];
      const valueId = field?.valueId ?? pivot?.layout.values[valueIndex]?.valueId;
      const sourceFieldId = field?.sourceFieldId ?? field?.fieldId ?? pivot?.layout.values[valueIndex]?.fieldId;
      const valueCaption = field?.displayName ?? fieldById.get(sourceFieldId ?? '')?.name ?? sourceFieldId ?? 'Value';
      const name = columnCaption ? `${columnCaption} ${valueCaption}` : valueCaption;
      series.push({
        id: `${columnPathKey}|value:${valueIndex}:${valueId ?? 'unknown'}`,
        name,
        columnPath: [...columnPath],
        ...(valueId ? { valueId } : {}),
        valueIndex,
        values: cellsByLeafAndPath.map((cells) => cells.get(columnPathKey)?.values[valueIndex] ?? null),
      });
    }
  }
  return { categories, series };
}

function pivotRowPathLabels(node: PivotResultNode, memberLabelsByField: ReadonlyMap<string, ReadonlyMap<string, string>>): string[] {
  if (!node.path?.length || (node.path.length === 1 && node.path[0] === '__root__')) return [node.label];
  return node.path.map((segment) => {
    const separator = segment.indexOf('=');
    if (separator <= 0) return segment;
    const fieldId = segment.slice(0, separator);
    const memberToken = segment.slice(separator + 1);
    const matchingLabel = memberLabelsByField.get(fieldId)?.get(memberToken);
    if (matchingLabel !== undefined) return matchingLabel;
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
function pivotPathKey(path: readonly PivotScalar[]): string {
  return JSON.stringify(path.map((value) => pivotMemberKey(createPivotMemberKey(value))));
}
function containsHidden(collection: ReadonlySet<number> | readonly number[], value: number): boolean {
  return 'has' in collection ? collection.has(value) : collection.indexOf(value) >= 0;
}

/**
 * Return every worksheet range that can affect a chart projection. Keeping
 * this dependency calculation next to the resolver prevents cache invalidation
 * and active-sheet projection planning from drifting apart as chart families
 * add specialised ranges (stock roles, error bars, and label sources).
 */
export function chartSourceRanges(
  payload: ChartDrawingPayload,
  tables: readonly WorkbookTableModel[] = [],
  owner?: ChartFormulaOwnerContext,
): RangeRef[] {
  const ranges: RangeRef[] = [];
  const seen = new Set<string>();
  const add = (range: RangeRef | undefined): void => {
    if (!range) return;
    const key = `${range.sheetId}:${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`;
    if (seen.has(key)) return;
    seen.add(key);
    ranges.push({ ...range });
  };
  if (payload.source.kind === 'worksheet-ranges') {
    for (const range of payload.source.ranges) add(range);
  } else if (payload.source.kind === 'table') {
    const tableId = payload.source.tableId;
    add(tables.find((table) => table.id === tableId)?.sourceRange);
  } else if (payload.source.kind === 'report-range') {
    add(payload.source.range);
  }
  add(payload.categoryRange);
  for (const series of payload.series ?? []) {
    add(series.range);
    add(series.xRange);
    add(series.yRange);
    add(series.sizeRange);
    add(series.categoryRange);
    add(series.stockRoles?.open);
    add(series.stockRoles?.high);
    add(series.stockRoles?.low);
    add(series.stockRoles?.close);
    add(series.stockRoles?.volume);
    add(series.errorBars?.plusRange);
    add(series.errorBars?.minusRange);
    add(series.dataLabels?.valuesFromCells);
  }
  const textFormulas = chartTextFormulaEntries(payload);
  if (textFormulas.length > 0) {
    if (!owner) throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: chart ${payload.chartId} formula ranges require an owner worksheet`);
    for (const { field } of textFormulas) {
      add(chartTextFormulaRange(payload, field, owner.ownerSheetId, owner.sheetOrder));
    }
  }
  return ranges;
}

/** Resolve a chart title formula from the same worksheet value projection used by chart data. */
export function resolveChartTitleText(
  payload: ChartDrawingPayload,
  owner: ChartFormulaOwnerContext,
  readCellText: (range: RangeRef) => string,
): string | undefined {
  const hasLinkedFormula = payload.elements.titleText?.linkedFormula !== undefined;
  const linkedRange = chartTextFormulaRange(payload, 'titleText.linkedFormula', owner.ownerSheetId, owner.sheetOrder);
  if (hasLinkedFormula) return linkedRange ? readCellText(linkedRange) : '#REF!';
  return payload.elements.titleText?.text ?? payload.elements.title;
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
    const trimmed = value.trim();
    const percent = trimmed.endsWith('%');
    const accounting = trimmed.startsWith('(') && trimmed.endsWith(')');
    const source = accounting ? trimmed.slice(1, -1) : trimmed;
    const numeric = Number(source.replace(/[$,\s]/g, '').replace(/%$/, ''));
    if (!Number.isFinite(numeric)) return undefined;
    const normalized = percent ? numeric / 100 : numeric;
    return accounting ? -normalized : normalized;
  }
  return undefined;
}

function scalarValue(sheet: StructuredChartSheet, row: number, column: number): PivotScalar {
  return sheet.getCell(row, column)?.value ?? null;
}

function scalarVector(sheet: StructuredChartSheet, range: RangeRef, hiddenData: ChartPayload['elements']['hiddenData']): PivotScalar[] {
  validateChartVector(range);
  const values: PivotScalar[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (hiddenData === 'hideRows' && containsHidden(sheet.hiddenRows, row)) continue;
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (hiddenData === 'hideColumns' && containsHidden(sheet.hiddenColumns, column)) continue;
      values.push(scalarValue(sheet, row, column));
    }
  }
  return values;
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

function assertVectorLength(seriesName: string, role: string, values: readonly PivotScalar[] | undefined, expected: number): void {
  if (values && values.length !== expected) throw new Error(`INVALID_CHART_SOURCE: ${seriesName} ${role} range has ${values.length} points; expected ${expected}`);
}

function chartSeriesFromDeclaration(payload: ChartPayload, declared: ChartSeries, getSheet: (sheetId: string) => StructuredChartSheet | undefined): ResolvedChartSeries {
  const valueRange = declared.yRange ?? declared.range;
  const sheet = sheetFor(getSheet, valueRange);
  const normalizedValues = normalizeEmptyValues(scalarVector(sheet, valueRange, payload.elements.hiddenData), payload.elements.emptyCells);
  const values = normalizedValues.values;
  const xValues = declared.xRange ? scalarVector(sheetFor(getSheet, declared.xRange), declared.xRange, payload.elements.hiddenData) : undefined;
  const sizeValues = declared.sizeRange ? scalarVector(sheetFor(getSheet, declared.sizeRange), declared.sizeRange, payload.elements.hiddenData) : undefined;
  const stockValues = declared.stockRoles ? {
    ...(declared.stockRoles.open ? { open: scalarVector(sheetFor(getSheet, declared.stockRoles.open), declared.stockRoles.open, payload.elements.hiddenData) } : {}),
    high: scalarVector(sheetFor(getSheet, declared.stockRoles.high), declared.stockRoles.high, payload.elements.hiddenData),
    low: scalarVector(sheetFor(getSheet, declared.stockRoles.low), declared.stockRoles.low, payload.elements.hiddenData),
    close: scalarVector(sheetFor(getSheet, declared.stockRoles.close), declared.stockRoles.close, payload.elements.hiddenData),
    ...(declared.stockRoles.volume ? { volume: scalarVector(sheetFor(getSheet, declared.stockRoles.volume), declared.stockRoles.volume, payload.elements.hiddenData) } : {}),
  } : undefined;
  const errorPlusValues = declared.errorBars?.plusRange ? scalarVector(sheetFor(getSheet, declared.errorBars.plusRange), declared.errorBars.plusRange, payload.elements.hiddenData) : undefined;
  const errorMinusValues = declared.errorBars?.minusRange ? scalarVector(sheetFor(getSheet, declared.errorBars.minusRange), declared.errorBars.minusRange, payload.elements.hiddenData) : undefined;
  const seriesLabel = declared.name || 'Series';
  assertVectorLength(seriesLabel, 'X', xValues, values.length);
  assertVectorLength(seriesLabel, 'Size', sizeValues, values.length);
  assertVectorLength(seriesLabel, 'positive error', errorPlusValues, values.length);
  assertVectorLength(seriesLabel, 'negative error', errorMinusValues, values.length);
  if (stockValues) {
    const stockLength = stockValues.high.length;
    assertVectorLength(seriesLabel, 'stock Low', stockValues.low, stockLength);
    assertVectorLength(seriesLabel, 'stock Close', stockValues.close, stockLength);
    assertVectorLength(seriesLabel, 'stock Open', stockValues.open, stockLength);
    assertVectorLength(seriesLabel, 'stock Volume', stockValues.volume, stockLength);
  }
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
    ...(errorPlusValues ? { errorPlusValues } : {}),
    ...(errorMinusValues ? { errorMinusValues } : {}),
    stockRoles: declared.stockRoles,
    ...(stockValues ? { stockValues } : {}),
  };
}

function rangeSourceData(payload: ChartPayload, getSheet: (sheetId: string) => StructuredChartSheet | undefined): { categories: PivotScalar[]; series: ResolvedChartSeries[] } {
  const binding = resolveWorksheetChartRanges(payload, range => scalarValue(sheetFor(getSheet, range), range.startRow, range.startColumn));
  const explicitCategoryRanges = binding.series.flatMap((entry) => entry.categoryRange ? [entry.categoryRange] : []);
  const categoryRangeKey = (range: RangeRef): string => `${range.sheetId}:${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`;
  if (explicitCategoryRanges.length > 0 && (explicitCategoryRanges.length !== binding.series.length || explicitCategoryRanges.some((range) => categoryRangeKey(range) !== categoryRangeKey(explicitCategoryRanges[0]!)))) {
    throw new Error('INVALID_CHART_SOURCE: all chart series must share one canonical category range');
  }
  const categoryRange = explicitCategoryRanges[0] ?? binding.categoryRange;
  const categories = scalarVector(sheetFor(getSheet, categoryRange), categoryRange, payload.elements.hiddenData);
  const series = binding.series.filter(entry => {
    if (payload.elements.hiddenData === 'show') return true;
    const range = entry.yRange ?? entry.range;
    const sheet = sheetFor(getSheet, range);
    return !(payload.elements.hiddenData === 'hideColumns' && range.startColumn === range.endColumn && containsHidden(sheet.hiddenColumns, range.startColumn))
      && !(payload.elements.hiddenData === 'hideRows' && range.startRow === range.endRow && containsHidden(sheet.hiddenRows, range.startRow));
  }).map(entry => chartSeriesFromDeclaration(payload, entry, getSheet));
  for (const entry of series) {
    const expected = entry.stockValues?.high.length ?? entry.values.length;
    if (categories.length !== expected) throw new Error(`INVALID_CHART_SOURCE: ${entry.name} has ${expected} points but the category range has ${categories.length}`);
  }
  return { categories, series };
}

function resolvePivotData(payload: ChartPayload, tree: PivotResultTree): { categories: PivotScalar[]; series: ResolvedChartSeries[] } {
  const projected = buildPivotChartData(tree);
  const declared = payload.series ?? [];
  const declaredById = new Map(declared.flatMap((entry) => entry.id ? [[entry.id, entry] as const] : []));
  const declaredByName = new Map<string, ChartSeries[]>();
  for (const entry of declared) {
    const matches = declaredByName.get(entry.name);
    if (matches) matches.push(entry);
    else declaredByName.set(entry.name, [entry]);
  }
  const series: ResolvedChartSeries[] = projected.series.map((entry) => {
    const byName = declaredByName.get(entry.name);
    const declaredSeries = declaredById.get(entry.id) ?? (byName?.length === 1 ? byName[0] : undefined);
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
    // Resolved chart vectors are immutable projection output. Sharing them
    // with the binding avoids duplicating every point for large charts.
    categories,
    series,
    hierarchyLevels: options.hierarchyLevels ?? [],
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

function loadingData(source: ChartSource, message: string): ResolvedChartData {
  return { categories: [], series: [], source: source.kind === 'worksheet-ranges' ? 'range' : source.kind, binding: bindingFor(source, [], []), status: { kind: 'loading', message } };
}

/** Resolve a canonical chart against a worksheet reader without constructing a second model. */
export function resolveChartDataFromSources(payload: ChartPayload, getSheet: (sheetId: string) => StructuredChartSheet | undefined, pivotResults: Readonly<Record<string, PivotResultTree>> = {}, tables: readonly WorkbookTableModel[] = [], loadingPivotIds: ReadonlySet<string> = new Set()): ResolvedChartData {
  try {
    if (payload.source.kind === 'pivot') {
      const tree = pivotResults[payload.source.pivotId];
      if (!tree) {
        if (loadingPivotIds.has(payload.source.pivotId)) return loadingData(payload.source, `Loading PivotTable chart data: ${payload.source.pivotId}`);
        return invalidData(payload.source, 'PIVOT_REFERENCE_UNAVAILABLE', `Pivot reference unavailable: ${payload.source.pivotId}`);
      }
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
    const message = error instanceof Error ? error.message : String(error);
    return invalidData(payload.source, message.startsWith('UNSUPPORTED_FEATURE') ? 'UNSUPPORTED_FEATURE' : 'INVALID_CHART_SOURCE', message);
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
  if (!Number.isSafeInteger(sourceRange.startRow) || !Number.isSafeInteger(sourceRange.endRow)
    || !Number.isSafeInteger(sourceRange.startColumn) || !Number.isSafeInteger(sourceRange.endColumn)
    || sourceRange.startRow < 0 || sourceRange.startColumn < 0
    || sourceRange.endRow < sourceRange.startRow || sourceRange.endColumn < sourceRange.startColumn) {
    throw new Error('INVALID_CHART_SOURCE: structured chart source range is invalid');
  }
  const sheet = getSheet(sourceRange.sheetId);
  if (!sheet) throw new Error(`Chart source sheet not found: ${sourceRange.sheetId}`);
  const fields = source.kind === 'table'
    ? table!.fields.map((field) => ({ id: field.id, name: field.name, ordinal: field.ordinal }))
    : Array.from({ length: sourceRange.endColumn - sourceRange.startColumn + 1 }, (_, offset) => ({ id: `report-column-${offset}`, name: String(sheet.getCell(sourceRange.startRow, sourceRange.startColumn + offset)?.value ?? `Column ${offset + 1}`), ordinal: offset }));
  const sourceWidth = sourceRange.endColumn - sourceRange.startColumn + 1;
  const fieldIds = new Set<string>();
  const fieldOrdinals = new Set<number>();
  for (const field of fields) {
    if (!field.id.trim() || fieldIds.has(field.id)) throw new Error(`INVALID_CHART_SOURCE: duplicate field identity ${field.id}`);
    if (!Number.isSafeInteger(field.ordinal) || field.ordinal < 0 || field.ordinal >= sourceWidth) {
      throw new Error(`INVALID_CHART_SOURCE: field ${field.id} ordinal is outside the source range`);
    }
    if (fieldOrdinals.has(field.ordinal)) throw new Error(`INVALID_CHART_SOURCE: duplicate field ordinal ${field.ordinal}`);
    fieldOrdinals.add(field.ordinal);
    fieldIds.add(field.id);
  }
  const fieldById = new Map(fields.map((field) => [field.id, field]));
  const bindingAreas = ['values', 'category', 'details', 'color', 'size', 'tooltip', 'filter'] as const;
  for (const area of bindingAreas) {
    const bindings = source.bindings[area];
    if (!Array.isArray(bindings)) throw new Error(`INVALID_CHART_SOURCE: ${area} bindings must be an array`);
    for (const binding of bindings) {
      if (!binding || typeof binding.fieldId !== 'string' || binding.fieldId.trim() === '' || binding.area !== area) {
        throw new Error(`INVALID_CHART_SOURCE: ${area} binding identity is invalid`);
      }
      if (!['sum', 'average', 'count', 'min', 'max', 'none'].includes(binding.aggregate)) {
        throw new Error(`INVALID_CHART_SOURCE: unsupported aggregate ${String(binding.aggregate)}`);
      }
      if (binding.sort !== undefined && binding.sort !== 'asc' && binding.sort !== 'desc') {
        throw new Error(`INVALID_CHART_SOURCE: unsupported sort ${String(binding.sort)}`);
      }
    }
  }
  if (source.bindings.category.length > 1) throw new Error('INVALID_CHART_SOURCE: only one category binding is supported');
  if (source.bindings.values.length === 0) throw new Error(`Chart source ${source.kind} has no value bindings`);
  const valueFieldIds = new Set<string>();
  for (const binding of source.bindings.values) {
    if (valueFieldIds.has(binding.fieldId)) throw new Error(`INVALID_CHART_SOURCE: duplicate value binding ${binding.fieldId}`);
    valueFieldIds.add(binding.fieldId);
  }
  if (source.bindings.values.filter((binding) => binding.sort !== undefined).length > 1) {
    throw new Error('INVALID_CHART_SOURCE: only one sorted value binding is supported');
  }
  for (const binding of [...source.bindings.category, ...source.bindings.values, ...source.bindings.details, ...source.bindings.color, ...source.bindings.size, ...source.bindings.tooltip, ...source.bindings.filter]) {
    if (!fieldById.has(binding.fieldId)) throw new Error(`INVALID_CHART_SOURCE: binding field not found: ${binding.fieldId}`);
  }
  for (const area of ['details', 'color', 'size', 'tooltip', 'filter'] as const) {
    if (source.bindings[area].length > 0) throw new Error(`UNSUPPORTED_FEATURE: structured chart binding area ${area} is not supported by the canonical renderer`);
  }
  const hiddenData = payload.elements.hiddenData ?? 'show';
  const hideRows = hiddenData === 'hideRows';
  const hideColumns = hiddenData === 'hideColumns';
  const visible = (field: { ordinal: number } | undefined): boolean => Boolean(field && (!hideColumns || !containsHidden(sheet.hiddenColumns, sourceRange.startColumn + field.ordinal)));
  const categoryBinding = source.bindings.category[0];
  const categoryField = categoryBinding ? fieldById.get(categoryBinding.fieldId) : undefined;
  if (categoryBinding && !categoryField) throw new Error(`INVALID_CHART_SOURCE: binding field not found: ${categoryBinding.fieldId}`);
  if (categoryBinding && categoryBinding.aggregate !== 'none') throw new Error('INVALID_CHART_SOURCE: category bindings cannot aggregate values');
  if (categoryBinding?.sort !== undefined) throw new Error('UNSUPPORTED_FEATURE: category binding sorting is not supported by the canonical renderer');
  if (categoryField && !visible(categoryField)) throw new Error('INVALID_CHART_SOURCE: the bound category field is hidden by the chart visibility policy');
  const valueBindings = source.bindings.values.filter((binding) => visible(fieldById.get(binding.fieldId)));
  if (!valueBindings.length) throw new Error(`Chart source ${source.kind} has no visible numeric value bindings`);
  const rows: Array<{ category: string; categoryKey: string; values: Map<string, number[]> }> = [];
  const buckets = new Map<string, Array<{ category: string; categoryKey: string; values: Map<string, number[]> }>>();
  let categoryOrdinal = 0;
  for (let row = sourceRange.startRow + 1; row <= sourceRange.endRow; row += 1) {
    if (hideRows && containsHidden(sheet.hiddenRows, row)) continue;
    const rawCategory = categoryField ? sheet.getCell(row, sourceRange.startColumn + categoryField.ordinal)?.value ?? '' : ++categoryOrdinal;
    const category = String(rawCategory);
    const categoryKey = `${typeof rawCategory}:${JSON.stringify(rawCategory)}`;
    const byField = new Map<string, number[]>();
    for (const binding of valueBindings) {
      const field = fieldById.get(binding.fieldId);
      if (!field) throw new Error(`Chart binding field not found: ${binding.fieldId}`);
      const rawValue = sheet.getCell(row, sourceRange.startColumn + field.ordinal)?.value ?? null;
      const numeric = chartNumericValue(rawValue);
      if (numeric !== undefined) {
        byField.set(binding.fieldId, [numeric]);
      } else if (binding.aggregate === 'count' && rawValue !== null && rawValue !== '') {
        byField.set(binding.fieldId, [1]);
      }
    }
    const entry = { category, categoryKey, values: byField };
    rows.push(entry);
    const group = buckets.get(categoryKey);
    if (group) group.push(entry);
    else buckets.set(categoryKey, [entry]);
  }
  const aggregate = (values: number[], mode: typeof source.bindings.values[number]['aggregate']): number | null => {
    if (!values.length) return payload.elements.emptyCells === 'zero' ? 0 : null;
    switch (mode) {
      case 'none':
        if (values.length !== 1) throw new Error('INVALID_CHART_SOURCE: none aggregate received multiple values');
        return values[0]!;
      case 'count': return values.length;
      case 'min': return Math.min(...values);
      case 'max': return Math.max(...values);
      case 'average': return values.reduce((sum, value) => sum + value, 0) / values.length;
      case 'sum': return values.reduce((sum, value) => sum + value, 0);
      default: throw new Error(`INVALID_CHART_SOURCE: unsupported aggregate ${String(mode)}`);
    }
  };
  const rowLevel = valueBindings.some((binding) => binding.aggregate === 'none');
  let entries = rowLevel
    ? rows.map((entry) => ({ category: entry.category, rows: [entry], group: buckets.get(entry.categoryKey) ?? [entry] }))
    : [...buckets.values()].map((group) => ({ category: group[0]!.category, rows: group, group }));
  let categories = entries.map((entry) => entry.category);
  let series = valueBindings.map((binding) => ({
    id: binding.fieldId,
    name: fieldById.get(binding.fieldId)!.name,
    values: entries.map((entry) => aggregate(
      (binding.aggregate === 'none' ? entry.rows : entry.group)
        .flatMap((row) => row.values.get(binding.fieldId) ?? []),
      binding.aggregate,
    )),
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
  const hideCells = (group?.hiddenCells ?? sparkline.hiddenCells ?? 'show') === 'hide';
  const visibleColumns: number[] = [];
  for (let column = source.startColumn; column <= source.endColumn; column += 1) {
    if (!hideCells || !containsHidden(sheet.hiddenColumns, column)) visibleColumns.push(column);
  }
  const rows: Array<Array<PivotScalar>> = [];
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    if (hideCells && containsHidden(sheet.hiddenRows, row)) continue;
    const values: PivotScalar[] = [];
    for (const column of visibleColumns) {
      values.push(scalarValue(sheet, row, column));
    }
    rows.push(values);
  }
  const values = orientation === 'columns'
    ? visibleColumns.map((_column, column) => rows.map((row) => row[column] ?? null)).flat()
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
