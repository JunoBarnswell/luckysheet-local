import type {
  AnalysisChartBinding,
  AnalysisFilterClause,
  AnalysisViewDefinition,
  FormulaErrorCode,
  TableScalar,
  WorkbookTableField,
  WorkbookTableModel,
} from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '../../ui-snapshot';

export interface AnalysisErrorValue {
  kind: 'error';
  code: FormulaErrorCode;
}

export type AnalysisCellValue = TableScalar | AnalysisErrorValue;

export interface AnalysisProjectedRow {
  sourceRow: number;
  values: Readonly<Record<string, AnalysisCellValue>>;
}

export interface AnalysisChartPoint {
  category: AnalysisCellValue;
  series?: AnalysisCellValue;
  value: number | null;
}

export interface AnalysisChartProjection {
  chartId: string;
  points: AnalysisChartPoint[];
  valueFieldId?: string;
  categoryFieldId?: string;
  seriesFieldId?: string;
}

export interface AnalysisViewProjection {
  status: 'ready' | 'unavailable';
  message?: string;
  viewId: string;
  tableId: string;
  sourceRange?: WorkbookTableModel['sourceRange'];
  fields: WorkbookTableField[];
  rows: AnalysisProjectedRow[];
  totalRows: number;
  filteredRows: number;
  errorRows: number;
  charts: AnalysisChartProjection[];
}

function isErrorValue(value: AnalysisCellValue): value is AnalysisErrorValue {
  return typeof value === 'object' && value !== null && value.kind === 'error';
}

function scalarFromCell(sheet: CanvasSheetSnapshot, row: number, column: number): AnalysisCellValue {
  const cell = sheet.getCell(row, column);
  const raw = cell?.formulaValue !== undefined ? cell.formulaValue : cell?.rawValue;
  if (raw === null || typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') return raw;
  if (typeof raw === 'object' && raw !== null && 'kind' in raw && raw.kind === 'error' && 'code' in raw) {
    return { kind: 'error', code: raw.code };
  }
  return null;
}

function equalValue(left: AnalysisCellValue, right: TableScalar): boolean {
  return !isErrorValue(left) && left === right;
}

function numberValue(value: AnalysisCellValue): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  return undefined;
}

function matchesFilter(value: AnalysisCellValue, filter: AnalysisFilterClause): boolean {
  if (isErrorValue(value)) return false;
  switch (filter.operator) {
    case 'equals': return equalValue(value, filter.values[0] ?? null);
    case 'not-equals': return !equalValue(value, filter.values[0] ?? null);
    case 'contains': return typeof value === 'string' && value.toLocaleLowerCase().includes(String(filter.values[0] ?? '').toLocaleLowerCase());
    case 'in': return filter.values.some((candidate) => equalValue(value, candidate));
    case 'between': {
      const lower = filter.values[0];
      const upper = filter.values[1];
      const numeric = numberValue(value);
      const lowerNumeric = lower === undefined ? undefined : numberValue(lower);
      const upperNumeric = upper === undefined ? undefined : numberValue(upper);
      if (numeric !== undefined && lowerNumeric !== undefined && upperNumeric !== undefined) return numeric >= lowerNumeric && numeric <= upperNumeric;
      if (typeof value === 'string' && typeof lower === 'string' && typeof upper === 'string') return value >= lower && value <= upper;
      return false;
    }
  }
}

function chartBindingProjection(binding: AnalysisChartBinding, rows: readonly AnalysisProjectedRow[]): AnalysisChartProjection {
  const categoryFieldId = binding.fieldMap.category;
  const seriesFieldId = binding.fieldMap.series;
  const valueFieldId = binding.fieldMap.value;
  const points: AnalysisChartPoint[] = [];
  for (const row of rows) {
    const category = categoryFieldId ? row.values[categoryFieldId] ?? null : row.sourceRow;
    const series = seriesFieldId ? row.values[seriesFieldId] : undefined;
    const value = valueFieldId ? numberValue(row.values[valueFieldId] ?? null) ?? null : null;
    points.push({ category, ...(series === undefined ? {} : { series }), value });
  }
  return {
    chartId: binding.chartId,
    points,
    ...(valueFieldId ? { valueFieldId } : {}),
    ...(categoryFieldId ? { categoryFieldId } : {}),
    ...(seriesFieldId ? { seriesFieldId } : {}),
  };
}

/**
 * Projects a persisted analysis view from the same worksheet-backed table
 * rows consumed by chart and table renderers.  The projection never invents
 * rows for block-backed sources: callers receive an explicit unavailable
 * state until the authoritative data block is available.
 */
export function buildAnalysisViewProjection(
  view: AnalysisViewDefinition,
  table: WorkbookTableModel | undefined,
  sourceSheet: CanvasSheetSnapshot | undefined,
): AnalysisViewProjection {
  const empty = (status: AnalysisViewProjection['status'], message: string): AnalysisViewProjection => ({
    status,
    message,
    viewId: view.id,
    tableId: view.tableId,
    ...(table?.sourceRange ? { sourceRange: structuredClone(table.sourceRange) } : {}),
    fields: table ? structuredClone(table.fields) : [],
    rows: [],
    totalRows: 0,
    filteredRows: 0,
    errorRows: 0,
    charts: [],
  });
  if (!table) return empty('unavailable', `分析视图引用的表不存在：${view.tableId}`);
  if (!table.sourceRange) return empty('unavailable', '当前表的数据块尚未提供工作表投影，无法在浏览器中预览。');
  if (!sourceSheet || sourceSheet.id !== table.sourceRange.sheetId) return empty('unavailable', `分析视图源工作表不可用：${table.sourceRange.sheetId}`);

  const rows: AnalysisProjectedRow[] = [];
  let errorRows = 0;
  for (let sourceRow = table.sourceRange.startRow + 1; sourceRow <= table.sourceRange.endRow; sourceRow += 1) {
    const values: Record<string, AnalysisCellValue> = {};
    let rowHasError = false;
    for (const field of table.fields) {
      const value = scalarFromCell(sourceSheet, sourceRow, table.sourceRange.startColumn + field.ordinal);
      values[field.id] = value;
      rowHasError ||= isErrorValue(value);
    }
    if (rowHasError) errorRows += 1;
    rows.push({ sourceRow, values });
  }
  const filtered = rows.filter((row) => view.filters.every((filter) => matchesFilter(row.values[filter.fieldId] ?? null, filter)));
  return {
    status: 'ready',
    viewId: view.id,
    tableId: view.tableId,
    sourceRange: structuredClone(table.sourceRange),
    fields: structuredClone(table.fields),
    rows,
    totalRows: rows.length,
    filteredRows: filtered.length,
    errorRows,
    charts: view.charts.map((binding) => chartBindingProjection(binding, filtered)),
  };
}
