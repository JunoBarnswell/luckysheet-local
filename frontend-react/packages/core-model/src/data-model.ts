import type { RangeRef, SheetId } from './index';

export type TableScalar = string | number | boolean | null;

export type TableFieldType = 'text' | 'number' | 'boolean' | 'date' | 'mixed';

export type RecordFieldCalculation =
  | { kind: 'formula'; formula: string }
  | { kind: 'lookup'; relationshipId: string; targetFieldId: string; direction: 'forward' | 'reverse' }
  | { kind: 'rollup'; relationshipId: string; targetFieldId: string; direction: 'forward' | 'reverse'; aggregate: 'SUM' | 'AVERAGE' | 'COUNT' | 'COUNTA' | 'MIN' | 'MAX' | 'PRODUCT' };

export interface RecordFieldAddress { tableId: string; recordId: string; fieldId: string }

export interface WorkbookTableField {
  id: string;
  name: string;
  ordinal: number;
  type: TableFieldType;
  calculation?: RecordFieldCalculation;
}

export interface WorkbookTableBlock {
  id: string;
  tableId: string;
  startRow: number;
  rowCount: number;
  storageKey: string;
  encoding: 'typed-column';
}

export interface WorkbookTableModel {
  id: string;
  name: string;
  /** Block-backed query table source. Rows stay outside WorkbookSnapshot. */
  sourceId?: string;
  /** The immutable identity column in the canonical source, independent of display order. */
  recordIdFieldId?: string;
  sourceSheetId?: SheetId;
  /** The canonical source range when rows remain sheet-backed in local mode. */
  sourceRange?: RangeRef;
  rowCount: number;
  fields: WorkbookTableField[];
  blockSize: number;
  blocks: WorkbookTableBlock[];
  revision: number;
}

export interface DataRelationship {
  id: string;
  fromTableId: string;
  fromFieldId: string;
  toTableId: string;
  toFieldId: string;
  cardinality: 'one-to-one' | 'one-to-many' | 'many-to-one';
}

export interface DataViewField {
  fieldId: string;
  caption: string;
  formula?: string;
  widthPx?: number;
}

export interface DataViewDefinition {
  id: string;
  name: string;
  tableId: string;
  fields: DataViewField[];
  groupBy?: string[];
  sort?: Array<{ fieldId: string; direction: 'asc' | 'desc' }>;
  /** Present only for the persisted analysis/dashboard view variant. */
  kind?: 'analysis';
  filters?: AnalysisFilterClause[];
  charts?: AnalysisChartBinding[];
  layout?: AnalysisViewLayout;
  revision?: number;
}

export type AnalysisFilterOperator = 'equals' | 'not-equals' | 'contains' | 'in' | 'between';

export interface AnalysisFilterClause {
  id: string;
  fieldId: string;
  operator: AnalysisFilterOperator;
  values: TableScalar[];
}

export interface AnalysisChartFieldMap {
  category?: string;
  series?: string;
  value?: string;
  color?: string;
  size?: string;
  tooltip?: string;
}

export interface AnalysisChartBinding {
  chartId: string;
  fieldMap: AnalysisChartFieldMap;
}

export interface AnalysisViewLayout {
  columns: number;
  rowHeightPx: number;
  gapPx: number;
}

/**
 * Persisted dashboard/analysis state. It intentionally extends the existing
 * data view shape so older table views remain valid and the workbook keeps a
 * single canonical `dataModel.views` collection.
 */
export interface AnalysisViewDefinition extends DataViewDefinition {
  kind: 'analysis';
  filters: AnalysisFilterClause[];
  charts: AnalysisChartBinding[];
  layout: AnalysisViewLayout;
  revision: number;
}

export function isAnalysisViewDefinition(value: DataViewDefinition): value is AnalysisViewDefinition {
  return value.kind === 'analysis'
    && Array.isArray(value.filters)
    && Array.isArray(value.charts)
    && typeof value.layout === 'object'
    && value.layout !== null
    && Number.isSafeInteger(value.revision)
    && Number(value.revision) >= 0;
}

/** Canonical structured-data graph consumed by TableSheet, GanttSheet and Chart sources. */
export interface ExternalLinkBinding {
  id: string;
  /** Formula workbook qualifier; stable until an explicit binding edit. */
  token: string;
  sourceUnitId: string;
  sheets: Array<{ token: string; sheetId: string }>;
}

export interface WorkbookDataModel {
  sources: import('./data-source').DataSourceManifest[];
  tables: WorkbookTableModel[];
  relationships: DataRelationship[];
  views: DataViewDefinition[];
  externalLinks: ExternalLinkBinding[];
}

export interface TableSheetColumn {
  fieldId: string;
  caption: string;
  widthPx?: number;
  type?: TableFieldType | 'formula' | 'lookup' | 'checkbox' | 'select' | 'currency' | 'percent' | 'barcode';
  formula?: string;
}

export interface TableSheetDefinition {
  viewId: string;
  columns: TableSheetColumn[];
  grouping: Array<{ fieldId: string; collapsed?: boolean }>;
  sortState?: Array<{ fieldId: string; direction: 'asc' | 'desc' }>;
}

export interface GanttTaskFieldMap {
  id: string;
  title: string;
  start: string;
  end: string;
  progress: string;
  parentId?: string;
  dependencies?: string;
}

export interface GanttSheetDefinition {
  viewId: string;
  fieldMap: GanttTaskFieldMap;
  calendar: { workingDays: number[]; dayStartHour: number; dayEndHour: number };
  timeline: { unit: 'day' | 'week' | 'month' | 'quarter'; start?: string; end?: string };
  dependencyStyle: { color: string; width: number };
}

export interface ReportBinding {
  cell: { row: number; column: number };
  expression: string;
  kind: 'static' | 'field' | 'formula' | 'group' | 'summary';
  direction?: 'vertical' | 'horizontal';
  fill?: 'none' | 'down' | 'right';
  summary?: 'sum' | 'count' | 'average' | 'min' | 'max';
}

export interface ReportLayoutDefinition {
  orientation: 'portrait' | 'landscape';
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
}

export interface ReportDataEntryRule {
  fieldId: string;
  writable: boolean;
  required?: boolean;
}

export interface ReportSheetDefinition {
  templateSheetId: SheetId;
  tableId?: string;
  bindings: ReportBinding[];
  pagination: { enabled: boolean; rowsPerPage?: number; repeatHeaderRows?: number[] };
  renderMode: 'design' | 'preview' | 'paginated';
  layout: ReportLayoutDefinition;
  dataEntry: ReportDataEntryRule[];
}

export function assertExternalLinkBinding(link: ExternalLinkBinding): void {
  if (!link || Object.keys(link).some(key => !['id', 'token', 'sourceUnitId', 'sheets'].includes(key))
    || [link.id, link.token, link.sourceUnitId].some(value => typeof value !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(value))
    || !Array.isArray(link.sheets) || link.sheets.length < 1 || link.sheets.length > 1000) throw new Error('EXTERNAL_LINK_BINDING_INVALID');
  const aliases = new Set<string>(), ids = new Set<string>();
  for (const sheet of link.sheets) {
    if (!sheet || Object.keys(sheet).some(key => !['token', 'sheetId'].includes(key)) || typeof sheet.token !== 'string' || !sheet.token.trim() || sheet.token.length > 255
      || typeof sheet.sheetId !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(sheet.sheetId)
      || aliases.has(sheet.token.toUpperCase()) || ids.has(sheet.sheetId)) throw new Error('EXTERNAL_LINK_SHEET_BINDING_INVALID');
    aliases.add(sheet.token.toUpperCase()); ids.add(sheet.sheetId);
  }
}
