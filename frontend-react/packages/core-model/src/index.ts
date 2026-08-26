export type UnitId = string;
export type SheetId = string;
export type Row = number;
export type Column = number;

import type {
  CellHyperlink,
  CellNote,
  CommentThread,
  DrawingObject,
  DrawingPayload,
  ImageCrop,
  ImageEffects,
  SparklineGroup,
  SheetTableModel,
  OutlineModel,
  SpillRange,
  ProtectionRule,
  ProtectionAllow,
  DefinedNameModel,
  DefinedNameScope,
} from './domain';
import { isFormulaError, normalizeDefinedNameModel } from './domain';
import type { FormulaErrorCode } from './domain';
import type { WorkbookDimensionMetrics, WorkbookSnapshot } from './snapshot';
import {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
  type QueryLoadTargetSnapshot,
} from './workbook-state';
import { normalizeFontFamily } from './font-family';

export * from './font-family';
export {
  HORIZONTAL_ALIGNMENTS,
  VERTICAL_ALIGNMENTS,
  READING_ORDERS,
  TEXT_ORIENTATIONS,
  isHorizontalAlignment,
  isVerticalAlignment,
  isReadingOrder,
  type HorizontalAlignment,
  type VerticalAlignment,
  type ReadingOrder,
  type TextOrientation,
  type UnsupportedCellAlignment,
} from './alignment';
import type { HorizontalAlignment, VerticalAlignment, ReadingOrder, TextOrientation, UnsupportedCellAlignment } from './alignment';

export type CellValue = string | number | boolean | null;

export interface CellBorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'double';
  color: string;
}

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
}

export interface CellStyle {
  textRotate?: number;
  textOrientation?: TextOrientation;
  fontFamily?: string;
  /** Font size in 96-DPI CSS pixels. OOXML point sizes are converted at the import boundary. */
  fontSizePx?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  textColor?: string;
  background?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  /** Excel alignment indentation level. One level maps to three rendered spaces. */
  indent?: number;
  /** OOXML readingOrder: context, left-to-right, or right-to-left. */
  readingOrder?: ReadingOrder;
  /** Native alignment values retained explicitly when the editor cannot execute them. */
  unsupportedAlignment?: UnsupportedCellAlignment;
  numberFormat?: string;
  borders?: CellBorders;
  padding?: number;
  locked?: boolean;
  formulaHidden?: boolean;
}

/** Editable cell behavior that can be expressed through the canonical workbook model. */
export type CellEditorKind = 'text' | 'number' | 'date' | 'list' | 'checkbox';

export interface CellEditorConfig {
  kind: CellEditorKind;
  /** List editors use an explicit canonical value list; range/formula lists belong to validation. */
  values?: string[];
  placeholder?: string;
}

export interface CellComment {
  id: string;
  author: string;
  text: string;
  createdAt: string;
  mentions?: string[];
  replies?: CellCommentReply[];
  resolved?: boolean;
  resolvedAt?: string;
}

export interface CellCommentReply {
  id: string;
  author: string;
  text: string;
  createdAt: string;
}

export interface CellData {
  value: CellValue;
  formula?: string;
  displayValue?: string;
  styleId?: string;
  style?: CellStyle;
  /** Workbook-owned editor configuration. It is never a React component payload. */
  editor?: CellEditorConfig;
  presentation?: CellPresentation;
  numberFormat?: string;
  /** Canonical rich text. `value` remains the plain-text projection used by formulas and search. */
  richText?: RichTextRun[];
  /** OOXML formula provenance used to preserve cached values for Excel-only formula families. */
  formulaMetadata?: FormulaMetadata;
  /** 公式引擎结果（含错误）。禁止再用 error: string 当真相 */
  formulaValue?: import('./domain').FormulaValue;
  note?: import('./domain').CellNote;
  comment?: CellComment;
  /** @deprecated prefer hyperlinkDetail */
  hyperlink?: string;
  hyperlinkDetail?: CellHyperlink;
  /** Native AutoFilter color/icon identity resolved at the import boundary. */
  filterMetadata?: {
    color?: { target: 'cell' | 'font'; dxfId?: number; value?: string };
    icon?: { iconSet: string; iconId: number };
  };
}

export const BARCODE_SYMBOLOGIES = ['qr', 'code128', 'code39', 'code93', 'code49', 'codabar', 'ean13', 'ean8', 'upca', 'gs1-128', 'pdf417', 'data-matrix'] as const;
export type BarcodeSymbology = typeof BARCODE_SYMBOLOGIES[number];

export type BarcodeLabelPosition = 'above' | 'below' | 'none';
export type BarcodeParameters =
  | { symbology: 'qr'; errorCorrection?: 'low' | 'medium' | 'quartile' | 'high' }
  | { symbology: 'data-matrix' }
  | { symbology: 'pdf417'; securityLevel?: number }
  | { symbology: 'ean13' | 'ean8' | 'upca'; addOnText?: string; includeCheckDigit?: boolean }
  | { symbology: 'code128' | 'code39' | 'code93' | 'code49' | 'codabar' | 'gs1-128'; fullAscii?: boolean; includeCheckDigit?: boolean; wideNarrowRatio?: number };

export interface BarcodeCellPresentation {
  kind: 'barcode';
  symbology: BarcodeSymbology;
  source: { kind: 'cell-value' } | { kind: 'formula'; formula: string };
  parameters: BarcodeParameters;
  options: { foreground: string; background: string; showText: boolean; labelPosition: BarcodeLabelPosition; quietZone: number; fontSize?: number };
}

export interface ImageCellPresentation {
  kind: 'image';
  src: string;
  altText?: string;
  fit: 'contain' | 'cover' | 'stretch';
  crop?: ImageCrop;
  effects?: ImageEffects;
}

export type CellPresentation = BarcodeCellPresentation | ImageCellPresentation;

export interface RichTextRun {
  text: string;
  style?: Pick<CellStyle, 'fontFamily' | 'fontSizePx' | 'bold' | 'italic' | 'underline' | 'strikethrough' | 'textColor'>;
  /** Names of OOXML run properties retained by the source package but not editable in the canonical model. */
  preservedProperties?: string[];
}

export interface FormulaMetadata {
  kind: 'normal' | 'shared' | 'array' | 'dataTable';
  sharedIndex?: number;
  sharedMaster?: boolean;
  range?: string;
  preservedOnly?: boolean;
  reason?: string;
  sourceFormula?: string;
}

export interface RangeRef {
  sheetId: SheetId;
  startRow: Row;
  endRow: Row;
  startColumn: Column;
  endColumn: Column;
}

export interface MergeSpan {
  range: RangeRef;
  anchor: { row: Row; column: Column };
}

export type WorksheetPane =
  | { kind: 'none' }
  | {
      kind: 'frozen';
      xSplit: number;
      ySplit: number;
      startRow: Row;
      startColumn: Column;
      activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
      state: 'frozen' | 'frozenSplit';
    }
  | {
      kind: 'split';
      /** Native OOXML split positions. They are not row/column counts. */
      xSplit: number;
      ySplit: number;
      startRow: Row;
      startColumn: Column;
      activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
      state: 'split';
    };

export function normalizeWorksheetPane(pane: WorksheetPane): WorksheetPane {
  if (pane.kind === 'none') return { kind: 'none' };
  const activePane = pane.activePane ?? (pane.xSplit > 0 && pane.ySplit > 0 ? 'bottomRight' : pane.xSplit > 0 ? 'topRight' : pane.ySplit > 0 ? 'bottomLeft' : 'topLeft');
  return pane.kind === 'frozen'
    ? { ...pane, activePane, state: pane.state }
    : { ...pane, activePane, state: 'split' };
}

export type {
  SelectionSnapshot,
  EditSession,
  SheetTableModel,
  OutlineGroup,
  OutlineModel,
  DrawingKind,
  DrawingTransform,
  DrawingObject,
  CellHyperlink,
  HyperlinkTarget,
  DrawingPayload,
  ImageDrawingPayload,
  ImageCrop,
  ImageEffects,
  ShapeDrawingPayload,
  ShapeDrawingType,
  ShapeDrawingCategory,
  TextBoxDrawingPayload,
  TextBoxTextFrame,
  TextBoxHorizontalAlignment,
  TextBoxVerticalAlignment,
  TextBoxTextDirection,
  TextBoxAutofit,
  ChartDrawingPayload,
  ChartSeriesType,
  ChartAxisModel,
  ChartGridlineModel,
  ChartAreaStyle,
  ChartMarkerModel,
  ChartTrendlineModel,
  ChartErrorBarsModel,
  ChartDataLabelsModel,
  ChartSeriesModel,
  ChartElementModel,
  DataChartDrawingPayload,
  DataChartPlotType,
  DataChartBindingArea,
  DataChartSource,
  DataChartFieldBinding,
  DataChartInspectorModel,
  CameraDrawingPayload,
  FormControlStyle,
  FormControlCellLink,
  FormControlAction,
  ButtonFormControlPayload,
  SpinButtonFormControlPayload,
  ListBoxFormControlPayload,
  ComboBoxFormControlPayload,
  CheckboxFormControlPayload,
  OptionButtonFormControlPayload,
  GroupBoxFormControlPayload,
  LabelFormControlPayload,
  ScrollbarFormControlPayload,
  FormControlDrawingPayload,
  FormControlType,
  PivotControlFilter,
  PivotControlConnection,
  PivotTimelinePeriod,
  PivotTimelineLevel,
  PivotTimelineFilterType,
  PivotControlStyle,
  PivotSlicerSettings,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  SparklineGroup,
  CellNote,
  CommentThread,
  CommentReply,
  SpillRange,
  SpillState,
  ProtectionRule,
  ProtectionAllow,
  DefinedNameModel,
  DefinedNameScope,
  ProtectionScope,
  FormulaErrorCode,
  FormulaValue,
  StructuralOpKind,
  CellShiftSpec,
  StructuralTransformParams,
} from './domain';
export { createDefaultTextBoxTextFrame } from './domain';
export {
  createEmptySelection,
  isFormulaError,
  createFormulaError,
  normalizeDefinedNameModel,
  isPivotControlFilter,
  isPivotTimelinePeriod,
  isPivotControlStyle,
  isPivotSlicerSettings,
  isPivotSlicerDrawingPayload,
  isPivotTimelineDrawingPayload,
  isFormControlDrawingPayload,
} from './domain';
export { StructuralTransform, planCellShift, type StructuralTransformResult, type CellShiftPlan, ensureDrawing } from './structural-transform';
export {
  planBorderChange,
  isBorderPlacement,
  isBorderLine,
  type BorderPlacement,
  type BorderLine,
  type BorderPlan,
  type BorderPlanCell,
  type BorderPlanBounds,
} from './border-planner';
export { ProtectionResolver, protectionResolver, type ProtectionAction, type ProtectionCellResolution, type ProtectionDecision, type ProtectionResolveRequest } from './protection';
export {
  canonicalExcelDateDayOfWeek,
  canonicalExcelDateFromParts,
  canonicalExcelDateFromSerial,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateFromValue,
  canonicalExcelDatePartsFromSerial,
  canonicalExcelDateToIso,
  canonicalExcelDateToSerial,
  canonicalExcelDateToUtcDate,
  compareCanonicalExcelDates,
  shiftCanonicalExcelDate,
  type CanonicalExcelDate,
  type CanonicalExcelDateParts,
  type ExcelDateEvaluationContext,
  type ExcelDateSystem,
} from '@react-sheets/formula-engine';
export { applyRowPermutation, createRowPermutationPlan, validatePermutationMetadata, type RowPermutationPlan } from './data-transform';
export { columnLabel, parseColumnLabel, cellAddress, parseAddress, a1Range } from './address';
export {
  loadWorkbookFromSnapshot,
  createWorkbookSnapshot,
  migrateStoredWorkbookSnapshot,
  assertCanonicalWorkbookSnapshot,
  MAX_DRAWING_SOURCE_CELLS,
  type WorkbookSnapshot,
  type WorkbookDimensionMetrics,
} from './snapshot';
export {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
  type QueryLoadTargetSnapshot,
  type QueryStepSnapshot,
} from './workbook-state';

import { canonicalizePivotDefinition, type PivotModel } from './pivot';
export * from './pivot';
import type { GanttSheetDefinition, ReportSheetDefinition, TableSheetDefinition, WorkbookDataModel, WorkbookTableModel } from './data-model';
import { normalizeDataSourceManifest, type DataSourceManifest, type SheetDataRegion } from './data-source';
export * from './data-model';
export * from './data-source';

export type SheetKind = 'worksheet' | 'table-sheet' | 'gantt-sheet' | 'report-sheet';

export interface SparklineModel {
  id: string;
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
  sourceRange: RangeRef;
  type: 'line' | 'column' | 'win-loss';
  color: string;
  negativeColor?: string;
  highlightMax?: boolean;
  highlightMin?: boolean;
  highlightFirst?: boolean;
  highlightLast?: boolean;
  highlightNegative?: boolean;
  groupId?: string;
  showAxis?: boolean;
  showMarkers?: boolean;
}

/** 隔行色带规则 */
export interface BandedRule {
  range: RangeRef;
  firstColor: string;
  secondColor: string;
}

export type ConditionalFormatType = 'highlight' | 'dataBar' | 'colorScale' | 'iconSet' | 'topBottom';
export type ConditionalFormatOperator =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equal'
  | 'notEqual'
  | 'containsText'
  | 'notContainsText'
  | 'duplicate'
  | 'unique'
  | 'formula'
  | 'top'
  | 'bottom';

export interface ConditionalFormatTopBottom {
  direction: 'top' | 'bottom';
  /** Number of values, or a percentage when `percent` is true. */
  rank: number;
  percent?: boolean;
}

export interface ConditionalFormatRule {
  id: string;
  sheetId: SheetId;
  ranges: RangeRef[];
  type: ConditionalFormatType;
  /** Lower values are evaluated first. Excel defaults to the insertion order. */
  priority?: number;
  /** Stop evaluating lower-priority rules after this rule matches a cell. */
  stopIfTrue?: boolean;
  operator?: ConditionalFormatOperator;
  value1?: string | number;
  value2?: string | number;
  style?: CellStyle;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  barColor?: string;
  iconSet?: string;
  iconThresholds?: Array<{ type: 'percent' | 'percentile' | 'num' | 'formula'; value?: number }>;
  topBottom?: ConditionalFormatTopBottom;
}

export type DataValidationType = 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'checkbox' | 'textLength' | 'custom';
export type DataValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan';

export interface DataValidationRule {
  id: string;
  sheetId: SheetId;
  ranges: RangeRef[];
  type: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  /** Excel error alert style. Only STOP blocks a write. */
  alertStyle?: 'stop' | 'warning' | 'information';
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showDropdown?: boolean;
  /** Allows comma-separated values for list validation when enabled. */
  multiSelect?: boolean;
  listSource?:
    | { kind: 'values'; values: string[] }
    | { kind: 'range'; range: RangeRef }
    | { kind: 'formula'; formula: string };
  promptTitle?: string;
  promptMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}

/** A range-free validation shape stored by a workbook template. */
export type CellValidationTemplate = Omit<DataValidationRule, 'id' | 'sheetId' | 'ranges'>;

/** Persisted workbook-native template, reusable across sheets and collaboration revisions. */
export interface CellStyleTemplate {
  id: string;
  name: string;
  style: CellStyle;
  dataValidation?: CellValidationTemplate;
  editor?: CellEditorConfig;
}

export type FilterScalar = string | number | boolean | null;

/**
 * The only value projection accepted by filter evaluation.  `CellData.value`
 * is authored storage; formulas, spills, and data blocks may expose a
 * different current result.  Callers resolve that result before constructing
 * this carrier, while the source cell remains available for style/icon/date
 * metadata.
 */
export interface FilterCellValue {
  cell?: CellData;
  value: FilterScalar;
  text: string;
  dateSystem?: '1900' | '1904';
  errorCode?: FormulaErrorCode;
}

export function resolveFilterCellValue(cell?: CellData, evaluated?: unknown, dateSystem?: '1900' | '1904'): FilterCellValue {
  const candidate = evaluated === undefined
    ? cell?.formula !== undefined
      ? cell?.formulaValue ?? null
      : cell?.formulaValue ?? cell?.value ?? null
    : evaluated;
  if (isFormulaError(candidate)) {
    return { cell, value: null, text: '', ...(dateSystem ? { dateSystem } : {}), errorCode: candidate.code };
  }
  if (candidate === null || typeof candidate === 'string' || typeof candidate === 'number' || typeof candidate === 'boolean') {
    return { cell, value: candidate, text: candidate == null ? '' : String(candidate), ...(dateSystem ? { dateSystem } : {}) };
  }
  return { cell, value: null, text: '', ...(dateSystem ? { dateSystem } : {}) };
}

export interface DateGroupItem {
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
}

export type FilterComparisonOperator =
  | 'equals'
  | 'notEquals'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'contains'
  | 'notContains'
  | 'beginsWith'
  | 'endsWith';

export interface FilterComparison {
  operator: FilterComparisonOperator;
  value: FilterScalar;
}

export type DynamicFilterType =
  | 'today' | 'yesterday' | 'tomorrow'
  | 'thisWeek' | 'lastWeek' | 'nextWeek'
  | 'thisMonth' | 'lastMonth' | 'nextMonth'
  | 'thisQuarter' | 'lastQuarter' | 'nextQuarter'
  | 'thisYear' | 'lastYear' | 'nextYear' | 'yearToDate';

const DYNAMIC_FILTER_TYPES: ReadonlySet<string> = new Set([
  'today', 'yesterday', 'tomorrow',
  'thisWeek', 'lastWeek', 'nextWeek',
  'thisMonth', 'lastMonth', 'nextMonth',
  'thisQuarter', 'lastQuarter', 'nextQuarter',
  'thisYear', 'lastYear', 'nextYear', 'yearToDate',
]);

export function isDynamicFilterType(value: unknown): value is DynamicFilterType {
  return typeof value === 'string' && DYNAMIC_FILTER_TYPES.has(value);
}

export type FilterCriterion =
  | { kind: 'values'; values: FilterScalar[]; includeBlank: boolean; dateGroups?: DateGroupItem[] }
  | { kind: 'custom'; join: 'and' | 'or'; conditions: [FilterComparison, FilterComparison?] }
  | { kind: 'dynamic'; type: DynamicFilterType; value?: number; maxValue?: number }
  | { kind: 'top10'; top: boolean; percent: boolean; rank: number; filterValue?: number }
  | { kind: 'color'; target: 'cell' | 'font'; dxfId: number; style?: Partial<CellStyle> }
  | { kind: 'icon'; iconSet: string; iconId: number };

export interface SortStateModel {
  ref: RangeRef;
  conditions: Array<{ ref: RangeRef; descending: boolean; customList?: string[] }>;
}

export interface AutoFilterColumn {
  column: Column;
  criterion?: FilterCriterion;
  showButton: boolean;
  hiddenButton: boolean;
  preservedXml?: unknown;
}

export interface AutoFilterModel {
  sheetId: SheetId;
  range: RangeRef;
  columns: Record<Column, AutoFilterColumn>;
  sortState?: SortStateModel;
  preservedXml?: unknown;
}

export interface SortCriterion {
  column: Column;
  ascending: boolean;
}

export class CellMatrix {
  private readonly rows = new Map<Row, Map<Column, CellData>>();
  private revisionCounter = 0;

  /** Monotonic content revision used by derived caches; it is not persisted. */
  get revision(): number {
    return this.revisionCounter;
  }

  get(row: Row, column: Column): CellData | undefined {
    return this.rows.get(row)?.get(column);
  }

  set(row: Row, column: Column, cell: CellData): void {
    let rowMap = this.rows.get(row);
    if (!rowMap) {
      rowMap = new Map<Column, CellData>();
      this.rows.set(row, rowMap);
    }
    const fontFamily = cell.style?.fontFamily;
    const normalizedCell = fontFamily === undefined
      ? cell
      : { ...cell, style: { ...cell.style, fontFamily: normalizeFontFamily(fontFamily) } };
    rowMap.set(column, normalizedCell);
    this.revisionCounter += 1;
  }

  delete(row: Row, column: Column): void {
    const rowMap = this.rows.get(row);
    const existed = rowMap?.has(column) ?? false;
    rowMap?.delete(column);
    if (rowMap?.size === 0) this.rows.delete(row);
    if (existed) this.revisionCounter += 1;
  }

  has(row: Row, column: Column): boolean {
    return this.rows.get(row)?.has(column) ?? false;
  }

  clear(): void {
    if (this.rows.size > 0) this.revisionCounter += 1;
    this.rows.clear();
  }

  count(): number {
    let count = 0;
    for (const columns of this.rows.values()) {
      count += columns.size;
    }
    return count;
  }

  forEach(callback: (cell: CellData, row: Row, column: Column) => void): void {
    for (const [row, columns] of this.rows) {
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  clone(): CellMatrix {
    const copy = new CellMatrix();
    this.forEach((cell, row, column) => copy.set(row, column, { ...cell }));
    return copy;
  }

  toJSON(): Record<string, Record<string, CellData>> {
    const result: Record<string, Record<string, CellData>> = {};
    this.forEach((cell, row, column) => {
      result[row] ??= {};
      result[row][column] = { ...cell };
    });
    return result;
  }

  static fromJSON(input: Record<string, Record<string, CellData>> | undefined): CellMatrix {
    const matrix = new CellMatrix();
    for (const [row, columns] of Object.entries(input ?? {})) {
      for (const [column, cell] of Object.entries(columns)) {
        matrix.set(Number(row), Number(column), { ...cell });
      }
    }
    return matrix;
  }

  /** 沿行轴整体平移:dir=+1 下移(插入),dir=-1 上移(删除);越界丢弃 */
  shiftRows(at: Row, count: number, direction: 1 | -1): void {
    const entries: Array<[Row, Column, CellData]> = [];
    const delta = direction * count;
    this.forEach((cell, row, column) => {
      if (row >= at) entries.push([row, column, cell]);
    });
    if (direction === -1) {
      // 从小到大删除,避免覆盖
      entries.sort((a, b) => a[0] - b[0]);
    } else {
      entries.sort((a, b) => b[0] - a[0]);
    }
    for (const [row, column] of entries) this.delete(row, column);
    for (const [row, column, cell] of entries) {
      this.set(row + delta, column, cell);
    }
  }

  /** 沿列轴整体平移:dir=+1 右移(插入),dir=-1 左移(删除) */
  shiftColumns(at: Column, count: number, direction: 1 | -1): void {
    const entries: Array<[Row, Column, CellData]> = [];
    const delta = direction * count;
    this.forEach((cell, row, column) => {
      if (column >= at) entries.push([row, column, cell]);
    });
    if (direction === -1) entries.sort((a, b) => a[1] - b[1]);
    else entries.sort((a, b) => b[1] - a[1]);
    for (const [row, column] of entries) this.delete(row, column);
    for (const [row, column, cell] of entries) {
      this.set(row, column + delta, cell);
    }
  }

  /** 摘除区间内全部单元格并返回(用于删除行的逆操作恢复) */
  extractRegion(startRow: Row, endRow: Row, startColumn: Column, endColumn: Column): Array<{ row: Row; column: Column; cell: CellData }> {
    const extracted: Array<{ row: Row; column: Column; cell: CellData }> = [];
    this.forEach((cell, row, column) => {
      if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
        extracted.push({ row, column, cell: structuredClone(cell) });
      }
    });
    for (const item of extracted) this.delete(item.row, item.column);
    return extracted;
  }

  placeRegion(items: ReadonlyArray<{ row: Row; column: Column; cell: CellData }>): void {
    for (const item of items) this.set(item.row, item.column, structuredClone(item.cell));
  }
}

export class WorksheetModel {
  kind: SheetKind = 'worksheet';
  tableSheet?: TableSheetDefinition;
  ganttSheet?: GanttSheetDefinition;
  reportSheet?: ReportSheetDefinition;
  readonly cells = new CellMatrix();
  /** Block-backed regions are metadata only; their bytes never enter CellMatrix. */
  readonly dataRegions: SheetDataRegion[] = [];
  readonly merges: MergeSpan[] = [];
  readonly pivots: PivotModel[] = [];
  readonly sparklines: SparklineModel[] = [];
  readonly conditionalFormats: ConditionalFormatRule[] = [];
  readonly dataValidations: DataValidationRule[] = [];
  readonly sheetTables: SheetTableModel[] = [];
  readonly drawings: DrawingObject[] = [];
  readonly drawingPayloads = new Map<string, DrawingPayload>();
  /** Canonical persisted hyperlink metadata keyed by row:column. */
  readonly hyperlinks = new Map<string, CellHyperlink>();
  readonly notes = new Map<string, CellNote>();
  readonly commentThreads: CommentThread[] = [];
  readonly spillRanges: SpillRange[] = [];
  readonly protectionRules: ProtectionRule[] = [];
  readonly sparklineGroups: SparklineGroup[] = [];
  outline?: OutlineModel;
  showGridlines = true;
  showHeaders = true;
  zoom = 100;
  hidden = false;
  autoFilter?: AutoFilterModel;
  bandedRule?: BandedRule;
  defaultRowHeightPx = 20;
  defaultColumnWidthPx = 64;
  readonly rowHeightsPx: Record<number, number> = {};
  readonly columnWidthsPx: Record<number, number> = {};
  readonly hiddenRows = new Set<number>();
  readonly hiddenColumns = new Set<number>();
  tabColor?: string;
  pane: WorksheetPane = { kind: 'none' };

  /** 深拷贝当前工作表(删除工作表撤销恢复用) */
  cloneSheet(): WorksheetModel {
    return this.cloneWithIdentity(this.id, this.name);
  }

  cloneWithIdentity(id: SheetId, name: string): WorksheetModel {
    const copy = new WorksheetModel(id, name, this.rowCount, this.columnCount);
    copy.kind = this.kind;
    copy.tableSheet = this.tableSheet ? structuredClone(this.tableSheet) : undefined;
    copy.ganttSheet = this.ganttSheet ? structuredClone(this.ganttSheet) : undefined;
    copy.reportSheet = this.reportSheet ? structuredClone(this.reportSheet) : undefined;
    this.cells.forEach((cell, row, column) => copy.cells.set(row, column, structuredClone(cell)));
    copy.dataRegions.push(...structuredClone(this.dataRegions));
    copy.merges.push(...structuredClone(this.merges));
    copy.pivots.push(...structuredClone(this.pivots));
    copy.sparklines.push(...structuredClone(this.sparklines));
    copy.conditionalFormats.push(...structuredClone(this.conditionalFormats));
    copy.dataValidations.push(...structuredClone(this.dataValidations));
    copy.autoFilter = this.autoFilter ? structuredClone(this.autoFilter) : undefined;
    copy.bandedRule = this.bandedRule ? structuredClone(this.bandedRule) : undefined;
    copy.defaultRowHeightPx = this.defaultRowHeightPx;
    copy.defaultColumnWidthPx = this.defaultColumnWidthPx;
    Object.assign(copy.rowHeightsPx, this.rowHeightsPx);
    Object.assign(copy.columnWidthsPx, this.columnWidthsPx);
    for (const row of this.hiddenRows) copy.hiddenRows.add(row);
    for (const column of this.hiddenColumns) copy.hiddenColumns.add(column);
    copy.sheetTables.push(...structuredClone(this.sheetTables));
    copy.drawings.push(...structuredClone(this.drawings));
    for (const [key, payload] of this.drawingPayloads) copy.drawingPayloads.set(key, structuredClone(payload));
    for (const [key, hyperlink] of this.hyperlinks) copy.hyperlinks.set(key, structuredClone(hyperlink));
    for (const [key, note] of this.notes) copy.notes.set(key, structuredClone(note));
    copy.commentThreads.push(...structuredClone(this.commentThreads));
    copy.spillRanges.push(...structuredClone(this.spillRanges));
    copy.protectionRules.push(...structuredClone(this.protectionRules));
    copy.sparklineGroups.push(...structuredClone(this.sparklineGroups));
    copy.outline = this.outline ? structuredClone(this.outline) : undefined;
    copy.showGridlines = this.showGridlines;
    copy.showHeaders = this.showHeaders;
    copy.zoom = this.zoom;
    copy.hidden = this.hidden;
    copy.tabColor = this.tabColor;
    copy.pane = normalizeWorksheetPane(this.pane);
    return copy;
  }

  constructor(
    readonly id: SheetId,
    public name: string,
    public rowCount = 1000,
    public columnCount = 26,
  ) {}

  isMerged(row: Row, column: Column): MergeSpan | undefined {
    return this.merges.find(
      (m) =>
        row >= m.range.startRow &&
        row <= m.range.endRow &&
        column >= m.range.startColumn &&
        column <= m.range.endColumn,
    );
  }

  isMergeAnchor(row: Row, column: Column): boolean {
    const merge = this.isMerged(row, column);
    return !merge || (merge.anchor.row === row && merge.anchor.column === column);
  }
}

export function noteCellKey(row: Row, column: Column): string {
  return `${row}:${column}`;
}

export function getDrawingPayload(sheet: WorksheetModel, payloadId: string): DrawingPayload | undefined {
  return sheet.drawingPayloads.get(payloadId);
}

export function getCellNote(sheet: WorksheetModel, row: Row, column: Column): CellNote | undefined {
  return sheet.notes.get(noteCellKey(row, column));
}

export interface SheetSnapshot {
  kind: SheetKind;
  id: SheetId;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, Record<string, CellData>>;
  dataRegions?: SheetDataRegion[];
  merges: MergeSpan[];
  pane: WorksheetPane;
  pivots: PivotModel[];
  sparklines: SparklineModel[];
  sparklineGroups?: SparklineGroup[];
  /** Canonical floating-object collection. Legacy per-kind collections are not part of snapshots. */
  drawings: DrawingObject[];
  drawingPayloads: Record<string, DrawingPayload>;
  hyperlinks?: Array<{ row: number; column: number; hyperlink: CellHyperlink }>;
  notes?: Array<{ row: number; column: number; note: CellNote }>;
  commentThreads?: CommentThread[];
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
  defaultRowHeightPx: number;
  defaultColumnWidthPx: number;
  rowHeightsPx?: Record<number, number>;
  columnWidthsPx?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  tabColor?: string;
  bandedRule?: BandedRule;
  autoFilter?: AutoFilterModel;
  sheetTables?: SheetTableModel[];
  spillRanges?: SpillRange[];
  protectionRules?: ProtectionRule[];
  showGridlines?: boolean;
  showHeaders?: boolean;
  zoom?: number;
  hidden?: boolean;
  outline?: OutlineModel;
  tableSheet?: TableSheetDefinition;
  ganttSheet?: GanttSheetDefinition;
  reportSheet?: ReportSheetDefinition;
}

export class WorkbookModel {
  readonly sheets = new Map<SheetId, WorksheetModel>();
  /** Sole canonical structured-data owner; bytes referenced by sources remain in the block store. */
  readonly dataModel = {
    sources: new Map<string, DataSourceManifest>(),
    tables: new Map<string, WorkbookTableModel>(),
    relationships: new Map<string, import('./data-model').DataRelationship>(),
    views: new Map<string, import('./data-model').DataViewDefinition>(),
  };
  /** Canonical workbook-owned print state; no host-side cache is authoritative. */
  readonly printDocuments = new Map<SheetId, PrintDocumentSnapshot>();
  /** Persistence-safe query definitions; connector credentials are redacted. */
  readonly queryDefinitions = new Map<string, QueryDefinitionSnapshot>();
  /** Canonical workbook-owned style/template library. */
  readonly cellStyleTemplates = new Map<string, CellStyleTemplate>();
  /** 工作表 Tab 顺序 */
  sheetOrder: SheetId[] = [];
  /** The sole canonical defined-name store. Formula consumers receive a derived workbook-scope view. */
  readonly definedNameModels: DefinedNameModel[] = [];
  dimensionMetrics: WorkbookDimensionMetrics = { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 };

  /**
   * Formula engines still accept a workbook-scope string map. This is a
   * read-only projection of `definedNameModels`, never an independently
   * mutable source of truth. Sheet-scoped names are resolved through
   * `getDefinedName(name, sheetId)` by callers that have a sheet context.
   */
  get definedNames(): Readonly<Record<string, string>> {
    const result: Record<string, string> = {};
    for (const entry of this.definedNameModels) {
      if (entry.scope === 'workbook') result[entry.name] = entry.formula;
    }
    return result;
  }

  constructor(readonly unitId: UnitId, public name: string) {
    const sheet = new WorksheetModel('sheet-1', 'Sheet1');
    this.sheets.set(sheet.id, sheet);
    this.sheetOrder = [sheet.id];
  }

  listCellStyleTemplates(): CellStyleTemplate[] {
    return [...this.cellStyleTemplates.values()].map((template) => structuredClone(template));
  }

  setCellStyleTemplate(template: CellStyleTemplate): void {
    const id = template.id.trim();
    const name = template.name.trim();
    if (!id) throw new Error('Cell style template id is required');
    if (!name) throw new Error('Cell style template name is required');
    if (template.style.indent !== undefined && (!Number.isInteger(template.style.indent) || template.style.indent < 0 || template.style.indent > 250)) {
      throw new Error('Cell style template indent is invalid');
    }
    if (template.editor && !['text', 'number', 'date', 'list', 'checkbox'].includes(template.editor.kind)) {
      throw new Error('Cell style template editor is invalid');
    }
    if (template.editor?.kind === 'list' && (!Array.isArray(template.editor.values) || template.editor.values.some((value) => !value.trim()))) {
      throw new Error('Cell style template list editor values are invalid');
    }
    this.cellStyleTemplates.set(id, structuredClone({ ...template, id, name }));
  }

  removeCellStyleTemplate(templateId: string): CellStyleTemplate | undefined {
    const previous = this.cellStyleTemplates.get(templateId);
    this.cellStyleTemplates.delete(templateId);
    return previous ? structuredClone(previous) : undefined;
  }

  /** Stable workbook default. UI selection belongs exclusively to WorkbookSession. */
  get primarySheetId(): SheetId {
    const sheetId = this.sheetOrder[0];
    if (!sheetId) throw new Error('A workbook must contain at least one worksheet');
    return sheetId;
  }

  getSheet(sheetId: SheetId): WorksheetModel {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetId}`);
    return sheet;
  }

  getSheetByName(name: string): WorksheetModel | undefined {
    for (const sheet of this.sheets.values()) {
      if (sheet.name.toLowerCase() === name.toLowerCase()) return sheet;
    }
    return undefined;
  }

  getSheets(): WorksheetModel[] {
    return this.sheetOrder
      .map((id) => this.sheets.get(id))
      .filter((sheet): sheet is WorksheetModel => sheet !== undefined);
  }

  getVisibleSheets(): WorksheetModel[] {
    return this.getSheets().filter((sheet) => !sheet.hidden);
  }

  getTable(tableId: string): WorkbookTableModel {
    const table = this.dataModel.tables.get(tableId);
    if (!table) throw new Error(`Unknown table: ${tableId}`);
    return table;
  }

  getDataSource(dataSourceId: string): DataSourceManifest {
    const source = this.dataModel.sources.get(dataSourceId);
    if (!source) throw new Error(`Unknown data source: ${dataSourceId}`);
    return structuredClone(source);
  }

  getDataModel(): WorkbookDataModel {
    return {
      sources: [...this.dataModel.sources.values()].map((source) => structuredClone(source)),
      tables: [...this.dataModel.tables.values()].map((table) => structuredClone(table)),
      relationships: [...this.dataModel.relationships.values()].map((relationship) => structuredClone(relationship)),
      views: [...this.dataModel.views.values()].map((view) => structuredClone(view)),
    };
  }

  getPrintDocument(sheetId: SheetId): PrintDocumentSnapshot | undefined {
    this.getSheet(sheetId);
    const document = this.printDocuments.get(sheetId);
    return document ? structuredClone(document) : undefined;
  }

  setPrintDocument(document: PrintDocumentSnapshot): void {
    if (document.unitId !== this.unitId) throw new Error(`Print document unit mismatch: expected ${this.unitId}, received ${document.unitId}`);
    this.getSheet(document.sheetId);
    this.printDocuments.set(document.sheetId, normalizePrintDocumentSnapshot(document));
  }

  removePrintDocument(sheetId: SheetId): PrintDocumentSnapshot | undefined {
    this.getSheet(sheetId);
    const document = this.printDocuments.get(sheetId);
    this.printDocuments.delete(sheetId);
    return document ? structuredClone(document) : undefined;
  }

  clearPrintDocuments(): void {
    this.printDocuments.clear();
  }

  listPrintDocuments(): PrintDocumentSnapshot[] {
    return [...this.printDocuments.values()].map((document) => structuredClone(document));
  }

  getQueryDefinition(queryId: string): QueryDefinitionSnapshot | undefined {
    const definition = this.queryDefinitions.get(queryId);
    return definition ? structuredClone(definition) : undefined;
  }

  setQueryDefinition(definition: QueryDefinitionSnapshot): void {
    const normalized = normalizeQueryDefinitionSnapshot(definition);
    this.queryDefinitions.set(normalized.id, normalized);
  }

  removeQueryDefinition(queryId: string): QueryDefinitionSnapshot | undefined {
    const definition = this.queryDefinitions.get(queryId);
    this.queryDefinitions.delete(queryId);
    return definition ? structuredClone(definition) : undefined;
  }

  clearQueryDefinitions(): void {
    this.queryDefinitions.clear();
  }

  listQueryDefinitions(): QueryDefinitionSnapshot[] {
    return [...this.queryDefinitions.values()].map((definition) => structuredClone(definition));
  }

  getDefinedName(name: string, sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    if (sheetId) {
      const local = this.definedNameModels.find((entry) => entry.scope === 'sheet'
        && entry.sheetId === sheetId
        && entry.name.toLocaleLowerCase() === normalized);
      if (local) return structuredClone(local);
    }
    const global = this.definedNameModels.find((entry) => entry.scope === 'workbook'
      && entry.name.toLocaleLowerCase() === normalized);
    return global ? structuredClone(global) : undefined;
  }

  getDefinedNameExact(name: string, scope: DefinedNameScope, sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    const exact = this.definedNameModels.find((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toLocaleLowerCase() === normalized);
    return exact ? structuredClone(exact) : undefined;
  }

  listDefinedNames(sheetId?: SheetId): DefinedNameModel[] {
    return this.definedNameModels
      .filter((entry) => entry.scope === 'workbook' || entry.sheetId === sheetId)
      .map((entry) => structuredClone(entry));
  }

  setDefinedName(input: DefinedNameModel): DefinedNameModel {
    const model = normalizeDefinedNameModel(input);
    const index = this.definedNameModels.findIndex((entry) => entry.scope === model.scope
      && entry.sheetId === model.sheetId
      && entry.name.toLocaleLowerCase() === model.name.toLocaleLowerCase());
    if (index >= 0) this.definedNameModels[index] = structuredClone(model);
    else this.definedNameModels.push(structuredClone(model));
    return structuredClone(model);
  }

  removeDefinedName(name: string, scope: DefinedNameScope = 'workbook', sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    const index = this.definedNameModels.findIndex((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toLocaleLowerCase() === normalized);
    const previous = index >= 0 ? this.definedNameModels[index] : undefined;
    if (index >= 0) this.definedNameModels.splice(index, 1);
    return previous ? structuredClone(previous) : undefined;
  }

  addTable(table: WorkbookTableModel): void {
    if (this.dataModel.tables.has(table.id)) throw new Error(`Table already exists: ${table.id}`);
    this.dataModel.tables.set(table.id, structuredClone(table));
  }

  addDataSource(source: DataSourceManifest): void {
    const normalized = normalizeDataSourceManifest(source);
    if (this.dataModel.sources.has(normalized.id)) throw new Error(`Data source already exists: ${normalized.id}`);
    this.dataModel.sources.set(normalized.id, structuredClone(normalized));
  }

  updateDataSource(source: DataSourceManifest): void {
    const normalized = normalizeDataSourceManifest(source);
    if (!this.dataModel.sources.has(normalized.id)) throw new Error(`Unknown data source: ${normalized.id}`);
    this.dataModel.sources.set(normalized.id, structuredClone(normalized));
  }

  removeDataSource(dataSourceId: string): DataSourceManifest {
    const source = this.getDataSource(dataSourceId);
    if (this.getSheets().some((sheet) => sheet.dataRegions.some((region) => region.sourceId === dataSourceId))) {
      throw new Error(`Data source is still referenced by a sheet region: ${dataSourceId}`);
    }
    this.dataModel.sources.delete(dataSourceId);
    return source;
  }

  removeTable(tableId: string): WorkbookTableModel {
    const table = this.getTable(tableId);
    this.dataModel.tables.delete(tableId);
    return table;
  }

  addSheet(id: SheetId, name: string, rowCount = 1000, columnCount = 26): WorksheetModel {
    if (this.sheets.has(id)) throw new Error(`Sheet already exists: ${id}`);
    const sheet = new WorksheetModel(id, name, rowCount, columnCount);
    this.sheets.set(id, sheet);
    this.sheetOrder.push(id);
    return sheet;
  }

  addAdvancedSheet(input: {
    id: SheetId;
    name: string;
    kind: Exclude<SheetKind, 'worksheet'>;
    rowCount?: number;
    columnCount?: number;
    tableSheet?: TableSheetDefinition;
    ganttSheet?: GanttSheetDefinition;
    reportSheet?: ReportSheetDefinition;
  }): WorksheetModel {
    const sheet = this.addSheet(input.id, input.name, input.rowCount, input.columnCount);
    sheet.kind = input.kind;
    sheet.tableSheet = input.tableSheet ? structuredClone(input.tableSheet) : undefined;
    sheet.ganttSheet = input.ganttSheet ? structuredClone(input.ganttSheet) : undefined;
    sheet.reportSheet = input.reportSheet ? structuredClone(input.reportSheet) : undefined;
    return sheet;
  }

  duplicateSheet(sourceSheetId: SheetId, newId: SheetId, newName: string): WorksheetModel {
    const source = this.getSheet(sourceSheetId);
    const copy = source.cloneWithIdentity(newId, newName);
    this.sheets.set(newId, copy);
    const scopedNames = this.definedNameModels
      .filter((entry) => entry.scope === 'sheet' && entry.sheetId === sourceSheetId)
      .map((entry) => ({ ...entry, sheetId: newId }));
    this.definedNameModels.push(...structuredClone(scopedNames));
    const printDocument = this.printDocuments.get(sourceSheetId);
    if (printDocument) {
      this.printDocuments.set(newId, structuredClone({
        ...printDocument,
        sheetId: newId,
        printAreas: printDocument.printAreas.map((area) => ({ sheetId: newId, range: { ...area.range, sheetId: newId } })),
        pageBreaks: printDocument.pageBreaks.map((pageBreak) => pageBreak.row !== undefined
          ? { sheetId: newId, row: pageBreak.row }
          : { sheetId: newId, column: pageBreak.column }),
      }));
    }
    const sourceIndex = this.sheetOrder.indexOf(sourceSheetId);
    this.sheetOrder.splice(sourceIndex + 1, 0, newId);
    return copy;
  }

  reorderSheet(sheetId: SheetId, toIndex: number): void {
    const fromIndex = this.sheetOrder.indexOf(sheetId);
    if (fromIndex < 0) throw new Error(`Unknown sheet: ${sheetId}`);
    const clamped = Math.max(0, Math.min(toIndex, this.sheetOrder.length - 1));
    this.sheetOrder.splice(fromIndex, 1);
    this.sheetOrder.splice(clamped, 0, sheetId);
  }

  removeSheet(sheetId: SheetId): WorksheetModel {
    if (this.sheets.size <= 1) throw new Error('A workbook must keep at least one worksheet');
    const sheet = this.getSheet(sheetId);
    this.sheets.delete(sheetId);
    this.printDocuments.delete(sheetId);
    for (let index = this.definedNameModels.length - 1; index >= 0; index -= 1) {
      if (this.definedNameModels[index]?.scope === 'sheet' && this.definedNameModels[index]?.sheetId === sheetId) {
        this.definedNameModels.splice(index, 1);
      }
    }
    this.sheetOrder = this.sheetOrder.filter((id) => id !== sheetId);
    return sheet;
  }

  getSheetSnapshot(sheetId: SheetId): SheetSnapshot {
    const sheet = this.snapshot().sheets.find((entry) => entry.id === sheetId);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetId}`);
    return structuredClone(sheet);
  }

  restoreSheetSnapshot(snapshot: SheetSnapshot, index = this.sheetOrder.length): void {
    if (this.sheets.has(snapshot.id)) throw new Error(`Sheet already exists: ${snapshot.id}`);
    const current = this.snapshot();
    const hydrated = WorkbookModel.fromSnapshot({ ...current, sheets: [structuredClone(snapshot)] });
    const sheet = hydrated.getSheet(snapshot.id);
    this.sheets.set(sheet.id, sheet);
    const bounded = Math.max(0, Math.min(index, this.sheetOrder.length));
    this.sheetOrder.splice(bounded, 0, sheet.id);
  }

  snapshot(): WorkbookSnapshot {
    return {
      schema: 'WorkbookSnapshot',
      version: 5,
      unitId: this.unitId,
      name: this.name,
      dimensionMetrics: structuredClone(this.dimensionMetrics),
      // Keep the legacy formula-map field as a derived wire projection for
      // import/export consumers. It is never hydrated as mutable state.
      definedNames: { ...this.definedNames },
      definedNameModels: structuredClone(this.definedNameModels),
      dataModel: this.getDataModel(),
      printDocuments: this.listPrintDocuments(),
      queryDefinitions: this.listQueryDefinitions(),
      cellStyleTemplates: this.listCellStyleTemplates(),
      sheets: this.getSheets().map((sheet) => ({
        kind: sheet.kind,
        id: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        cells: sheet.cells.toJSON(),
        dataRegions: structuredClone(sheet.dataRegions),
        merges: structuredClone(sheet.merges),
        pane: normalizeWorksheetPane(sheet.pane),
        pivots: structuredClone(sheet.pivots),
        sparklines: structuredClone(sheet.sparklines),
        conditionalFormats: structuredClone(sheet.conditionalFormats),
        dataValidations: structuredClone(sheet.dataValidations),
        defaultRowHeightPx: sheet.defaultRowHeightPx,
        defaultColumnWidthPx: sheet.defaultColumnWidthPx,
        rowHeightsPx: { ...sheet.rowHeightsPx },
        columnWidthsPx: { ...sheet.columnWidthsPx },
        hiddenRows: [...sheet.hiddenRows],
        hiddenColumns: [...sheet.hiddenColumns],
        tabColor: sheet.tabColor,
        bandedRule: sheet.bandedRule ? structuredClone(sheet.bandedRule) : undefined,
        autoFilter: sheet.autoFilter ? structuredClone(sheet.autoFilter) : undefined,
        sheetTables: structuredClone(sheet.sheetTables),
        sparklineGroups: structuredClone(sheet.sparklineGroups),
        drawings: structuredClone(sheet.drawings),
        drawingPayloads: Object.fromEntries([...sheet.drawingPayloads.entries()].map(([k, v]) => [k, structuredClone(v)])),
        hyperlinks: [...sheet.hyperlinks.entries()].map(([key, hyperlink]) => {
          const [row, column] = key.split(':').map(Number);
          return { row: row!, column: column!, hyperlink: structuredClone(hyperlink) };
        }),
        notes: [...sheet.notes.entries()].map(([key, note]) => {
          const [row, column] = key.split(':').map(Number);
          return { row: row!, column: column!, note: structuredClone(note) };
        }),
        commentThreads: structuredClone(sheet.commentThreads),
        spillRanges: structuredClone(sheet.spillRanges),
        protectionRules: structuredClone(sheet.protectionRules),
        showGridlines: sheet.showGridlines,
        showHeaders: sheet.showHeaders,
        zoom: sheet.zoom,
        hidden: sheet.hidden,
        outline: sheet.outline ? structuredClone(sheet.outline) : undefined,
        tableSheet: sheet.tableSheet ? structuredClone(sheet.tableSheet) : undefined,
        ganttSheet: sheet.ganttSheet ? structuredClone(sheet.ganttSheet) : undefined,
        reportSheet: sheet.reportSheet ? structuredClone(sheet.reportSheet) : undefined,
      })),
    };
  }

  static fromSnapshot(snapshot: WorkbookSnapshot): WorkbookModel {
    if (snapshot.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
    if (snapshot.version !== 5) throw new Error('Unsupported workbook snapshot version');
    if (snapshot.sheets.length === 0) throw new Error('Workbook snapshot must contain at least one sheet');
    const workbook = new WorkbookModel(snapshot.unitId, snapshot.name);
    workbook.dimensionMetrics = structuredClone(snapshot.dimensionMetrics);
    workbook.sheets.clear();
    // `definedNameModels` is canonical. The optional map is accepted only as
    // a boundary projection for older snapshots and is immediately folded
    // into the canonical scoped collection.
    const definedNameModels = snapshot.definedNameModels
      ?? Object.entries(snapshot.definedNames ?? {}).map(([name, formula]) => ({ name, formula, scope: 'workbook' as const }));
    for (const entry of definedNameModels) workbook.setDefinedName(entry);
    for (const table of snapshot.dataModel.tables) workbook.dataModel.tables.set(table.id, structuredClone(table));
    for (const source of snapshot.dataModel.sources) workbook.addDataSource(source);
    for (const relationship of snapshot.dataModel.relationships) workbook.dataModel.relationships.set(relationship.id, structuredClone(relationship));
    for (const view of snapshot.dataModel.views) workbook.dataModel.views.set(view.id, structuredClone(view));
    for (const input of snapshot.sheets) {
      const sheet = new WorksheetModel(input.id, input.name, input.rowCount, input.columnCount);
      sheet.kind = input.kind;
      sheet.tableSheet = input.tableSheet ? structuredClone(input.tableSheet) : undefined;
      sheet.ganttSheet = input.ganttSheet ? structuredClone(input.ganttSheet) : undefined;
      sheet.reportSheet = input.reportSheet ? structuredClone(input.reportSheet) : undefined;
      const matrix = CellMatrix.fromJSON(input.cells);
      matrix.forEach((cell, row, column) => {
        const normalized = structuredClone(cell);
        const legacy = normalized.hyperlinkDetail
          ?? (normalized.hyperlink ? {
            id: `legacy-hyperlink-${row}-${column}`,
            target: { kind: 'url' as const, url: normalized.hyperlink },
          } : undefined);
        delete normalized.hyperlink;
        delete normalized.hyperlinkDetail;
        sheet.cells.set(row, column, normalized);
        if (legacy) sheet.hyperlinks.set(noteCellKey(row, column), legacy);
      });
      if (input.dataRegions) sheet.dataRegions.push(...structuredClone(input.dataRegions));
      sheet.merges.push(...structuredClone(input.merges));
      sheet.pane = normalizeWorksheetPane(input.pane);
      sheet.pivots.push(...input.pivots.map((pivot) => canonicalizePivotDefinition(structuredClone(pivot))));
      sheet.sparklines.push(...structuredClone(input.sparklines));
      if (input.sparklineGroups) sheet.sparklineGroups.push(...structuredClone(input.sparklineGroups));
      sheet.drawings.push(...structuredClone(input.drawings));
      for (const [key, payload] of Object.entries(input.drawingPayloads)) {
        sheet.drawingPayloads.set(key, structuredClone(payload));
      }
      if (input.hyperlinks) {
        for (const entry of input.hyperlinks) sheet.hyperlinks.set(noteCellKey(entry.row, entry.column), structuredClone(entry.hyperlink));
      }
      if (input.notes) {
        for (const entry of input.notes) sheet.notes.set(noteCellKey(entry.row, entry.column), structuredClone(entry.note));
      }
      if (input.commentThreads) sheet.commentThreads.push(...structuredClone(input.commentThreads));
      if (input.conditionalFormats) sheet.conditionalFormats.push(...structuredClone(input.conditionalFormats));
      if (input.dataValidations) sheet.dataValidations.push(...structuredClone(input.dataValidations));
      sheet.defaultRowHeightPx = input.defaultRowHeightPx;
      sheet.defaultColumnWidthPx = input.defaultColumnWidthPx;
      if (input.rowHeightsPx) Object.assign(sheet.rowHeightsPx, input.rowHeightsPx);
      if (input.columnWidthsPx) Object.assign(sheet.columnWidthsPx, input.columnWidthsPx);
      if (input.hiddenRows) input.hiddenRows.forEach((r) => sheet.hiddenRows.add(r));
      if (input.hiddenColumns) input.hiddenColumns.forEach((c) => sheet.hiddenColumns.add(c));
      if (input.bandedRule) sheet.bandedRule = structuredClone(input.bandedRule);
      if (input.autoFilter) sheet.autoFilter = structuredClone(input.autoFilter);
      if (input.sheetTables) sheet.sheetTables.push(...structuredClone(input.sheetTables));
      if (input.spillRanges) sheet.spillRanges.push(...structuredClone(input.spillRanges));
      if (input.protectionRules) sheet.protectionRules.push(...structuredClone(input.protectionRules));
      if (input.showGridlines != null) sheet.showGridlines = input.showGridlines;
      if (input.showHeaders != null) sheet.showHeaders = input.showHeaders;
      if (input.zoom != null) sheet.zoom = input.zoom;
      if (input.hidden != null) sheet.hidden = input.hidden;
      if (input.outline) sheet.outline = structuredClone(input.outline);
      sheet.tabColor = input.tabColor;
      workbook.sheets.set(sheet.id, sheet);
    }
    for (const document of snapshot.printDocuments ?? []) workbook.setPrintDocument(document);
    for (const definition of snapshot.queryDefinitions ?? []) workbook.setQueryDefinition(definition);
    for (const template of snapshot.cellStyleTemplates ?? []) workbook.setCellStyleTemplate(template);
    workbook.sheetOrder = snapshot.sheets.map((sheet) => sheet.id);
    return workbook;
  }
}
