export type UnitId = string;
export type SheetId = string;
export type Row = number;
export type Column = number;

export interface CellAddress {
  readonly sheetId: SheetId;
  readonly row: Row;
  readonly column: Column;
}

import type {
  CellHyperlink,
  CellNote,
  DrawingObject,
  DrawingPayload,
  DrawingGroup,
  WorksheetSnapSettings,
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
import { DEFAULT_WORKSHEET_SNAP_SETTINGS, isFormulaError, normalizeDefinedNameModel } from './domain';
import type { FormulaErrorCode } from './domain';
import type { WorkbookDimensionMetrics, WorkbookSnapshot } from './snapshot';
import { isCellEditorConfig, type CellEditorConfig } from './cell-editor';
import { DEFAULT_WORKBOOK_EDITING_OPTIONS, normalizeWorkbookEditingOptions, type WorkbookEditingOptions } from './editing-options';
export { ASSET_REF_SCHEMA, assertAssetRef, isAssetRef, type AssetRef } from './asset';
export {
  checkboxStateFromValue,
  checkboxValueForState,
  isCellEditorConfig,
  isUnambiguousCheckboxEditor,
  nextCheckboxValue,
  normalizeCheckboxValue,
  type CellEditorConfig,
  type CellEditorKind,
  type CellEditorOptionValue,
  type CellEditorScalar,
  type CheckboxCellEditorConfig,
  type CheckboxCellState,
} from './cell-editor';
export { DEFAULT_WORKBOOK_EDITING_OPTIONS, isWorkbookEditingOptions, normalizeWorkbookEditingOptions, type WorkbookEditingOptions, type WorkbookEnterDirection } from './editing-options';
import {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
  type QueryLoadTargetSnapshot,
} from './workbook-state';
import { normalizeFontFamily } from './font-family';
import { DEFAULT_SHEET_COLUMN_COUNT, DEFAULT_SHEET_ROW_COUNT, SheetExtent } from './sheet-extent';
import { DEFAULT_WORKBOOK_CALCULATION_SETTINGS, DEFAULT_WORKBOOK_COLLATION, normalizeWorkbookCalculationSettings, normalizeWorkbookCollation, type WorkbookCalculationSettings, type WorkbookCollationContext } from '@react-sheets/formula-engine';
import { planSheetIdentityTransform } from './sheet-identity-transform';
import { ReviewStore } from './review-store';
import type { ReviewStoreSnapshot } from './review-store';

export * from './sheet-extent';
export * from './sheet-identity-transform';
export * from './review-store';

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

/** Canonical workbook-owned theme reference used by cross-workbook formatting operations. */
export interface WorkbookTheme {
  id: string;
  colors: Record<string, string>;
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
  asset: import('./asset').AssetRef;
  altText?: string;
  fit: 'contain' | 'cover' | 'stretch';
  crop?: ImageCrop;
  effects?: ImageEffects;
}

export type CellPresentation = BarcodeCellPresentation | ImageCellPresentation;

export interface RichTextRunStyle extends Pick<CellStyle, 'fontFamily' | 'fontSizePx' | 'bold' | 'italic' | 'underline' | 'strikethrough' | 'textColor'> {
  verticalAlignment?: 'baseline' | 'superscript' | 'subscript';
}

export interface RichTextRun {
  text: string;
  style?: RichTextRunStyle;
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

/**
 * User-authored cell writes replace the formula definition; OOXML provenance
 * belongs only to the imported definition that is being replaced.
 */
export function clearFormulaProvenance(cell: CellData): CellData {
  const next = structuredClone(cell);
  delete next.formulaMetadata;
  return next;
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
  SheetTableModel,
  OutlineGroup,
  OutlineModel,
  DrawingKind,
  DrawingTransform,
  DrawingAnchor,
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
  ShapeDrawingEffects,
  ShapeTextDirection,
  ShapeTextHorizontalAlignment,
  ShapeTextVerticalAlignment,
  ConnectorDrawingPayload,
  DrawingConnectorType,
  DrawingConnectionPoint,
  DrawingArrowhead,
  DrawingConnectionEndpoint,
  DrawingConnectorRoutePoint,
  DrawingConnectorRoute,
  DrawingGroup,
  WorksheetSnapSettings,
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
  isDrawingConnectorPayload,
  isShapeDrawingPayload,
  isDrawingGroup,
  isWorksheetSnapSettings,
  isDrawingConnectionPoint,
} from './domain';
export { DEFAULT_WORKSHEET_SNAP_SETTINGS } from './domain';
export {
  assertCanonicalConnector,
  canonicalSnapSettings,
  planConnectorRoute,
  recomputeConnectorRoutes,
  validateDrawingGraph,
  type ConnectorRoutePlan,
  type ConnectorTransformOverride,
  type DrawingGraphSheet,
} from './drawing-planner';
export { StructuralTransform, planCellShift, type StructuralTransformResult, type CellShiftPlan, ensureDrawing } from './structural-transform';
export { SheetRuleRegistry, sheetRuleRegistry, ruleRangesIntersect, type RuleTransform, type RulePasteTransform, type SheetRule, type SheetRuleKind } from './rule-lifecycle';
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
  /** Canonical origin used to project relative rule references. */
  formulaAnchor?: CellAddress;
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
  /** Canonical origin used to project relative custom/list formulas. */
  formulaAnchor?: CellAddress;
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

interface HeapEntry {
  coordinate: number;
  key: string;
}

/**
 * A tiny binary heap used by sparse bounds indexes. Stale entries are removed
 * lazily, so a cell delete never needs to scan every persisted coordinate.
 */
class CoordinateHeap {
  private readonly entries: HeapEntry[] = [];

  constructor(private readonly before: (left: number, right: number) => boolean) {}

  push(entry: HeapEntry): void {
    this.entries.push(entry);
    let index = this.entries.length - 1;
    while (index > 0) {
      const parent = Math.floor((index - 1) / 2);
      if (!this.before(this.entries[index]!.coordinate, this.entries[parent]!.coordinate)) break;
      [this.entries[index], this.entries[parent]] = [this.entries[parent]!, this.entries[index]!];
      index = parent;
    }
  }

  peek(): HeapEntry | undefined {
    return this.entries[0];
  }

  pop(): HeapEntry | undefined {
    const first = this.entries[0];
    const last = this.entries.pop();
    if (!first) return undefined;
    if (last && this.entries.length > 0) {
      this.entries[0] = last;
      let index = 0;
      while (true) {
        const left = index * 2 + 1;
        const right = left + 1;
        let next = index;
        if (left < this.entries.length && this.before(this.entries[left]!.coordinate, this.entries[next]!.coordinate)) next = left;
        if (right < this.entries.length && this.before(this.entries[right]!.coordinate, this.entries[next]!.coordinate)) next = right;
        if (next === index) break;
        [this.entries[index], this.entries[next]] = [this.entries[next]!, this.entries[index]!];
        index = next;
      }
    }
    return first;
  }

  clear(): void {
    this.entries.length = 0;
  }
}

/** Incremental sparse coordinate bounds with O(log n) updates and amortized O(log n) reads. */
class SparseAxisBounds {
  private readonly counts = new Map<number, number>();
  private readonly minimums = new CoordinateHeap((left, right) => left < right);
  private readonly maximums = new CoordinateHeap((left, right) => left > right);

  add(coordinate: number): void {
    const count = this.counts.get(coordinate) ?? 0;
    this.counts.set(coordinate, count + 1);
    if (count === 0) {
      const entry = { coordinate, key: String(coordinate) };
      this.minimums.push(entry);
      this.maximums.push(entry);
    }
  }

  remove(coordinate: number): void {
    const count = this.counts.get(coordinate);
    if (!count) return;
    if (count === 1) this.counts.delete(coordinate);
    else this.counts.set(coordinate, count - 1);
  }

  get minimum(): number | undefined {
    return this.read(this.minimums);
  }

  get maximum(): number | undefined {
    return this.read(this.maximums);
  }

  clear(): void {
    this.counts.clear();
    this.minimums.clear();
    this.maximums.clear();
  }

  private read(heap: CoordinateHeap): number | undefined {
    while (heap.peek() && !this.counts.has(heap.peek()!.coordinate)) heap.pop();
    return heap.peek()?.coordinate;
  }
}

/**
 * Worksheet-owned block ranges are indexed by their four boundaries. The
 * index intentionally stores no materialized cells for block-backed regions.
 */
class DataRegionBoundsIndex {
  private readonly ranges = new Map<string, RangeRef>();
  private readonly startRows = new CoordinateHeap((left, right) => left < right);
  private readonly endRows = new CoordinateHeap((left, right) => left > right);
  private readonly startColumns = new CoordinateHeap((left, right) => left < right);
  private readonly endColumns = new CoordinateHeap((left, right) => left > right);

  add(region: SheetDataRegion): void {
    if (this.ranges.has(region.id)) throw new Error(`Data region ${region.id} already exists`);
    const range = structuredClone(region.range);
    this.ranges.set(region.id, range);
    this.startRows.push({ coordinate: range.startRow, key: region.id });
    this.endRows.push({ coordinate: range.endRow, key: region.id });
    this.startColumns.push({ coordinate: range.startColumn, key: region.id });
    this.endColumns.push({ coordinate: range.endColumn, key: region.id });
  }

  remove(regionId: string): void {
    this.ranges.delete(regionId);
  }

  clear(): void {
    this.ranges.clear();
    this.startRows.clear();
    this.endRows.clear();
    this.startColumns.clear();
    this.endColumns.clear();
  }

  get range(): RangeRef | undefined {
    const startRow = this.read(this.startRows, 'startRow');
    const endRow = this.read(this.endRows, 'endRow');
    const startColumn = this.read(this.startColumns, 'startColumn');
    const endColumn = this.read(this.endColumns, 'endColumn');
    if (startRow === undefined || endRow === undefined || startColumn === undefined || endColumn === undefined) return undefined;
    return { sheetId: this.ranges.get(this.startRows.peek()!.key)!.sheetId, startRow, endRow, startColumn, endColumn };
  }

  private read(heap: CoordinateHeap, boundary: keyof Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>): number | undefined {
    while (heap.peek() && this.ranges.get(heap.peek()!.key)?.[boundary] !== heap.peek()!.coordinate) heap.pop();
    return heap.peek()?.coordinate;
  }
}

export class CellMatrix {
  private readonly rows = new Map<Row, Map<Column, CellData>>();
  private readonly rowBounds = new SparseAxisBounds();
  private readonly columnBounds = new SparseAxisBounds();
  private cellCount = 0;
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
    const existed = rowMap.has(column);
    rowMap.set(column, normalizedCell);
    if (!existed) {
      this.cellCount += 1;
      this.rowBounds.add(row);
      this.columnBounds.add(column);
    }
    this.revisionCounter += 1;
  }

  delete(row: Row, column: Column): void {
    const rowMap = this.rows.get(row);
    const existed = rowMap?.has(column) ?? false;
    rowMap?.delete(column);
    if (rowMap?.size === 0) this.rows.delete(row);
    if (existed) {
      this.cellCount -= 1;
      this.rowBounds.remove(row);
      this.columnBounds.remove(column);
      this.revisionCounter += 1;
    }
  }

  has(row: Row, column: Column): boolean {
    return this.rows.get(row)?.has(column) ?? false;
  }

  clear(): void {
    if (this.rows.size > 0) this.revisionCounter += 1;
    this.rows.clear();
    this.rowBounds.clear();
    this.columnBounds.clear();
    this.cellCount = 0;
  }

  count(): number {
    return this.cellCount;
  }

  /** Read the persisted-cell extent without walking CellMatrix rows. */
  occupiedRange(sheetId: SheetId): RangeRef {
    const startRow = this.rowBounds.minimum;
    const endRow = this.rowBounds.maximum;
    const startColumn = this.columnBounds.minimum;
    const endColumn = this.columnBounds.maximum;
    return {
      sheetId,
      startRow: startRow ?? 0,
      endRow: endRow ?? 0,
      startColumn: startColumn ?? 0,
      endColumn: endColumn ?? 0,
    };
  }

  forEach(callback: (cell: CellData, row: Row, column: Column) => void): void {
    for (const [row, columns] of this.rows) {
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  forEachInRows(rows: ReadonlySet<Row>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    for (const row of rows) {
      const columns = this.rows.get(row);
      if (!columns) continue;
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  forEachInColumns(columns: ReadonlySet<Column>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    for (const [row, rowCells] of this.rows) {
      for (const column of columns) {
        const cell = rowCells.get(column);
        if (cell) callback(cell, row, column);
      }
    }
  }

  *entriesInColumn(column: Column): IterableIterator<{ row: Row; cell: CellData }> {
    for (const [row, rowCells] of this.rows) {
      const cell = rowCells.get(column);
      if (cell) yield { row, cell };
    }
  }

  /** Enumerate only persisted cells inside a range; implicit cells are not materialized. */
  forEachInRange(
    startRow: Row,
    endRow: Row,
    startColumn: Column,
    endColumn: Column,
    callback: (cell: CellData, row: Row, column: Column) => void,
  ): void {
    for (const [row, columns] of this.rows) {
      if (row < startRow || row > endRow) continue;
      for (const [column, cell] of columns) {
        if (column >= startColumn && column <= endColumn) callback(cell, row, column);
      }
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
  private readonly dataRegionStore: SheetDataRegion[] = [];
  private readonly dataRegionBounds = new DataRegionBoundsIndex();
  readonly merges: MergeSpan[] = [];
  readonly pivots: PivotModel[] = [];
  readonly sparklines: SparklineModel[] = [];
  readonly conditionalFormats: ConditionalFormatRule[] = [];
  readonly dataValidations: DataValidationRule[] = [];
  readonly sheetTables: SheetTableModel[] = [];
  readonly drawings: DrawingObject[] = [];
  readonly drawingPayloads = new Map<string, DrawingPayload>();
  readonly drawingGroups: DrawingGroup[] = [];
  snapSettings: WorksheetSnapSettings = structuredClone(DEFAULT_WORKSHEET_SNAP_SETTINGS);
  /** Canonical persisted hyperlink metadata keyed by row:column. */
  readonly hyperlinks = new Map<string, CellHyperlink>();
  readonly review: ReviewStore;
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
    copy.replaceDataRegions(this.dataRegions);
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
    copy.drawingGroups.push(...structuredClone(this.drawingGroups));
    copy.snapSettings = structuredClone(this.snapSettings);
    for (const [key, hyperlink] of this.hyperlinks) copy.hyperlinks.set(key, structuredClone(hyperlink));
    copy.review.replaceNotes(this.review.noteEntries());
    copy.review.replaceThreads(this.review.threadEntries());
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

  private readonly extent: SheetExtent;

  constructor(
    readonly id: SheetId,
    public name: string,
    rowCount: number = DEFAULT_SHEET_ROW_COUNT,
    columnCount: number = DEFAULT_SHEET_COLUMN_COUNT,
  ) {
    this.extent = new SheetExtent(rowCount, columnCount);
    this.review = new ReviewStore(id);
  }

  get rowCount(): number { return this.extent.rowCount; }
  set rowCount(value: number) { this.extent.rowCount = value; }
  get columnCount(): number { return this.extent.columnCount; }
  set columnCount(value: number) { this.extent.columnCount = value; }

  get sheetExtent(): SheetExtent { return this.extent; }

  ensureCellExtent(row: number, column: number): void {
    this.extent.ensureCell(row, column);
  }

  ensureRangeExtent(startRow: number, endRow: number, startColumn: number, endColumn: number): void {
    this.extent.ensureRange(startRow, endRow, startColumn, endColumn);
  }

  /** Read-only block-backed region view; all mutations must update the range index below. */
  get dataRegions(): readonly SheetDataRegion[] {
    return this.dataRegionStore;
  }

  addDataRegion(region: SheetDataRegion, index = this.dataRegionStore.length): void {
    if (region.range.sheetId !== this.id) throw new Error(`Data region ${region.id} belongs to ${region.range.sheetId}, not worksheet ${this.id}`);
    const copy = structuredClone(region);
    this.dataRegionBounds.add(copy);
    this.dataRegionStore.splice(Math.min(Math.max(index, 0), this.dataRegionStore.length), 0, copy);
  }

  removeDataRegionAt(index: number): SheetDataRegion | undefined {
    const region = this.dataRegionStore[index];
    if (!region) return undefined;
    this.dataRegionStore.splice(index, 1);
    this.dataRegionBounds.remove(region.id);
    return structuredClone(region);
  }

  replaceDataRegions(regions: readonly SheetDataRegion[]): void {
    const copies = regions.map((region) => structuredClone(region));
    const nextBounds = new DataRegionBoundsIndex();
    for (const region of copies) {
      if (region.range.sheetId !== this.id) throw new Error(`Data region ${region.id} belongs to ${region.range.sheetId}, not worksheet ${this.id}`);
      nextBounds.add(region);
    }
    this.dataRegionStore.splice(0, this.dataRegionStore.length, ...copies);
    this.dataRegionBounds.clear();
    for (const region of copies) this.dataRegionBounds.add(region);
  }

  /** One incremental used-range authority for cells and block-backed regions. */
  get usedRange(): RangeRef {
    const cells = this.cells.occupiedRange(this.id);
    const regions = this.dataRegionBounds.range;
    if (!regions) return cells;
    if (this.cells.count() === 0) return regions;
    return {
      sheetId: this.id,
      startRow: Math.min(cells.startRow, regions.startRow),
      endRow: Math.max(cells.endRow, regions.endRow),
      startColumn: Math.min(cells.startColumn, regions.startColumn),
      endColumn: Math.max(cells.endColumn, regions.endColumn),
    };
  }

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

export function cellKey(row: Row, column: Column): string {
  return `${row}:${column}`;
}

export function getDrawingPayload(sheet: WorksheetModel, payloadId: string): DrawingPayload | undefined {
  return sheet.drawingPayloads.get(payloadId);
}

export function getCellNote(sheet: WorksheetModel, row: Row, column: Column): CellNote | undefined {
  return sheet.review.getNoteAt(row, column);
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
  drawingGroups?: DrawingGroup[];
  snapSettings?: WorksheetSnapSettings;
  hyperlinks?: Array<{ row: number; column: number; hyperlink: CellHyperlink }>;
  review: ReviewStoreSnapshot;
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
  /** Lifecycle inverse payload; owned workbook documents travel with the sheet. */
  lifecycleDefinedNames?: DefinedNameModel[];
  lifecyclePrintDocument?: PrintDocumentSnapshot;
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
  collationContext: WorkbookCollationContext = normalizeWorkbookCollation(DEFAULT_WORKBOOK_COLLATION);
  /** Canonical authored calculation policy shared by the runtime and workers. */
  calculationSettings: WorkbookCalculationSettings = normalizeWorkbookCalculationSettings(DEFAULT_WORKBOOK_CALCULATION_SETTINGS);
  editingOptions: WorkbookEditingOptions = normalizeWorkbookEditingOptions(DEFAULT_WORKBOOK_EDITING_OPTIONS);
  /** The sole theme owner. Clipboard and OOXML boundaries carry a reference to this state. */
  theme: WorkbookTheme = { id: 'workbook-theme-default', colors: {} };

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

  setCalculationSettings(settings: Partial<WorkbookCalculationSettings>): void {
    this.calculationSettings = normalizeWorkbookCalculationSettings({ ...this.calculationSettings, ...settings });
  }

  setEditingOptions(options: WorkbookEditingOptions): void {
    this.editingOptions = normalizeWorkbookEditingOptions(options);
  }

  setTheme(theme: WorkbookTheme): void {
    if (!theme.id.trim()) throw new Error('Workbook theme id is required');
    for (const [key, color] of Object.entries(theme.colors)) {
      if (!key.trim() || !/^#[0-9a-f]{6}$/i.test(color)) throw new Error('Workbook theme color is invalid');
    }
    this.theme = structuredClone({ id: theme.id.trim(), colors: theme.colors });
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
    if (template.editor && !isCellEditorConfig(template.editor)) {
      throw new Error('Cell style template editor is invalid');
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
    const normalized = name.trim().toUpperCase();
    if (sheetId) {
      const local = this.definedNameModels.find((entry) => entry.scope === 'sheet'
        && entry.sheetId === sheetId
        && entry.name.toUpperCase() === normalized);
      if (local) return structuredClone(local);
    }
    const global = this.definedNameModels.find((entry) => entry.scope === 'workbook'
      && entry.name.toUpperCase() === normalized);
    return global ? structuredClone(global) : undefined;
  }

  getDefinedNameExact(name: string, scope: DefinedNameScope, sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toUpperCase();
    const exact = this.definedNameModels.find((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toUpperCase() === normalized);
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
      && entry.name.toUpperCase() === model.name.toUpperCase());
    if (index >= 0) this.definedNameModels[index] = structuredClone(model);
    else this.definedNameModels.push(structuredClone(model));
    return structuredClone(model);
  }

  removeDefinedName(name: string, scope: DefinedNameScope = 'workbook', sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toUpperCase();
    const index = this.definedNameModels.findIndex((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toUpperCase() === normalized);
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

  addSheet(id: SheetId, name: string, rowCount: number = DEFAULT_SHEET_ROW_COUNT, columnCount: number = DEFAULT_SHEET_COLUMN_COUNT): WorksheetModel {
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
    const plan = planSheetIdentityTransform(this, {
      kind: 'duplicate',
      sourceSheetId,
      sourceName: source.name,
      targetSheetId: newId,
      targetName: newName,
    });
    plan.apply();
    return this.getSheet(newId);
  }

  reorderSheet(sheetId: SheetId, toIndex: number): void {
    const fromIndex = this.sheetOrder.indexOf(sheetId);
    if (fromIndex < 0) throw new Error(`Unknown sheet: ${sheetId}`);
    const clamped = Math.max(0, Math.min(toIndex, this.sheetOrder.length - 1));
    this.sheetOrder.splice(fromIndex, 1);
    this.sheetOrder.splice(clamped, 0, sheetId);
  }

  removeSheet(sheetId: SheetId): WorksheetModel {
    const sheet = this.getSheet(sheetId);
    const plan = planSheetIdentityTransform(this, {
      kind: 'delete',
      sourceSheetId: sheetId,
      sourceName: sheet.name,
    });
    plan.apply();
    return sheet;
  }

  renameSheet(sheetId: SheetId, name: string): void {
    const source = this.getSheet(sheetId);
    planSheetIdentityTransform(this, {
      kind: 'rename',
      sourceSheetId: sheetId,
      sourceName: source.name,
      targetName: name,
    }).apply();
  }

  getSheetSnapshot(sheetId: SheetId): SheetSnapshot {
    const sheet = this.snapshot().sheets.find((entry) => entry.id === sheetId);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetId}`);
    sheet.lifecycleDefinedNames = structuredClone(this.definedNameModels.filter((entry) => entry.scope === 'sheet' && entry.sheetId === sheetId));
    const printDocument = this.printDocuments.get(sheetId);
    if (printDocument) sheet.lifecyclePrintDocument = structuredClone(printDocument);
    return structuredClone(sheet);
  }

  restoreSheetSnapshot(snapshot: SheetSnapshot, index = this.sheetOrder.length): void {
    if (this.sheets.has(snapshot.id)) throw new Error(`Sheet already exists: ${snapshot.id}`);
    const current = this.snapshot();
    const hydrated = WorkbookModel.fromSnapshot({ ...current, sheets: [structuredClone(snapshot)] });
    const sheet = hydrated.getSheet(snapshot.id);
    this.sheets.set(sheet.id, sheet);
    if (snapshot.lifecycleDefinedNames) this.definedNameModels.push(...structuredClone(snapshot.lifecycleDefinedNames));
    if (snapshot.lifecyclePrintDocument) this.printDocuments.set(snapshot.id, structuredClone(snapshot.lifecyclePrintDocument));
    const bounded = Math.max(0, Math.min(index, this.sheetOrder.length));
    this.sheetOrder.splice(bounded, 0, sheet.id);
  }

  snapshot(): WorkbookSnapshot {
    return {
      schema: 'WorkbookSnapshot',
      version: 9,
      unitId: this.unitId,
      name: this.name,
      dimensionMetrics: structuredClone(this.dimensionMetrics),
      collationContext: structuredClone(this.collationContext),
      calculationSettings: structuredClone(this.calculationSettings),
      editingOptions: structuredClone(this.editingOptions),
      theme: structuredClone(this.theme),
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
        dataRegions: sheet.dataRegions.map((region) => structuredClone(region)),
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
        drawingGroups: structuredClone(sheet.drawingGroups),
        snapSettings: structuredClone(sheet.snapSettings),
        hyperlinks: [...sheet.hyperlinks.entries()].map(([key, hyperlink]) => {
          const [row, column] = key.split(':').map(Number);
          return { row: row!, column: column!, hyperlink: structuredClone(hyperlink) };
        }),
        review: sheet.review.toSnapshot(),
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
    if (snapshot.version !== 9) throw new Error('Unsupported workbook snapshot version');
    if (snapshot.sheets.length === 0) throw new Error('Workbook snapshot must contain at least one sheet');
    const workbook = new WorkbookModel(snapshot.unitId, snapshot.name);
    workbook.dimensionMetrics = structuredClone(snapshot.dimensionMetrics);
    if (snapshot.theme) workbook.setTheme(snapshot.theme);
    workbook.collationContext = normalizeWorkbookCollation(snapshot.collationContext ?? DEFAULT_WORKBOOK_COLLATION);
    workbook.setCalculationSettings(snapshot.calculationSettings);
    workbook.setEditingOptions(snapshot.editingOptions);
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
        if (legacy) sheet.hyperlinks.set(cellKey(row, column), legacy);
      });
      if (input.dataRegions) sheet.replaceDataRegions(input.dataRegions);
      sheet.merges.push(...structuredClone(input.merges));
      sheet.pane = normalizeWorksheetPane(input.pane);
      sheet.pivots.push(...input.pivots.map((pivot) => canonicalizePivotDefinition(structuredClone(pivot))));
      sheet.sparklines.push(...structuredClone(input.sparklines));
      if (input.sparklineGroups) sheet.sparklineGroups.push(...structuredClone(input.sparklineGroups));
      sheet.drawings.push(...structuredClone(input.drawings));
      for (const [key, payload] of Object.entries(input.drawingPayloads)) {
        sheet.drawingPayloads.set(key, structuredClone(payload));
      }
      if (input.drawingGroups) sheet.drawingGroups.push(...structuredClone(input.drawingGroups));
      sheet.snapSettings = input.snapSettings ? structuredClone(input.snapSettings) : structuredClone(DEFAULT_WORKSHEET_SNAP_SETTINGS);
      if (input.hyperlinks) {
        for (const entry of input.hyperlinks) sheet.hyperlinks.set(cellKey(entry.row, entry.column), structuredClone(entry.hyperlink));
      }
      const review = ReviewStore.fromSnapshot(input.id, input.review);
      sheet.review.replaceNotes(review.noteEntries());
      sheet.review.replaceThreads(review.threadEntries());
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
