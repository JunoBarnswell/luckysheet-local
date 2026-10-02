import { assertExternalLinkBinding } from './data-model';
import { assertRecordTable, assertRecordRelationship, assertRecordCalculations } from './record-domain';
import { rewriteSheetLifecycleFormula } from '@react-sheets/formula-engine';
import { planWorkbookFormulaRewrite } from './structural-transform';
import { parseCellMatrixCoordinate } from './cell-coordinates';

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
import { assertCanonicalDefinedNameModels, assertCanonicalWorksheetIdentities, assertCanonicalWorkbookOwnerIdentities, type WorkbookDimensionMetrics, type WorkbookSnapshot } from './snapshot';
import { isCellEditorConfig, type CellEditorConfig } from './cell-editor';
import { DEFAULT_WORKBOOK_EDITING_OPTIONS, normalizeWorkbookEditingOptions, type WorkbookEditingOptions } from './editing-options';
export { ASSET_REF_SCHEMA, assertAssetRef, isAssetRef, isSupportedAssetMime, type AssetRef } from './asset';
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
export { CALCULATION_CONTEXT_EFFECTS, isWorkbookCalculationContextEffect, type WorkbookCalculationContextAction, type WorkbookCalculationContextEffect } from './calculation-context-effect';
import {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
  type QueryLoadTargetSnapshot,
} from './workbook-state';
import { normalizeFontFamily } from './font-family';
import { DEFAULT_SHEET_COLUMN_COUNT, DEFAULT_SHEET_ROW_COUNT, SheetExtent } from './sheet-extent';
import { DEFAULT_WORKBOOK_CALCULATION_SETTINGS, DEFAULT_WORKBOOK_COLLATION, MAX_COLUMN_INDEX, MAX_ROW_INDEX, normalizeWorkbookCalculationSettings, normalizeWorkbookCollation, type WorkbookCalculationSettings, type WorkbookCollationContext } from '@react-sheets/formula-engine';
import { planSheetIdentityTransform, SheetIdentityTransformInvariantError } from './sheet-identity-transform';
import type { StructuralTransformResult } from './structural-transform';
import { ReviewStore } from './review-store';
import type { ReviewStoreSnapshot } from './review-store';

export * from './sheet-extent';
export * from './sheet-identity-transform';
export * from './review-store';
export * from './chart-text-reference';

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
  style: 'hair' | 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'dashDot' | 'dashDotDot' | 'double';
  color: string;
}

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
  diagonal?: CellBorderSide;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export type CellUnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
export type CellFillPattern = 'solid' | 'none' | 'gray125' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis' | 'gray0625' | 'lightGray' | 'darkGray' | 'mediumGray';

export interface CellFill {
  kind: 'solid' | 'pattern' | 'gradient';
  foreground?: string;
  background?: string;
  pattern?: CellFillPattern;
  gradientType?: 'linear' | 'path';
  degree?: number;
  stops?: Array<{ position: number; color: string }>;
}

export interface CellNumberFormatSpec {
  category?: 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom';
  locale?: string;
  decimalPlaces?: number;
  useThousandsSeparator?: boolean;
  negativeStyle?: 'minus' | 'parentheses' | 'red-minus' | 'red-parentheses';
  currencySymbol?: string;
  fractionType?: 'up-to-one-digit' | 'up-to-two-digits' | 'up-to-three-digits' | 'as-halves' | 'as-quarters' | 'as-eighths' | 'as-sixteenths' | 'as-tenths' | 'as-hundredths' | 'as-thousandths';
  sample?: string;
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
  underlineStyle?: CellUnderlineStyle;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontTheme?: string;
  textDirection?: 'context' | 'ltr' | 'rtl';
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
  numberFormatSpec?: CellNumberFormatSpec;
  borders?: CellBorders;
  fill?: CellFill;
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
  /** East Asian phonetic guide runs. Visibility is a render/edit property, not cell text. */
  phonetic?: import('./phonetic').CellPhoneticMetadata;
  /** OOXML formula provenance used to preserve cached values for Excel-only formula families. */
  formulaMetadata?: FormulaMetadata;
  /** 公式引擎结果（含错误）。禁止再用 error: string 当真相 */
  formulaValue?: import('./domain').FormulaValue;
  /** Native AutoFilter color/icon identity resolved at the import boundary. */
  filterMetadata?: {
    color?: { target: 'cell' | 'font'; dxfId?: number; value?: string };
    icon?: { iconSet: string; iconId: number };
  };
}

/** Apply the canonical model normalization shared by CellMatrix writes and structural preflight. */
export function normalizeCellDataForStorage(cell: CellData): CellData {
  const fontFamily = cell.style?.fontFamily;
  return fontFamily === undefined
    ? cell
    : { ...cell, style: { ...cell.style, fontFamily: normalizeFontFamily(fontFamily) } };
}

export type { CellPhoneticMetadata, PhoneticAlignment, PhoneticRun, PhoneticType } from './phonetic';
export { isCellPhoneticMetadata } from './phonetic';

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

export function hasFormulaGroupMetadata(cell: Pick<CellData, 'formulaMetadata'>): boolean {
  const metadata = cell.formulaMetadata;
  return metadata !== undefined
    && (metadata.preservedOnly === true || metadata.kind !== 'normal' || metadata.range !== undefined);
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

export function worksheetPaneValidationError(value: unknown): string | undefined {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return 'kind';
  const pane = value as Record<string, unknown>;
  const hasOwn = (field: string): boolean => Object.prototype.hasOwnProperty.call(pane, field);
  if (!hasOwn('kind')) return 'kind';
  if (pane.kind === 'none') {
    const fields = Object.keys(pane);
    if (fields.some((field) => field !== 'kind')) {
      return fields.some((field) => ['state', 'xSplit', 'ySplit', 'startRow', 'startColumn', 'activePane'].includes(field))
        ? 'none-state'
        : 'unknown-field';
    }
    return undefined;
  }
  if (pane.kind !== 'frozen' && pane.kind !== 'split') return 'kind';
  const allowedFields = ['kind', 'state', 'xSplit', 'ySplit', 'startRow', 'startColumn', 'activePane'];
  if (Object.keys(pane).some((field) => !allowedFields.includes(field))) return 'unknown-field';
  for (const field of ['state', 'xSplit', 'ySplit', 'startRow', 'startColumn']) {
    if (!hasOwn(field)) return field;
  }
  if (pane.kind === 'frozen' && pane.state !== 'frozen' && pane.state !== 'frozenSplit') return 'state';
  if (pane.kind === 'split' && pane.state !== 'split') return 'state';
  if (!Number.isSafeInteger(pane.startRow) || (pane.startRow as number) < 0 || (pane.startRow as number) > MAX_ROW_INDEX) return 'startRow';
  if (!Number.isSafeInteger(pane.startColumn) || (pane.startColumn as number) < 0 || (pane.startColumn as number) > MAX_COLUMN_INDEX) return 'startColumn';
  if (pane.kind === 'frozen') {
    if (!Number.isSafeInteger(pane.xSplit) || (pane.xSplit as number) < 0 || (pane.xSplit as number) > MAX_COLUMN_INDEX + 1) return 'xSplit';
    if (!Number.isSafeInteger(pane.ySplit) || (pane.ySplit as number) < 0 || (pane.ySplit as number) > MAX_ROW_INDEX + 1) return 'ySplit';
  } else {
    if (typeof pane.xSplit !== 'number' || !Number.isFinite(pane.xSplit) || pane.xSplit < 0) return 'xSplit';
    if (typeof pane.ySplit !== 'number' || !Number.isFinite(pane.ySplit) || pane.ySplit < 0) return 'ySplit';
  }
  if (hasOwn('activePane') && pane.activePane !== undefined
    && !['topLeft', 'topRight', 'bottomLeft', 'bottomRight'].includes(pane.activePane as string)) return 'activePane';
  return undefined;
}

export function isCanonicalWorksheetPane(value: unknown): value is WorksheetPane {
  return worksheetPaneValidationError(value) === undefined;
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
  P1ChartType,
  ChartNativeIdentity,
  ChartSubtype,
  ChartSeriesType,
  ChartAxisType,
  ChartAxisCross,
  ChartTickMark,
  ChartTickLabelPosition,
  ChartAxisModel,
  ChartGridlineModel,
  ChartAreaStyle,
  ChartFillKind,
  ChartFillModel,
  ChartEffectModel,
  ChartLineStyle,
  ChartTextModel,
  ChartTextFormulaField,
  ChartMarkerModel,
  ChartTrendlineModel,
  ChartErrorBarsModel,
  ChartDataLabelsModel,
  ChartDataLabelPosition,
  ChartDataLabelTarget,
  ChartDataTableModel,
  ChartPointModel,
  ChartStockRoles,
  ChartHistogramOptions,
  ChartBoxWhiskerOptions,
  ChartWaterfallOptions,
  ChartMapCoordinate,
  ChartMapRing,
  ChartMapFeature,
  ChartMapResource,
  ChartMapOptions,
  ChartSeriesModel,
  ChartElementModel,
  ChartAggregate,
  ChartBindingArea,
  ChartBindings,
  ChartSource,
  ChartFieldBinding,
  CameraDrawingPayload,
  ScreenshotDrawingPayload,
  IconDrawingPayload,
  Model3dGeometry,
  Model3dDrawingPayload,
  SmartArtNode,
  SmartArtDrawingPayload,
  WordArtDrawingPayload,
  SignatureLineDrawingPayload,
  EmbeddedObjectDrawingPayload,
  EquationDrawingPayload,
  LocalDrawingObjectKind,
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
  FormulaError,
  FormulaErrorCode,
  FormulaValue,
  StructuralOpKind,
  CellShiftSpec,
  StructuralTransformParams,
} from './domain';
export { createDefaultTextBoxTextFrame, SHAPE_DRAWING_PRESETS, isShapeDrawingType } from './domain';
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
  isIconDrawingPayload,
  isModel3dDrawingPayload,
  isSmartArtDrawingPayload,
  isWordArtDrawingPayload,
  isSignatureLineDrawingPayload,
  isEmbeddedObjectDrawingPayload,
  isEquationDrawingPayload,
  isScreenshotDrawingPayload,
  isDrawingConnectorPayload,
  isShapeDrawingPayload,
  isDrawingGroup,
  isWorksheetSnapSettings,
  isDrawingConnectionPoint,
  CHART_SUBTYPES_BY_TYPE,
  defaultChartSubtype,
  isChartSubtypeForType,
  isChartHistogramOptions,
  chartStackingForSubtype,
  chartSeriesSupportsErrorBars,
  chartSeriesSupportsTrendlines,
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
export {
  StructuralTransform,
  planSheetTableRename,
  planCellShift,
  type StructuralTransformResult,
  type StructuralFormulaOwnerDelta,
  type StructuralFormulaCellOwnerDelta,
  type StructuralFormulaRuleOwnerDelta,
  type StructuralFormulaRuleFormulaOwnerDelta,
  type StructuralFormulaRuleAnchorOwnerDelta,
  type StructuralFormulaObjectOwnerDelta,
  type StructuralDefinedNameOwnerDelta,
  type StructuralFormulaOwnerState,
  type StructuralReferenceOwnerAddress,
  type StructuralReferenceOwnerIndex,
  type SheetTableRenamePlan,
  type CellShiftPlan,
  ensureDrawing,
} from './structural-transform';
export { StructuralMutationApplyError } from './structural-mutation-apply-error';
export { structuralRuleFormulaFields, type StructuralFormulaRule, type StructuralFormulaRuleField } from './structural-formula-owner';
export { structuralRangeOwnerAffectedRanges, type StructuralRangeOwnerDelta } from './structural-range-owner';
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
export {
  applyRowPermutation,
  createRowPermutationPlan,
  rowPermutationAffectedColumnEnd,
  validatePermutationMetadata,
  type RowPermutationPlan,
  type RowPermutationResult,
} from './data-transform';
export { columnLabel, parseColumnLabel, cellAddress, parseAddress, a1Range } from './address';
export {
  loadWorkbookFromSnapshot,
  createWorkbookSnapshot,
  migrateStoredWorkbookSnapshot,
  assertCanonicalWorkbookHyperlinks,
  assertCanonicalWorkbookSnapshot,
  MAX_CHANGED_CELLS,
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
export { buildExplicitChartSeries, resolveWorksheetChartRanges, retargetChartPayload, validateChartVector } from './chart-range-bindings';
import type { AnalysisViewDefinition, GanttSheetDefinition, ReportSheetDefinition, TableSheetDefinition, WorkbookDataModel, WorkbookTableModel } from './data-model';
import { normalizeDataSourceManifest, type DataSourceManifest, type SheetDataRegion } from './data-source';
export * from './data-model';
export * from './data-source';

export type SheetKind = 'worksheet' | 'table-sheet' | 'gantt-sheet' | 'report-sheet';

export type SparklineDataOrientation = 'rows' | 'columns';
export type SparklineEmptyCells = 'gap' | 'zero' | 'connect';
export type SparklineHiddenCells = 'show' | 'hide';
export type SparklineAxisBoundsMode = 'automatic' | 'same-group' | 'custom';

export interface SparklineAxisBounds {
  mode: SparklineAxisBoundsMode;
  minimum?: number;
  maximum?: number;
}

export interface SparklineColors {
  series?: string;
  negative?: string;
  first?: string;
  last?: string;
  high?: string;
  low?: string;
  marker?: string;
  axis?: string;
}

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
  lineWeight?: number;
  dateAxis?: boolean;
  dataOrientation?: SparklineDataOrientation;
  rightToLeft?: boolean;
  hiddenCells?: SparklineHiddenCells;
  emptyCells?: SparklineEmptyCells;
  verticalAxis?: SparklineAxisBounds;
  axisColor?: string;
  firstColor?: string;
  lastColor?: string;
  highColor?: string;
  lowColor?: string;
  markerColor?: string;
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

function firstCoordinateAtLeast(coordinates: readonly number[], target: number): number {
  let low = 0;
  let high = coordinates.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (coordinates[middle]! < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function assertCellMatrixCoordinates(row: Row, column: Column): void {
  if (!Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0) {
    throw new Error(`CELL_MATRIX_INVALID_COORDINATE: cell address ${row}:${column} must use non-negative safe integers`);
  }
}

export class CellMatrix {
  private readonly rows = new Map<Row, Map<Column, CellData>>();
  private readonly rowBounds = new SparseAxisBounds();
  private readonly columnBounds = new SparseAxisBounds();
  private sortedRowCoordinates?: Row[];
  private cellCount = 0;
  private revisionCounter = 0;
  private deferredJSON?: Record<string, Record<string, CellData>>;
  private deferredRowCoordinates?: Row[];
  private deferredMaterializedCells = new WeakSet<CellData>();
  private readonly deferredCellOverlays = new Map<string, Map<string, CellData | null>>();
  private readonly deferredCellCountsByRow = new Map<string, number>();
  private deferredRevision = 0;
  private deferredBounds?: {
    count: number;
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
  };

  constructor(private readonly onWrite?: (row: Row, column: Column) => void) {}

  /** Whether indexed cell storage has been materialized; sparse writes can remain deferred. */
  get isHydrated(): boolean {
    return this.deferredJSON === undefined;
  }

  /**
   * Keep the canonical sparse wire object intact until a caller needs cell
   * semantics. The owning worksheet remains responsible for its extent.
   */
  deferJSON(input: Record<string, Record<string, CellData>> | undefined): void {
    if (this.rows.size > 0 || this.deferredJSON !== undefined) throw new Error('CellMatrix already contains data');
    this.deferredJSON = input ?? {};
    this.deferredRowCoordinates = undefined;
    this.deferredMaterializedCells = new WeakSet<CellData>();
    this.deferredCellOverlays.clear();
    this.deferredCellCountsByRow.clear();
    this.deferredRevision = 0;
    this.sortedRowCoordinates = undefined;
    this.deferredBounds = undefined;
  }

  /** Monotonic content revision used by derived caches; it is not persisted. */
  get revision(): number {
    return this.deferredJSON !== undefined
      ? this.revisionCounter + this.deferredRevision
      : this.revisionCounter;
  }

  /** Read one persisted cell, materializing and normalizing only that cell when storage is deferred. */
  get(row: Row, column: Column): CellData | undefined {
    const cell = this.getWithoutHydration(row, column);
    if (!cell || this.deferredJSON === undefined) return cell;
    if (this.deferredMaterializedCells.has(cell)) return cell;

    const materialized = normalizeCellDataForStorage(structuredClone(cell));
    this.writeDeferredCell(row, column, materialized);
    this.deferredMaterializedCells.add(materialized);
    return materialized;
  }

  /** Read a persisted sparse cell without materializing deferred worksheet data. */
  getWithoutHydration(row: Row, column: Column): CellData | undefined {
    assertCellMatrixCoordinates(row, column);
    return this.deferredJSON !== undefined
      ? this.getDeferredCell(row, column)
      : this.rows.get(row)?.get(column);
  }

  /** Read a formula-bearing persisted cell without materializing deferred worksheet data. */
  getFormulaOwnerWithoutHydration(row: Row, column: Column): CellData | undefined {
    const cell = this.getWithoutHydration(row, column);
    if (!cell) return undefined;
    return cell.formula !== undefined || cell.formulaMetadata?.sourceFormula !== undefined
      || (cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula')
      ? cell
      : undefined;
  }

  /** Replace one existing formula-bearing cell while preserving deferred sparse storage. */
  replaceFormulaOwnerWithoutHydration(row: Row, column: Column, replacement: CellData): boolean {
    if (this.getFormulaOwnerWithoutHydration(row, column) === undefined) return false;
    return this.replaceCellWithoutHydration(row, column, replacement);
  }

  /** Replace one existing sparse cell without materializing deferred worksheet data. */
  replaceCellWithoutHydration(row: Row, column: Column, replacement: CellData): boolean {
    if (this.getWithoutHydration(row, column) === undefined) return false;
    this.set(row, column, replacement);
    return true;
  }

  set(row: Row, column: Column, cell: CellData): void {
    this.writeNormalizedCell(row, column, normalizeCellDataForStorage(cell));
  }

  private writeNormalizedCell(row: Row, column: Column, cell: CellData, incrementRevision = true): void {
    assertCellMatrixCoordinates(row, column);
    this.onWrite?.(row, column);
    if (this.deferredJSON !== undefined) {
      const existed = this.getWithoutHydration(row, column) !== undefined;
      const rowKey = String(row);
      if (!existed) {
        const rowCount = this.getDeferredCellCountForRow(rowKey) + 1;
        this.deferredCellCountsByRow.set(rowKey, rowCount);
        if (rowCount === 1) this.deferredRowCoordinates = undefined;
        if (this.deferredBounds) {
          const bounds = this.deferredBounds;
          const wasEmpty = bounds.count === 0;
          bounds.count += 1;
          bounds.startRow = wasEmpty ? row : Math.min(bounds.startRow, row);
          bounds.endRow = wasEmpty ? row : Math.max(bounds.endRow, row);
          bounds.startColumn = wasEmpty ? column : Math.min(bounds.startColumn, column);
          bounds.endColumn = wasEmpty ? column : Math.max(bounds.endColumn, column);
        }
      }
      this.writeDeferredCell(row, column, cell);
      this.deferredMaterializedCells.add(cell);
      this.deferredRevision += 1;
      return;
    }
    let rowMap = this.rows.get(row);
    if (!rowMap) {
      rowMap = new Map<Column, CellData>();
      this.rows.set(row, rowMap);
      this.sortedRowCoordinates = undefined;
    }
    const existed = rowMap.has(column);
    rowMap.set(column, cell);
    if (!existed) {
      this.cellCount += 1;
      this.rowBounds.add(row);
      this.columnBounds.add(column);
    }
    if (incrementRevision) this.revisionCounter += 1;
  }

  delete(row: Row, column: Column): void {
    assertCellMatrixCoordinates(row, column);
    if (this.deferredJSON !== undefined) {
      if (this.getWithoutHydration(row, column) === undefined) return;
      const rowKey = String(row);
      const remaining = this.getDeferredCellCountForRow(rowKey) - 1;
      this.deferredCellCountsByRow.set(rowKey, remaining);
      this.writeDeferredCell(row, column, null);
      if (remaining === 0) this.deferredRowCoordinates = undefined;
      this.deferredBounds = undefined;
      this.deferredRevision += 1;
      return;
    }
    const rowMap = this.rows.get(row);
    const existed = rowMap?.has(column) ?? false;
    rowMap?.delete(column);
    if (rowMap?.size === 0) {
      this.rows.delete(row);
      this.sortedRowCoordinates = undefined;
    }
    if (existed) {
      this.cellCount -= 1;
      this.rowBounds.remove(row);
      this.columnBounds.remove(column);
      this.revisionCounter += 1;
    }
  }

  /** Check sparse-cell presence without materializing deferred worksheet data. */
  has(row: Row, column: Column): boolean {
    return this.getWithoutHydration(row, column) !== undefined;
  }

  clear(): void {
    if (this.deferredJSON !== undefined) {
      const revision = this.revision;
      const hadCells = this.count() > 0;
      this.revisionCounter = revision + (hadCells ? 1 : 0);
      this.deferredJSON = undefined;
      this.deferredRowCoordinates = undefined;
      this.deferredBounds = undefined;
      this.sortedRowCoordinates = undefined;
      this.rowBounds.clear();
      this.columnBounds.clear();
      this.cellCount = 0;
      this.deferredMaterializedCells = new WeakSet<CellData>();
      this.deferredCellOverlays.clear();
      this.deferredCellCountsByRow.clear();
      this.deferredRevision = 0;
      return;
    }
    this.hydrate();
    if (this.rows.size > 0) this.revisionCounter += 1;
    this.rows.clear();
    this.rowBounds.clear();
    this.columnBounds.clear();
    this.sortedRowCoordinates = undefined;
    this.cellCount = 0;
  }

  count(): number {
    if (this.deferredJSON !== undefined) return this.getDeferredBounds().count;
    return this.cellCount;
  }

  /** Read the persisted-cell extent without walking CellMatrix rows. */
  occupiedRange(sheetId: SheetId): RangeRef {
    if (this.deferredJSON !== undefined) {
      const bounds = this.getDeferredBounds();
      return {
        sheetId,
        startRow: bounds.startRow,
        endRow: bounds.endRow,
        startColumn: bounds.startColumn,
        endColumn: bounds.endColumn,
      };
    }
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
    this.hydrate();
    for (const [row, columns] of this.rows) {
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  private writeDeferredCell(row: Row, column: Column, cell: CellData | null): void {
    const rowKey = String(row);
    let overlay = this.deferredCellOverlays.get(rowKey);
    if (!overlay) {
      overlay = new Map<string, CellData | null>();
      this.deferredCellOverlays.set(rowKey, overlay);
    }
    overlay.set(String(column), cell);
  }

  private getDeferredCellCountForRow(rowKey: string): number {
    const cached = this.deferredCellCountsByRow.get(rowKey);
    if (cached !== undefined) return cached;
    const base = this.deferredJSON?.[rowKey];
    const overlay = this.deferredCellOverlays.get(rowKey);
    let count = Object.keys(base ?? {}).length;
    if (overlay) {
      for (const [columnKey, cell] of overlay) {
        const existed = base !== undefined && Object.prototype.hasOwnProperty.call(base, columnKey);
        if (cell === null && existed) count -= 1;
        else if (cell !== null && !existed) count += 1;
      }
    }
    this.deferredCellCountsByRow.set(rowKey, count);
    return count;
  }

  private parseDeferredCoordinate(key: string, axis: 'row' | 'column'): number {
    return parseCellMatrixCoordinate(key, axis, 'deferred');
  }

  private forEachDeferredCellInRow(row: Row, callback: (cell: CellData, column: Column) => void): void {
    const rowKey = String(row);
    const base = this.deferredJSON?.[rowKey];
    const overlay = this.deferredCellOverlays.get(rowKey);
    if (!overlay) {
      for (const columnKey of Object.keys(base ?? {})) callback(base![columnKey]!, this.parseDeferredCoordinate(columnKey, 'column'));
      return;
    }

    const addedColumns = [...overlay]
      .filter(([columnKey, cell]) => cell !== null && !(base && Object.prototype.hasOwnProperty.call(base, columnKey)))
      .map(([columnKey, cell]) => ({ column: this.parseDeferredCoordinate(columnKey, 'column'), cell: cell! }))
      .sort((left, right) => left.column - right.column);
    let addedIndex = 0;
    for (const columnKey of Object.keys(base ?? {})) {
      const baseColumn = this.parseDeferredCoordinate(columnKey, 'column');
      while (addedIndex < addedColumns.length && addedColumns[addedIndex]!.column < baseColumn) {
        const added = addedColumns[addedIndex++]!;
        callback(added.cell, added.column);
      }
      const cell = overlay.has(columnKey) ? overlay.get(columnKey) : base![columnKey];
      if (cell !== null && cell !== undefined) callback(cell, baseColumn);
    }
    while (addedIndex < addedColumns.length) {
      const added = addedColumns[addedIndex++]!;
      callback(added.cell, added.column);
    }
  }

  private getDeferredRow(row: Row): Record<string, CellData> | undefined {
    const rowKey = String(row);
    const base = this.deferredJSON?.[rowKey];
    if (!this.deferredCellOverlays.has(rowKey)) return base;
    const columns: Record<string, CellData> = {};
    this.forEachDeferredCellInRow(row, (cell, column) => { columns[String(column)] = cell; });
    return Object.keys(columns).length > 0 ? columns : undefined;
  }

  private getDeferredCell(row: Row, column: Column): CellData | undefined {
    const rowKey = String(row);
    const columnKey = String(column);
    const overlay = this.deferredCellOverlays.get(rowKey);
    if (overlay?.has(columnKey)) return overlay.get(columnKey) ?? undefined;
    return this.deferredJSON?.[rowKey]?.[columnKey];
  }

  /** Read persisted sparse cells without constructing row maps or changing storage ownership. */
  forEachWithoutHydration(callback: (cell: CellData, row: Row, column: Column) => void): void {
    const deferred = this.deferredJSON;
    if (deferred !== undefined) {
      for (const row of this.getDeferredRowCoordinates()) {
        this.forEachDeferredCellInRow(row, (cell, column) => callback(cell, row, column));
      }
      return;
    }
    for (const [row, columns] of this.rows) {
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  /** Enumerate only calculated formula inputs without materializing deferred worksheet data. */
  forEachFormula(callback: (cell: CellData & { formula: string }, row: Row, column: Column) => void): void {
    this.forEachWithoutHydration((cell, row, column) => {
      if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) callback(cell as CellData & { formula: string }, row, column);
    });
  }

  /** Enumerate persisted cells that own any formula text, without hydrating deferred JSON. */
  forEachFormulaOwner(callback: (cell: CellData, row: Row, column: Column) => void): void {
    const hasFormulaOwner = (cell: CellData): boolean => cell.formula !== undefined
      || cell.formulaMetadata?.sourceFormula !== undefined
      || (cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula');
    this.forEachWithoutHydration((cell, row, column) => {
      if (hasFormulaOwner(cell)) callback(cell, row, column);
    });
  }

  forEachInRows(rows: ReadonlySet<Row>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    if (this.deferredJSON !== undefined) {
      this.forEachInRowsWithoutHydration(rows, (_cell, row, column) => {
        const cell = this.get(row, column);
        if (cell) callback(cell, row, column);
      });
      return;
    }
    for (const row of rows) {
      const columns = this.rows.get(row);
      if (!columns) continue;
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  /** Enumerate selected sparse rows without materializing deferred worksheet cells. */
  forEachInRowsWithoutHydration(rows: ReadonlySet<Row>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    const deferred = this.deferredJSON;
    if (deferred !== undefined) {
      for (const row of rows) {
        if (!Number.isSafeInteger(row)) continue;
        this.forEachDeferredCellInRow(row, (cell, column) => {
          if (Number.isSafeInteger(column)) callback(cell, row, column);
        });
      }
      return;
    }
    for (const row of rows) {
      const columns = this.rows.get(row);
      if (!columns) continue;
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  forEachInColumns(columns: ReadonlySet<Column>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    if (this.deferredJSON !== undefined) {
      this.forEachInColumnsWithoutHydration(columns, (_cell, row, column) => {
        const cell = this.get(row, column);
        if (cell) callback(cell, row, column);
      });
      return;
    }
    for (const [row, rowCells] of this.rows) {
      for (const column of columns) {
        const cell = rowCells.get(column);
        if (cell) callback(cell, row, column);
      }
    }
  }

  *entriesInColumn(column: Column): IterableIterator<{ row: Row; cell: CellData }> {
    if (this.deferredJSON !== undefined) {
      for (const row of this.getDeferredRowCoordinates()) {
        const cell = this.get(row, column);
        if (cell) yield { row, cell };
      }
      return;
    }
    for (const [row, rowCells] of this.rows) {
      const cell = rowCells.get(column);
      if (cell) yield { row, cell };
    }
  }

  /** Read one sparse column without cloning cells or creating editable overlays. */
  *entriesInColumnWithoutHydration(column: Column): IterableIterator<{ row: Row; cell: CellData }> {
    if (this.deferredJSON !== undefined) {
      for (const row of this.getDeferredRowCoordinates()) {
        const cell = this.getDeferredCell(row, column);
        if (cell) yield { row, cell };
      }
      return;
    }
    for (const [row, rowCells] of this.rows) {
      const cell = rowCells.get(column);
      if (cell) yield { row, cell };
    }
  }

  /** Enumerate persisted cells inside a range, materializing only matching deferred cells. */
  forEachInRange(
    startRow: Row,
    endRow: Row,
    startColumn: Column,
    endColumn: Column,
    callback: (cell: CellData, row: Row, column: Column) => void,
  ): void {
    if (this.deferredJSON !== undefined) {
      this.forEachInRangeWithoutHydration(startRow, endRow, startColumn, endColumn, (_cell, row, column) => {
        const cell = this.get(row, column);
        if (cell) callback(cell, row, column);
      });
      return;
    }
    const rows = this.getSortedRowCoordinates();
    for (let index = firstCoordinateAtLeast(rows, startRow); index < rows.length; index += 1) {
      const row = rows[index]!;
      if (row > endRow) break;
      const columns = this.rows.get(row);
      if (!columns) continue;
      for (const [column, cell] of columns) {
        if (column >= startColumn && column <= endColumn) callback(cell, row, column);
      }
    }
  }

  /** Enumerate selected sparse columns without materializing deferred worksheet cells. */
  forEachInColumnsWithoutHydration(columns: ReadonlySet<Column>, callback: (cell: CellData, row: Row, column: Column) => void): void {
    if (this.deferredJSON !== undefined) {
      if (columns.size === 0) return;
      const requestedColumns = [...columns];
      const orderByColumn = columns.size > 4
        ? new Map(requestedColumns.map((column, index) => [column, index] as const))
        : undefined;
      for (const row of this.getDeferredRowCoordinates()) {
        if (!orderByColumn) {
          for (const column of requestedColumns) {
            const cell = this.getDeferredCell(row, column);
            if (cell) callback(cell, row, column);
          }
        } else {
          const selected: Array<{ cell: CellData; column: Column; requestedOrder: number }> = [];
          this.forEachDeferredCellInRow(row, (cell, column) => {
            const requestedOrder = orderByColumn.get(column);
            if (requestedOrder !== undefined) selected.push({ cell, column, requestedOrder });
          });
          selected.sort((left, right) => left.requestedOrder - right.requestedOrder);
          for (const entry of selected) callback(entry.cell, row, entry.column);
        }
      }
      return;
    }
    const requestedColumns = [...columns];
    const orderByColumn = columns.size > 4
      ? new Map(requestedColumns.map((column, index) => [column, index] as const))
      : undefined;
    for (const [row, rowCells] of this.rows) {
      if (!orderByColumn) {
        for (const column of requestedColumns) {
          const cell = rowCells.get(column);
          if (cell) callback(cell, row, column);
        }
      } else {
        const selected: Array<{ cell: CellData; column: Column; requestedOrder: number }> = [];
        for (const [column, cell] of rowCells) {
          const requestedOrder = orderByColumn.get(column);
          if (requestedOrder !== undefined) selected.push({ cell, column, requestedOrder });
        }
        selected.sort((left, right) => left.requestedOrder - right.requestedOrder);
        for (const entry of selected) callback(entry.cell, row, entry.column);
      }
    }
  }

  /** Enumerate a sparse range in row-major order without materializing deferred worksheet cells. */
  forEachInRangeWithoutHydration(
    startRow: Row,
    endRow: Row,
    startColumn: Column,
    endColumn: Column,
    callback: (cell: CellData, row: Row, column: Column) => void,
  ): void {
    const deferred = this.deferredJSON;
    if (deferred !== undefined) {
      const rows = this.getDeferredRowCoordinates();
      const rangeWidth = endColumn - startColumn;
      for (let index = firstCoordinateAtLeast(rows, startRow); index < rows.length; index += 1) {
        const row = rows[index]!;
        if (row > endRow) break;
        if (rangeWidth >= 0 && rangeWidth < 4) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            const cell = this.getDeferredCell(row, column);
            if (cell) callback(cell, row, column);
          }
          continue;
        }
        this.forEachDeferredCellInRow(row, (cell, column) => {
          if (Number.isSafeInteger(column) && column >= startColumn && column <= endColumn) callback(cell, row, column);
        });
      }
      return;
    }

    const rows = this.getSortedRowCoordinates();
    const rangeWidth = endColumn - startColumn;
    for (let index = firstCoordinateAtLeast(rows, startRow); index < rows.length; index += 1) {
      const row = rows[index]!;
      if (row > endRow) break;
      const columns = this.rows.get(row);
      if (!columns) continue;
      if (rangeWidth >= 0 && rangeWidth < 4) {
        for (let column = startColumn; column <= endColumn; column += 1) {
          const cell = columns.get(column);
          if (cell) callback(cell, row, column);
        }
        continue;
      }
      const selectedColumns = [...columns.keys()]
        .filter((column) => column >= startColumn && column <= endColumn)
        .sort((left, right) => left - right);
      for (const column of selectedColumns) callback(columns.get(column)!, row, column);
    }
  }

  private getSortedRowCoordinates(): Row[] {
    if (!this.sortedRowCoordinates) {
      this.sortedRowCoordinates = [...this.rows.keys()].sort((left, right) => left - right);
    }
    return this.sortedRowCoordinates;
  }

  clone(): CellMatrix {
    const copy = new CellMatrix();
    this.forEachWithoutHydration((cell, row, column) => copy.set(row, column, structuredClone(cell)));
    return copy;
  }

  toJSON(): Record<string, Record<string, CellData>> {
    if (this.deferredJSON !== undefined) {
      const result = structuredClone(this.deferredJSON);
      for (const [row, overlay] of this.deferredCellOverlays) {
        const columns = result[row] ?? (result[row] = {});
        for (const [column, cell] of overlay) {
          if (cell === null) delete columns[column];
          else columns[column] = structuredClone(cell);
        }
        if (Object.keys(columns).length === 0) delete result[row];
      }
      for (const columns of Object.values(result)) {
        for (const [column, cell] of Object.entries(columns)) {
          columns[column] = normalizeCellDataForStorage(cell);
        }
      }
      return result;
    }
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
        matrix.set(parseCellMatrixCoordinate(row, 'row', 'JSON import'), parseCellMatrixCoordinate(column, 'column', 'JSON import'), { ...cell });
      }
    }
    return matrix;
  }

  private hydrate(): void {
    const input = this.deferredJSON;
    if (input === undefined) return;
    const rows = this.getDeferredRowCoordinates().map((row) => [row, this.getDeferredRow(row)!] as const);
    const normalizedFontFamilies = new Map<CellData, string>();
    for (const [, columns] of rows) {
      for (const cell of Object.values(columns)) {
        const fontFamily = cell.style?.fontFamily;
        if (fontFamily === undefined) continue;
        const normalized = normalizeFontFamily(fontFamily);
        if (normalized !== fontFamily) normalizedFontFamilies.set(cell, normalized);
      }
    }
    this.deferredJSON = undefined;
    this.deferredRowCoordinates = undefined;
    this.deferredBounds = undefined;
    this.revisionCounter += this.deferredRevision;
    this.deferredRevision = 0;
    this.deferredMaterializedCells = new WeakSet<CellData>();
    this.deferredCellOverlays.clear();
    this.deferredCellCountsByRow.clear();
    for (const [row, columns] of rows) {
      for (const [column, cell] of Object.entries(columns)) {
        const fontFamily = cell.style?.fontFamily;
        const normalizedCell = fontFamily === undefined
          ? { ...cell }
          : { ...cell, style: { ...cell.style, fontFamily: normalizedFontFamilies.get(cell) ?? fontFamily } };
        this.writeNormalizedCell(Number(row), Number(column), normalizedCell, false);
      }
    }
  }

  private getDeferredBounds(): NonNullable<CellMatrix['deferredBounds']> {
    if (this.deferredBounds) return this.deferredBounds;
    let count = 0;
    let startRow = Number.POSITIVE_INFINITY;
    let endRow = Number.NEGATIVE_INFINITY;
    let startColumn = Number.POSITIVE_INFINITY;
    let endColumn = Number.NEGATIVE_INFINITY;
    for (const row of this.getDeferredRowCoordinates()) {
      this.forEachDeferredCellInRow(row, (_cell, column) => {
        count += 1;
        startRow = Math.min(startRow, row);
        endRow = Math.max(endRow, row);
        startColumn = Math.min(startColumn, column);
        endColumn = Math.max(endColumn, column);
      });
    }
    this.deferredBounds = {
      count,
      startRow: Number.isFinite(startRow) ? startRow : 0,
      endRow: Number.isFinite(endRow) ? endRow : 0,
      startColumn: Number.isFinite(startColumn) ? startColumn : 0,
      endColumn: Number.isFinite(endColumn) ? endColumn : 0,
    };
    return this.deferredBounds;
  }

  private getDeferredRowCoordinates(): Row[] {
    if (!this.deferredRowCoordinates) {
      const input = this.deferredJSON ?? {};
      const coordinates = new Set<Row>();
      let previous = Number.NEGATIVE_INFINITY;
      let ordered = true;
      for (const key of Object.keys(input)) {
        const row = this.parseDeferredCoordinate(key, 'row');
        if (row < previous) ordered = false;
        coordinates.add(row);
        previous = row;
      }
      for (const key of this.deferredCellOverlays.keys()) {
        const row = this.parseDeferredCoordinate(key, 'row');
        if (this.deferredCellCountsByRow.get(key) === 0) coordinates.delete(row);
        else coordinates.add(row);
      }
      const rows = [...coordinates];
      for (let index = 1; index < rows.length; index += 1) {
        if (rows[index]! < rows[index - 1]!) {
          ordered = false;
          break;
        }
      }
      if (!ordered) rows.sort((left, right) => left - right);
      this.deferredRowCoordinates = rows;
    }
    return this.deferredRowCoordinates;
  }

  /** 按稀疏行列索引读取范围内实际存在的单元格。 */
  getRegion(startRow: Row, endRow: Row, startColumn: Column, endColumn: Column): Array<{ row: Row; column: Column; cell: CellData }> {
    const extracted: Array<{ row: Row; column: Column; cell: CellData }> = [];
    this.forEachInRangeWithoutHydration(startRow, endRow, startColumn, endColumn, (cell, row, column) => {
      const copy = structuredClone(cell);
      const fontFamily = copy.style?.fontFamily;
      if (fontFamily !== undefined) copy.style = { ...copy.style, fontFamily: normalizeFontFamily(fontFamily) };
      extracted.push({ row, column, cell: copy });
    });
    return extracted;
  }

  /** 摘除范围内全部单元格并返回原坐标快照。 */
  extractRegion(startRow: Row, endRow: Row, startColumn: Column, endColumn: Column): Array<{ row: Row; column: Column; cell: CellData }> {
    const extracted = this.getRegion(startRow, endRow, startColumn, endColumn);
    for (const item of extracted) this.delete(item.row, item.column);
    return extracted;
  }

  placeRegion(items: ReadonlyArray<{ row: Row; column: Column; cell: CellData }>): void {
    for (const item of items) this.set(item.row, item.column, structuredClone(item.cell));
  }
}

export class CellHyperlinkMap extends Map<string, CellHyperlink> {
  private indexBuilt = false;
  private readonly hyperlinksByRow = new Map<Row, Map<Column, string>>();
  private sortedRows?: Row[];
  private readonly sortedColumnsByRow = new Map<Row, Column[]>();

  set(key: string, hyperlink: CellHyperlink): this {
    const address = this.parseAddress(key);
    const existed = super.has(key);
    super.set(key, hyperlink);
    if (this.indexBuilt && !existed) this.addAddress(address, key);
    return this;
  }

  delete(key: string): boolean {
    if (!super.has(key)) return false;
    const address = this.parseAddress(key);
    super.delete(key);
    if (this.indexBuilt) {
      const columns = this.hyperlinksByRow.get(address.row);
      columns?.delete(address.column);
      this.sortedColumnsByRow.delete(address.row);
      if (columns?.size === 0) {
        this.hyperlinksByRow.delete(address.row);
        this.sortedRows = undefined;
      }
    }
    return true;
  }

  clear(): void {
    super.clear();
    if (!this.indexBuilt) return;
    this.hyperlinksByRow.clear();
    this.sortedRows = [];
    this.sortedColumnsByRow.clear();
  }

  *entriesInRange(startRow: Row, endRow: Row, startColumn: Column, endColumn: Column): IterableIterator<{
    key: string;
    row: Row;
    column: Column;
    hyperlink: CellHyperlink;
  }> {
    if (![startRow, endRow, startColumn, endColumn].every((coordinate) => Number.isSafeInteger(coordinate) && coordinate >= 0)
      || endRow < startRow || endColumn < startColumn) {
      throw new Error('Hyperlink range is invalid');
    }
    const rowCount = endRow - startRow + 1;
    const columnCount = endColumn - startColumn + 1;
    if (rowCount <= 16 && columnCount <= 16 && rowCount * columnCount <= 16) {
      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = startColumn; column <= endColumn; column += 1) {
          const key = cellKey(row, column);
          const hyperlink = super.get(key);
          if (hyperlink) yield { key, row, column, hyperlink };
        }
      }
      return;
    }
    this.ensureIndex();
    const rows = this.sortedRows ??= [...this.hyperlinksByRow.keys()].sort((left, right) => left - right);
    for (let rowIndex = firstCoordinateAtLeast(rows, startRow); rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex]!;
      if (row > endRow) break;
      const columnsByAddress = this.hyperlinksByRow.get(row)!;
      const columns = this.sortedColumnsByRow.get(row) ?? [...columnsByAddress.keys()].sort((left, right) => left - right);
      this.sortedColumnsByRow.set(row, columns);
      for (let columnIndex = firstCoordinateAtLeast(columns, startColumn); columnIndex < columns.length; columnIndex += 1) {
        const column = columns[columnIndex]!;
        if (column > endColumn) break;
        const key = columnsByAddress.get(column)!;
        const hyperlink = super.get(key);
        if (!hyperlink) throw new Error(`Hyperlink coordinate index is dangling: ${key}`);
        yield { key, row, column, hyperlink };
      }
    }
  }

  private ensureIndex(): void {
    if (this.indexBuilt) return;
    for (const key of super.keys()) this.addAddress(this.parseAddress(key), key);
    this.indexBuilt = true;
  }

  private addAddress(address: { row: Row; column: Column }, key: string): void {
    let columns = this.hyperlinksByRow.get(address.row);
    if (!columns) {
      columns = new Map();
      this.hyperlinksByRow.set(address.row, columns);
      this.sortedRows = undefined;
    }
    if (columns.has(address.column)) throw new Error(`Hyperlink coordinate is duplicated: ${key}`);
    columns.set(address.column, key);
    this.sortedColumnsByRow.delete(address.row);
  }

  private parseAddress(key: string): { row: Row; column: Column } {
    const separator = key.indexOf(':');
    if (separator <= 0 || separator !== key.lastIndexOf(':')) throw new Error(`Hyperlink cell key is invalid: ${key}`);
    const row = Number(key.slice(0, separator));
    const column = Number(key.slice(separator + 1));
    if (cellKey(row, column) !== key) throw new Error(`Hyperlink cell key is not canonical: ${key}`);
    return { row, column };
  }
}

export class WorksheetModel {
  kind: SheetKind = 'worksheet';
  tableSheet?: TableSheetDefinition;
  ganttSheet?: GanttSheetDefinition;
  reportSheet?: ReportSheetDefinition;
  readonly cells: CellMatrix;
  /** Block-backed regions are metadata only; their bytes never enter CellMatrix. */
  private readonly dataRegionStore: SheetDataRegion[] = [];
  private dataRegionBounds = new DataRegionBoundsIndex();
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
  readonly hyperlinks = new CellHyperlinkMap();
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

  snapshot(): SheetSnapshot {
    return {
      kind: this.kind,
      id: this.id,
      name: this.name,
      rowCount: this.rowCount,
      columnCount: this.columnCount,
      cells: this.cells.toJSON(),
      dataRegions: this.dataRegions.map((region) => structuredClone(region)),
      merges: structuredClone(this.merges),
      pane: normalizeWorksheetPane(this.pane),
      pivots: structuredClone(this.pivots),
      sparklines: structuredClone(this.sparklines),
      conditionalFormats: structuredClone(this.conditionalFormats),
      dataValidations: structuredClone(this.dataValidations),
      defaultRowHeightPx: this.defaultRowHeightPx,
      defaultColumnWidthPx: this.defaultColumnWidthPx,
      rowHeightsPx: { ...this.rowHeightsPx },
      columnWidthsPx: { ...this.columnWidthsPx },
      hiddenRows: [...this.hiddenRows],
      hiddenColumns: [...this.hiddenColumns],
      tabColor: this.tabColor,
      bandedRule: this.bandedRule ? structuredClone(this.bandedRule) : undefined,
      autoFilter: this.autoFilter ? structuredClone(this.autoFilter) : undefined,
      sheetTables: structuredClone(this.sheetTables),
      sparklineGroups: structuredClone(this.sparklineGroups),
      drawings: structuredClone(this.drawings),
      drawingPayloads: Object.fromEntries([...this.drawingPayloads.entries()].map(([k, v]) => [k, structuredClone(v)])),
      drawingGroups: structuredClone(this.drawingGroups),
      snapSettings: structuredClone(this.snapSettings),
      hyperlinks: [...this.hyperlinks.entries()].map(([key, hyperlink]) => {
        const [row, column] = key.split(':').map(Number);
        return { row: row!, column: column!, hyperlink: structuredClone(hyperlink) };
      }),
      review: this.review.toSnapshot(),
      spillRanges: structuredClone(this.spillRanges),
      protectionRules: structuredClone(this.protectionRules),
      showGridlines: this.showGridlines,
      showHeaders: this.showHeaders,
      zoom: this.zoom,
      hidden: this.hidden,
      outline: this.outline ? structuredClone(this.outline) : undefined,
      tableSheet: this.tableSheet ? structuredClone(this.tableSheet) : undefined,
      ganttSheet: this.ganttSheet ? structuredClone(this.ganttSheet) : undefined,
      reportSheet: this.reportSheet ? structuredClone(this.reportSheet) : undefined,
    };
  }

  static fromSnapshot(input: SheetSnapshot): WorksheetModel {
    const paneError = worksheetPaneValidationError(input.pane);
    if (paneError !== undefined) throw new Error(`Workbook snapshot pane ${paneError} is invalid`);
    const sheet = new WorksheetModel(input.id, input.name, input.rowCount, input.columnCount);
    sheet.kind = input.kind;
    sheet.tableSheet = input.tableSheet ? structuredClone(input.tableSheet) : undefined;
    sheet.ganttSheet = input.ganttSheet ? structuredClone(input.ganttSheet) : undefined;
    sheet.reportSheet = input.reportSheet ? structuredClone(input.reportSheet) : undefined;
    sheet.cells.deferJSON(input.cells);
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
    for (const entry of input.hyperlinks) sheet.hyperlinks.set(cellKey(entry.row, entry.column), structuredClone(entry.hyperlink));
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
    return sheet;
  }

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
    this.cells.forEachWithoutHydration((cell, row, column) => copy.cells.set(row, column, structuredClone(cell)));
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
    this.cells = new CellMatrix((row, column) => this.extent.ensureCell(row, column));
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
    this.dataRegionBounds = nextBounds;
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
  hyperlinks: Array<{ row: number; column: number; hyperlink: CellHyperlink }>;
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

function definedNameStoreKey(name: string, scope: DefinedNameScope, sheetId?: SheetId): string {
  return JSON.stringify([scope, scope === 'sheet' ? sheetId : null, name.trim().toUpperCase()]);
}

export class WorkbookModel {
  readonly sheets = new Map<SheetId, WorksheetModel>();
  /** Sole canonical structured-data owner; bytes referenced by sources remain in the block store. */
  readonly dataModel = {
    externalLinks: new Map<string, import('./data-model').ExternalLinkBinding>(),
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
  private definedNamesByIdentity = new Map<string, DefinedNameModel>();
  private definedNameModelsProjection: readonly DefinedNameModel[] | undefined;
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
    return Object.fromEntries(this.definedNameModels
      .filter((entry) => entry.scope === 'workbook')
      .map((entry) => [entry.name, entry.formula]));
  }

  get definedNameModels(): readonly DefinedNameModel[] {
    this.definedNameModelsProjection ??= Object.freeze([...this.definedNamesByIdentity.values()].map((entry) => {
      const copy = structuredClone(entry);
      if (copy.anchor) Object.freeze(copy.anchor);
      return Object.freeze(copy);
    }));
    return this.definedNameModelsProjection;
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
      externalLinks: [...this.dataModel.externalLinks.values()].map(link => structuredClone(link)),
      sources: [...this.dataModel.sources.values()].map((source) => structuredClone(source)),
      tables: [...this.dataModel.tables.values()].map((table) => structuredClone(table)),
      relationships: [...this.dataModel.relationships.values()].map((relationship) => structuredClone(relationship)),
      views: [...this.dataModel.views.values()].map((view) => structuredClone(view)),
    };
  }

  setAnalysisView(view: AnalysisViewDefinition): void {
    if (view.kind !== 'analysis') throw new Error('Analysis view kind is required');
    if (!view.id.trim() || !view.name.trim()) throw new Error('Analysis view id and name are required');
    if (!this.dataModel.tables.has(view.tableId)) throw new Error(`Analysis view table not found: ${view.tableId}`);
    this.dataModel.views.set(view.id, structuredClone(view));
  }

  removeAnalysisView(viewId: string): AnalysisViewDefinition {
    const current = this.dataModel.views.get(viewId);
    if (!current || current.kind !== 'analysis') throw new Error(`Analysis view not found: ${viewId}`);
    this.dataModel.views.delete(viewId);
    return structuredClone(current) as AnalysisViewDefinition;
  }

  getAnalysisView(viewId: string): AnalysisViewDefinition | undefined {
    const current = this.dataModel.views.get(viewId);
    return current?.kind === 'analysis' ? structuredClone(current) as AnalysisViewDefinition : undefined;
  }

  listAnalysisViews(): AnalysisViewDefinition[] {
    return [...this.dataModel.views.values()]
      .filter((view): view is AnalysisViewDefinition => view.kind === 'analysis')
      .map((view) => structuredClone(view));
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
    if (sheetId) {
      const local = this.definedNamesByIdentity.get(definedNameStoreKey(name, 'sheet', sheetId));
      if (local) return structuredClone(local);
    }
    const global = this.definedNamesByIdentity.get(definedNameStoreKey(name, 'workbook'));
    return global ? structuredClone(global) : undefined;
  }

  getDefinedNameExact(name: string, scope: DefinedNameScope, sheetId?: SheetId): DefinedNameModel | undefined {
    if ((scope === 'workbook' && sheetId !== undefined) || (scope === 'sheet' && !sheetId)) return undefined;
    const exact = this.definedNamesByIdentity.get(definedNameStoreKey(name, scope, sheetId));
    return exact ? structuredClone(exact) : undefined;
  }

  listDefinedNames(sheetId?: SheetId): DefinedNameModel[] {
    return this.definedNameModels
      .filter((entry) => entry.scope === 'workbook' || entry.sheetId === sheetId)
      .map((entry) => structuredClone(entry));
  }

  setDefinedName(input: DefinedNameModel): DefinedNameModel {
    const model = normalizeDefinedNameModel(input);
    this.definedNamesByIdentity.set(definedNameStoreKey(model.name, model.scope, model.sheetId), structuredClone(model));
    this.definedNameModelsProjection = undefined;
    return structuredClone(model);
  }

  removeDefinedName(name: string, scope: DefinedNameScope = 'workbook', sheetId?: SheetId): DefinedNameModel | undefined {
    if ((scope === 'workbook' && sheetId !== undefined) || (scope === 'sheet' && !sheetId)) return undefined;
    const key = definedNameStoreKey(name, scope, sheetId);
    const previous = this.definedNamesByIdentity.get(key);
    if (previous) {
      this.definedNamesByIdentity.delete(key);
      this.definedNameModelsProjection = undefined;
    }
    return previous ? structuredClone(previous) : undefined;
  }

  replaceDefinedNames(inputs: readonly DefinedNameModel[]): void {
    const next = new Map<string, DefinedNameModel>();
    for (const input of inputs) {
      const model = normalizeDefinedNameModel(input);
      const key = definedNameStoreKey(model.name, model.scope, model.sheetId);
      if (next.has(key)) throw new Error(`Defined-name owner identity is duplicated: ${model.scope}:${model.sheetId ?? '*'}:${model.name}`);
      next.set(key, structuredClone(model));
    }
    this.definedNamesByIdentity = next;
    this.definedNameModelsProjection = undefined;
  }

  addTable(table: WorkbookTableModel): void {
    assertRecordTable(this, table);
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
    assertCanonicalWorksheetIdentities([...this.getSheets(), { id, name }]);
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

  reorderSheet(sheetId: SheetId, toIndex: number): StructuralTransformResult {
    const fromIndex = this.sheetOrder.indexOf(sheetId);
    if (fromIndex < 0) throw new Error(`Unknown sheet: ${sheetId}`);
    if (!Number.isSafeInteger(toIndex)) throw new Error('Sheet destination index must be an integer');
    const before = this.getSheets().map(({ id, name }) => ({ id, name }));
    const order = [...this.sheetOrder];
    order.splice(fromIndex, 1);
    order.splice(Math.max(0, Math.min(toIndex, this.sheetOrder.length - 1)), 0, sheetId);
    const after = order.map(id => ({ id, name: this.getSheet(id).name }));
    const plan = planWorkbookFormulaRewrite(this, this.getSheet(sheetId),
      formula => rewriteSheetLifecycleFormula(formula, { before, after, sheetId, kind: 'move' }), undefined, undefined, false);
    const effect = plan.apply();
    this.sheetOrder = order;
    return { ...effect, calculationContextEffect: { kind: 'calculation-context', action: 'sync-sheet-order' } };
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

  renameSheet(sheetId: SheetId, name: string): StructuralTransformResult {
    const source = this.getSheet(sheetId);
    const effect = planSheetIdentityTransform(this, {
      kind: 'rename',
      sourceSheetId: sheetId,
      sourceName: source.name,
      targetName: name,
    }).apply();
    if (!effect) throw new SheetIdentityTransformInvariantError('Rename plan returned no structural effect');
    return effect;
  }

  /** Apply only the identity field when history replay will apply server-owned reference deltas. */
  renameSheetIdentity(sheetId: SheetId, name: string): StructuralTransformResult | undefined {
    const sheet = this.getSheet(sheetId);
    const targetName = name.trim();
    if (!targetName) throw new SheetIdentityTransformInvariantError('Sheet rename requires a non-empty targetName');
    assertCanonicalWorksheetIdentities(this.getSheets().map((candidate) => ({
      id: candidate.id,
      name: candidate.id === sheetId ? targetName : candidate.name,
    })));
    if (sheet.name === targetName) return undefined;
    sheet.name = targetName;
    return {
      kind: 'structural-transform',
      removedCells: [],
      clearInputRanges: [],
      populateInputRanges: [],
      rewrittenFormulaOwners: [],
    };
  }

  getSheetSnapshot(sheetId: SheetId): SheetSnapshot {
    const sheet = this.getSheet(sheetId).snapshot();
    sheet.lifecycleDefinedNames = structuredClone(this.definedNameModels.filter((entry) => entry.scope === 'sheet' && entry.sheetId === sheetId));
    const printDocument = this.printDocuments.get(sheetId);
    if (printDocument) sheet.lifecyclePrintDocument = structuredClone(printDocument);
    return structuredClone(sheet);
  }

  restoreSheetSnapshot(snapshot: SheetSnapshot, index = this.sheetOrder.length): void {
    if (this.sheets.has(snapshot.id)) throw new Error(`Sheet already exists: ${snapshot.id}`);
    if (!Number.isSafeInteger(index)) throw new Error(`Invalid sheet restore index: ${index}`);
    assertCanonicalWorksheetIdentities([...this.getSheets(), snapshot]);
    const sheet = WorksheetModel.fromSnapshot(structuredClone(snapshot));
    const lifecycleNames = snapshot.lifecycleDefinedNames ?? [];
    for (const entry of lifecycleNames) {
      if (entry.scope !== 'sheet' || entry.sheetId !== sheet.id) {
        throw new Error(`Sheet restore defined-name owner does not match worksheet: ${entry.name}`);
      }
    }
    const definedNames = [...this.definedNameModels, ...lifecycleNames];
    assertCanonicalDefinedNameModels({
      sheets: [...this.getSheets(), sheet],
      definedNameModels: definedNames,
      definedNames: this.definedNames,
    });
    const printDocument = snapshot.lifecyclePrintDocument;
    if (printDocument && (printDocument.unitId !== this.unitId || printDocument.sheetId !== sheet.id)) {
      throw new Error(`Sheet restore print document owner does not match worksheet: ${sheet.id}`);
    }
    const normalizedPrintDocument = printDocument ? normalizePrintDocumentSnapshot(printDocument) : undefined;
    if (lifecycleNames.length > 0) this.replaceDefinedNames(definedNames);
    this.sheets.set(sheet.id, sheet);
    if (normalizedPrintDocument) this.printDocuments.set(sheet.id, normalizedPrintDocument);
    const bounded = Math.max(0, Math.min(index, this.sheetOrder.length));
    this.sheetOrder.splice(bounded, 0, sheet.id);
  }

  snapshot(): WorkbookSnapshot {
    return {
      schema: 'WorkbookSnapshot',
      version: 11,
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
      definedNameModels: structuredClone([...this.definedNameModels]),
      dataModel: this.getDataModel(),
      printDocuments: this.listPrintDocuments(),
      queryDefinitions: this.listQueryDefinitions(),
      cellStyleTemplates: this.listCellStyleTemplates(),
      sheets: this.getSheets().map((sheet) => sheet.snapshot()),
    };
  }

  static fromSnapshot(snapshot: WorkbookSnapshot): WorkbookModel {
    if (snapshot.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
    if (snapshot.version !== 11) throw new Error('Unsupported workbook snapshot version');
    if (snapshot.sheets.length === 0) throw new Error('Workbook snapshot must contain at least one sheet');
    assertCanonicalWorksheetIdentities(snapshot.sheets);
    assertCanonicalWorkbookOwnerIdentities(snapshot);
    for (const sheet of snapshot.sheets) {
      const paneError = worksheetPaneValidationError(sheet.pane);
      if (paneError !== undefined) throw new Error(`Workbook snapshot pane ${paneError} is invalid`);
    }
    assertCanonicalDefinedNameModels(snapshot);
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
    if (!Array.isArray(snapshot.dataModel.externalLinks)) throw new Error('Canonical external link definitions are required');
    for (const link of snapshot.dataModel.externalLinks) {
      assertExternalLinkBinding(link);
      if (link.sourceUnitId === snapshot.unitId) throw new Error('EXTERNAL_LINK_SOURCE_INVALID');
      workbook.dataModel.externalLinks.set(link.id, structuredClone(link));
    }
    for (const table of snapshot.dataModel.tables) workbook.dataModel.tables.set(table.id, structuredClone(table));
    for (const source of snapshot.dataModel.sources) workbook.addDataSource(source);
    for (const relationship of snapshot.dataModel.relationships) workbook.dataModel.relationships.set(relationship.id, structuredClone(relationship));
    for (const view of snapshot.dataModel.views) workbook.dataModel.views.set(view.id, structuredClone(view));
    for (const input of snapshot.sheets) {
      const sheet = WorksheetModel.fromSnapshot(input);
      workbook.sheets.set(sheet.id, sheet);
    }
    for (const document of snapshot.printDocuments ?? []) workbook.setPrintDocument(document);
    for (const definition of snapshot.queryDefinitions ?? []) workbook.setQueryDefinition(definition);
    for (const template of snapshot.cellStyleTemplates ?? []) workbook.setCellStyleTemplate(template);
    workbook.sheetOrder = snapshot.sheets.map((sheet) => sheet.id);
    for (const table of workbook.dataModel.tables.values()) assertRecordTable(workbook, table);
    for (const relation of workbook.dataModel.relationships.values()) if (workbook.getTable(relation.fromTableId).recordIdFieldId) assertRecordRelationship(workbook, relation);
    assertRecordCalculations(workbook);
    return workbook;
  }
}

export type { ExternalLinkBinding } from './data-model';

export { assertExternalLinkBinding } from './data-model';

export type { RecordFieldAddress, RecordFieldCalculation } from './data-model';

export { assertRecordCalculations, assertRecordFieldWrite, assertRecordTable, assertRecordRelationship, canonicalRecordFieldFormula, recordRows, resolveRecordField, RecordDomainError } from './record-domain';

export { guardRecordWorksheetWrites } from './record-domain';
