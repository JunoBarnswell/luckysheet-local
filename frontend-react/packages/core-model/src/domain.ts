import type { AutoFilterModel, CellAddress, CellData, RangeRef, Row, Column, SheetId, UnitId } from './index';
import type { PivotMemberKey } from './pivot';

/**
 * A defined name is scoped either to the workbook or to one worksheet.
 * `formula` intentionally keeps the authored formula/reference text instead
 * of compiling it to an A1 range; this is required for dynamic names and for
 * round-tripping OOXML name definitions.
 */
export type DefinedNameScope = 'workbook' | 'sheet';

export interface DefinedNameModel {
  name: string;
  formula: string;
  scope: DefinedNameScope;
  sheetId?: SheetId;
  /** Active-cell anchor captured when a relative Refers To was created. */
  anchor?: CellAddress;
  hidden?: boolean;
  comment?: string;
}

export function normalizeDefinedNameModel(input: DefinedNameModel): DefinedNameModel {
  const name = input.name.trim();
  if (!name) throw new Error('Defined name is required');
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) {
    throw new Error(`Invalid defined name: ${input.name}`);
  }
  if (input.scope === 'sheet' && !input.sheetId) {
    throw new Error(`Sheet-scoped defined name ${name} requires a sheetId`);
  }
  if (input.scope === 'workbook' && input.sheetId !== undefined) {
    throw new Error(`Workbook-scoped defined name ${name} cannot specify sheetId`);
  }
  if (input.anchor && (!input.anchor.sheetId || !Number.isSafeInteger(input.anchor.row) || input.anchor.row < 0 || !Number.isSafeInteger(input.anchor.column) || input.anchor.column < 0)) {
    throw new Error(`Defined name ${name} has an invalid anchor`);
  }
  const formula = input.formula.trim();
  if (!formula) throw new Error(`Defined name ${name} requires a formula`);
  return {
    name,
    formula,
    scope: input.scope,
    ...(input.sheetId ? { sheetId: input.sheetId } : {}),
    ...(input.anchor ? { anchor: structuredClone(input.anchor) } : {}),
    ...(input.hidden === undefined ? {} : { hidden: input.hidden }),
    ...(input.comment === undefined ? {} : { comment: input.comment }),
  };
}

/** 公式引擎错误值 — CellData.formulaValue 是唯一真相，禁止再用 error: string */
export type FormulaErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#CALC!'
  | '#BLOCKED!'
  | '#SPILL!'
  | '#PARSE!';

export interface FormulaError {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
  readonly message?: string;
  readonly position?: number;
}

export type FormulaValue = string | number | boolean | null | FormulaError;

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'object' && value !== null && 'kind' in value && (value as { kind: string }).kind === 'error';
}

export function createFormulaError(code: FormulaErrorCode, message = '', position?: number): FormulaError {
  return position === undefined ? { kind: 'error', code, message } : { kind: 'error', code, message, position };
}

/** 选区唯一类型 — anchorCell 是 Shift 扩展原点，primaryCell 是活动单元格 */
export interface SelectionSnapshot {
  unitId: UnitId;
  sheetId: SheetId;
  ranges: RangeRef[];
  primaryRangeIndex: number;
  primaryCell: { row: Row; column: Column };
  anchorCell: { row: Row; column: Column };
  phase: 'idle' | 'selected' | 'selecting' | 'editing' | 'preview';
}

export function createEmptySelection(unitId: UnitId, sheetId: SheetId): SelectionSnapshot {
  return {
    unitId,
    sheetId,
    ranges: [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    primaryRangeIndex: 0,
    primaryCell: { row: 0, column: 0 },
    anchorCell: { row: 0, column: 0 },
    phase: 'selected',
  };
}

/** Excel Sheet Table — 与 WorkbookTableModel（列存/查询结果）是两个概念 */
export interface SheetTableColumn {
  id: string;
  name: string;
  totalsFunction?: 'none' | 'sum' | 'count' | 'average' | 'min' | 'max';
}

export interface SheetTableModel {
  id: string;
  sheetId: SheetId;
  name: string;
  range: RangeRef;
  hasHeaderRow: boolean;
  hasTotalRow: boolean;
  showBandedRows: boolean;
  showBandedColumns: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showFilterButton: boolean;
  autoExpand: 'none' | 'rows' | 'columns' | 'both';
  autoFilter?: AutoFilterModel;
  columns: SheetTableColumn[];
  styleName?: string;
}

export type DrawingKind = 'image' | 'shape' | 'connector' | 'chart' | 'camera' | 'textbox' | 'form-control' | 'slicer' | 'timeline';

export interface DrawingTransform {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
}

export interface DrawingAnchor {
  kind: 'absolute' | 'one-cell' | 'two-cell';
  row?: Row;
  column?: Column;
  endRow?: Row;
  endColumn?: Column;
}

/** Canonical normalized crop fractions shared by in-cell and floating images. */
export interface ImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Supported editable picture appearance; unsupported OOXML effects remain opaque. */
export interface ImageEffects {
  brightness?: number;
  contrast?: number;
  transparency?: number;
}

export interface ImageDrawingPayload {
  kind: 'image';
  asset: import('./asset').AssetRef;
  altText?: string;
  name?: string;
  crop?: ImageCrop;
  effects?: ImageEffects;
}

/** Renderer-backed auto-shape identity shared by gallery, payload, and export. */
export type ShapeDrawingType = 'rectangle' | 'rounded-rectangle' | 'ellipse' | 'line' | 'arrow' | 'callout' | 'star';

export type ShapeDrawingCategory = 'basic-shapes' | 'lines' | 'callouts-and-stars';

/** Canonical shape gallery ownership. UI galleries must project this registry instead of redefining model identities. */
export const SHAPE_DRAWING_PRESETS: readonly { category: ShapeDrawingCategory; type: ShapeDrawingType }[] = [
  { category: 'basic-shapes', type: 'rectangle' },
  { category: 'basic-shapes', type: 'rounded-rectangle' },
  { category: 'basic-shapes', type: 'ellipse' },
  { category: 'lines', type: 'line' },
  { category: 'lines', type: 'arrow' },
  { category: 'callouts-and-stars', type: 'callout' },
  { category: 'callouts-and-stars', type: 'star' },
] as const;

export type ShapeTextDirection = 'horizontal' | 'vertical';
export type ShapeTextHorizontalAlignment = 'left' | 'center' | 'right';
export type ShapeTextVerticalAlignment = 'top' | 'middle' | 'bottom';

export interface ShapeDrawingEffects {
  shadow?: { color: string; blur: number; offsetX: number; offsetY: number; opacity: number };
  glow?: { color: string; radius: number; opacity: number };
}

export interface ShapeDrawingPayload {
  kind: 'shape';
  type: ShapeDrawingType;
  fill: string;
  stroke: string;
  strokeWidth?: number;
  text?: string;
  textColor?: string;
  fontSize?: number;
  effects?: ShapeDrawingEffects;
  textDirection?: ShapeTextDirection;
  textAlignment?: ShapeTextHorizontalAlignment;
  textVerticalAlignment?: ShapeTextVerticalAlignment;
  hyperlink?: HyperlinkTarget;
  propertyFormula?: string;
}

export type DrawingConnectorType = 'straight' | 'elbow' | 'curved';
export type DrawingConnectionPoint = 'top' | 'right' | 'bottom' | 'left' | 'center';
export type DrawingArrowhead = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval';

export interface DrawingConnectionEndpoint {
  drawingId: string;
  connectionPoint: DrawingConnectionPoint;
}

export interface DrawingConnectorRoutePoint {
  x: number;
  y: number;
}

export interface DrawingConnectorRoute {
  points: DrawingConnectorRoutePoint[];
}

/** Canonical connector semantics. Route points are derived from endpoint geometry. */
export interface ConnectorDrawingPayload {
  kind: 'connector';
  connectorType: DrawingConnectorType;
  start: DrawingConnectionEndpoint;
  end: DrawingConnectionEndpoint;
  stroke: string;
  strokeWidth?: number;
  startArrowhead: DrawingArrowhead;
  endArrowhead: DrawingArrowhead;
  route: DrawingConnectorRoute;
}

export interface DrawingGroup {
  id: string;
  sheetId: SheetId;
  memberDrawingIds: string[];
}

/** Worksheet-owned snapping policy. Pointer code must not invent a grid. */
export interface WorksheetSnapSettings {
  enabled: boolean;
  snapToGrid: boolean;
  snapToShape: boolean;
  gridSize: number;
}

export const DEFAULT_WORKSHEET_SNAP_SETTINGS: WorksheetSnapSettings = {
  enabled: true,
  snapToGrid: true,
  snapToShape: true,
  gridSize: 8,
};

export type TextBoxHorizontalAlignment = 'left' | 'center' | 'right';
export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';
export type TextBoxTextDirection = 'horizontal' | 'vertical';
export type TextBoxAutofit = 'none' | 'shrink-text' | 'resize-shape';

/** Canonical OOXML text-frame semantics shared by rendering, commands and export. */
export interface TextBoxTextFrame {
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  textColor: string;
  horizontalAlignment: TextBoxHorizontalAlignment;
  verticalAlignment: TextBoxVerticalAlignment;
  direction: TextBoxTextDirection;
  margin: { top: number; right: number; bottom: number; left: number };
  wrap: boolean;
  autofit: TextBoxAutofit;
}

export function createDefaultTextBoxTextFrame(): TextBoxTextFrame {
  return {
    fontFamily: 'Inter',
    fontSize: 14,
    bold: false,
    italic: false,
    underline: false,
    textColor: '#1f2937',
    horizontalAlignment: 'left',
    verticalAlignment: 'top',
    direction: 'horizontal',
    margin: { top: 8, right: 8, bottom: 8, left: 8 },
    wrap: true,
    autofit: 'none',
  };
}

export interface TextBoxDrawingPayload {
  kind: 'textbox';
  text: string;
  textFrame: TextBoxTextFrame;
}

export interface CameraDrawingPayload {
  kind: 'camera';
  sourceRange: RangeRef;
  refreshPolicy: 'live';
}

export type FormControlType = 'button' | 'spin-button' | 'list-box' | 'combo-box' | 'checkbox' | 'option-button' | 'group-box' | 'label' | 'scrollbar';

export interface FormControlStyle {
  fill: string;
  border: string;
  textColor: string;
  fontSize?: number;
}

export interface FormControlCellLink {
  sheetId: SheetId;
  row: Row;
  column: Column;
}

export interface FormControlAction {
  kind: 'event';
  eventId: string;
}

interface FormControlBase {
  kind: 'form-control';
  text?: string;
  enabled: boolean;
  style: FormControlStyle;
}

interface FormControlLinkedBase extends FormControlBase {
  cellLink?: FormControlCellLink;
}

export interface ButtonFormControlPayload extends FormControlBase {
  controlType: 'button';
  value: null;
  action: FormControlAction;
}

export interface SpinButtonFormControlPayload extends FormControlLinkedBase {
  controlType: 'spin-button';
  value: number;
  minValue: number;
  maxValue: number;
  step: number;
}

export interface ListBoxFormControlPayload extends FormControlLinkedBase {
  controlType: 'list-box';
  value: string | null;
  inputRange: RangeRef;
  selectionType: 'single' | 'multiple';
  selectedIndices: number[];
}

export interface ComboBoxFormControlPayload extends FormControlLinkedBase {
  controlType: 'combo-box';
  value: string | null;
  inputRange: RangeRef;
  dropDownLines: number;
}

export interface CheckboxFormControlPayload extends FormControlLinkedBase {
  controlType: 'checkbox';
  value: boolean;
}

export interface OptionButtonFormControlPayload extends FormControlLinkedBase {
  controlType: 'option-button';
  value: boolean;
  groupId?: string;
}

export interface GroupBoxFormControlPayload extends FormControlBase {
  controlType: 'group-box';
  value: null;
  groupId: string;
}

export interface LabelFormControlPayload extends FormControlBase {
  controlType: 'label';
  value: null;
}

export interface ScrollbarFormControlPayload extends FormControlLinkedBase {
  controlType: 'scrollbar';
  value: number;
  minValue: number;
  maxValue: number;
  step: number;
  pageChange: number;
}

export type FormControlDrawingPayload =
  | ButtonFormControlPayload
  | SpinButtonFormControlPayload
  | ListBoxFormControlPayload
  | ComboBoxFormControlPayload
  | CheckboxFormControlPayload
  | OptionButtonFormControlPayload
  | GroupBoxFormControlPayload
  | LabelFormControlPayload
  | ScrollbarFormControlPayload;

const isFormPayloadRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const isFiniteNumber = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value);
const isRangeRef = (value: unknown): value is RangeRef => {
  if (!isFormPayloadRecord(value)) return false;
  const startRow = value.startRow as number;
  const endRow = value.endRow as number;
  const startColumn = value.startColumn as number;
  const endColumn = value.endColumn as number;
  return typeof value.sheetId === 'string' && value.sheetId.length > 0
    && Number.isSafeInteger(startRow) && Number.isSafeInteger(endRow)
    && Number.isSafeInteger(startColumn) && Number.isSafeInteger(endColumn)
    && startRow >= 0 && endRow >= startRow
    && startColumn >= 0 && endColumn >= startColumn;
};
const isCellLink = (value: unknown): value is FormControlCellLink => {
  if (!isFormPayloadRecord(value)) return false;
  const row = value.row as number;
  const column = value.column as number;
  return typeof value.sheetId === 'string' && value.sheetId.length > 0
    && Number.isSafeInteger(row) && Number.isSafeInteger(column)
    && row >= 0 && column >= 0;
};

/** Canonical runtime validator shared by drawing storage and form-control commands. */
export function isFormControlDrawingPayload(value: unknown): value is FormControlDrawingPayload {
  if (!isFormPayloadRecord(value) || value.kind !== 'form-control' || typeof value.controlType !== 'string'
    || typeof value.enabled !== 'boolean' || !isFormPayloadRecord(value.style)
    || typeof value.style.fill !== 'string' || typeof value.style.border !== 'string' || typeof value.style.textColor !== 'string'
    || (value.style.fontSize !== undefined && (!isFiniteNumber(value.style.fontSize) || value.style.fontSize <= 0))) return false;
  if (value.text !== undefined && typeof value.text !== 'string') return false;
  if (value.cellLink !== undefined && !isCellLink(value.cellLink)) return false;
  switch (value.controlType) {
    case 'button':
      return value.value === null && isFormPayloadRecord(value.action) && value.action.kind === 'event'
        && typeof value.action.eventId === 'string' && value.action.eventId.length > 0
        && value.cellLink === undefined;
    case 'spin-button':
      return isFiniteNumber(value.value) && isFiniteNumber(value.minValue) && isFiniteNumber(value.maxValue)
        && isFiniteNumber(value.step) && (value.minValue as number) <= (value.maxValue as number) && (value.step as number) > 0
        && (value.value as number) >= (value.minValue as number) && (value.value as number) <= (value.maxValue as number);
    case 'scrollbar':
      return isFiniteNumber(value.value) && isFiniteNumber(value.minValue) && isFiniteNumber(value.maxValue)
        && isFiniteNumber(value.step) && isFiniteNumber(value.pageChange) && (value.minValue as number) <= (value.maxValue as number)
        && (value.step as number) > 0 && (value.pageChange as number) > 0 && (value.value as number) >= (value.minValue as number) && (value.value as number) <= (value.maxValue as number);
    case 'list-box': {
      if ((value.value !== null && typeof value.value !== 'string') || !isRangeRef(value.inputRange)
        || !['single', 'multiple'].includes(String(value.selectionType)) || !Array.isArray(value.selectedIndices)) return false;
      if (!value.selectedIndices.every((index) => Number.isSafeInteger(index) && index >= 0)) return false;
      return value.selectionType === 'multiple' || value.selectedIndices.length <= 1;
    }
    case 'combo-box':
      return (value.value === null || typeof value.value === 'string') && isRangeRef(value.inputRange)
        && Number.isSafeInteger(value.dropDownLines) && (value.dropDownLines as number) >= 1 && (value.dropDownLines as number) <= 100;
    case 'checkbox':
      return typeof value.value === 'boolean';
    case 'option-button':
      return typeof value.value === 'boolean' && (value.groupId === undefined || (typeof value.groupId === 'string' && value.groupId.length > 0));
    case 'group-box':
      return value.value === null && typeof value.groupId === 'string' && value.groupId.length > 0 && value.cellLink === undefined;
    case 'label':
      return value.value === null && value.cellLink === undefined;
    default:
      return false;
  }
}

export type P1ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'bubble'
  | 'treemap'
  | 'sunburst'
  | 'histogram'
  | 'pareto'
  | 'box-whisker'
  | 'waterfall'
  | 'funnel'
  | 'stock'
  | 'surface'
  | 'radar'
  | 'map'
  | 'combo';

export type ChartSubtype =
  | 'clustered'
  | 'stacked'
  | 'percent-stacked'
  | 'three-dimensional'
  | 'three-dimensional-stacked'
  | 'three-dimensional-percent-stacked'
  | 'line'
  | 'line-markers'
  | 'area'
  | 'pie'
  | 'pie-of-pie'
  | 'bar-of-pie'
  | 'doughnut'
  | 'scatter-markers'
  | 'scatter-smooth-lines'
  | 'scatter-straight-lines'
  | 'bubble'
  | 'bubble-three-dimensional'
  | 'treemap'
  | 'sunburst'
  | 'histogram'
  | 'pareto'
  | 'box-whisker'
  | 'waterfall'
  | 'funnel'
  | 'stock-high-low-close'
  | 'stock-open-high-low-close'
  | 'surface-three-dimensional'
  | 'surface-wireframe'
  | 'radar'
  | 'radar-markers'
  | 'radar-filled'
  | 'custom-combo'
  | 'filled-map';

export const CHART_SUBTYPES_BY_TYPE: Readonly<Record<P1ChartType, readonly ChartSubtype[]>> = {
  column: ['clustered', 'stacked', 'percent-stacked', 'three-dimensional', 'three-dimensional-stacked', 'three-dimensional-percent-stacked'],
  bar: ['clustered', 'stacked', 'percent-stacked', 'three-dimensional', 'three-dimensional-stacked', 'three-dimensional-percent-stacked'],
  line: ['line', 'line-markers', 'stacked', 'percent-stacked', 'three-dimensional'],
  area: ['area', 'stacked', 'percent-stacked', 'three-dimensional'],
  pie: ['pie', 'three-dimensional', 'pie-of-pie', 'bar-of-pie'],
  doughnut: ['doughnut'],
  scatter: ['scatter-markers', 'scatter-smooth-lines', 'scatter-straight-lines'],
  bubble: ['bubble', 'bubble-three-dimensional'],
  treemap: ['treemap'],
  sunburst: ['sunburst'],
  histogram: ['histogram'],
  pareto: ['pareto'],
  'box-whisker': ['box-whisker'],
  waterfall: ['waterfall'],
  funnel: ['funnel'],
  stock: ['stock-high-low-close', 'stock-open-high-low-close'],
  surface: ['surface-three-dimensional', 'surface-wireframe'],
  radar: ['radar', 'radar-markers', 'radar-filled'],
  map: ['filled-map'],
  combo: ['custom-combo'],
};

export function defaultChartSubtype(type: P1ChartType): ChartSubtype {
  return CHART_SUBTYPES_BY_TYPE[type][0]!;
}

export function isChartSubtypeForType(type: P1ChartType, subtype: ChartSubtype): boolean {
  return CHART_SUBTYPES_BY_TYPE[type].includes(subtype);
}

export function chartStackingForSubtype(subtype: ChartSubtype): 'stacked' | 'percent' | undefined {
  return subtype === 'stacked' || subtype === 'three-dimensional-stacked' ? 'stacked' : subtype === 'percent-stacked' || subtype === 'three-dimensional-percent-stacked' ? 'percent' : undefined;
}

export type ChartSeriesType = Exclude<P1ChartType, 'combo'>;
export type ChartAxisPosition = 'top' | 'bottom' | 'left' | 'right';
export type ChartAxisScale = 'linear' | 'logarithmic';
export type ChartDataLabelPosition = 'best-fit' | 'center' | 'inside-end' | 'inside-base' | 'outside-end';

export interface ChartAxisModel {
  id: string;
  position: ChartAxisPosition;
  visible?: boolean;
  title?: string;
  scale?: ChartAxisScale;
  minimum?: number;
  maximum?: number;
  majorUnit?: number;
  minorUnit?: number;
  numberFormat?: string;
  crossesAt?: number;
  majorGridlines?: ChartGridlineModel;
  minorGridlines?: ChartGridlineModel;
}

export interface ChartGridlineModel {
  visible: boolean;
  color?: string;
  width?: number;
  dash?: 'solid' | 'dash' | 'dot';
}

export interface ChartAreaStyle {
  fill?: string;
  border?: string;
  borderWidth?: number;
  borderDash?: 'solid' | 'dash' | 'dot';
}

export interface ChartMarkerModel {
  enabled: boolean;
  shape?: 'circle' | 'square' | 'diamond' | 'triangle';
  size?: number;
  fill?: string;
  border?: string;
}

export interface ChartTrendlineModel {
  type: 'linear' | 'exponential' | 'polynomial' | 'moving-average';
  order?: number;
  period?: number;
  color?: string;
  width?: number;
}

export interface ChartErrorBarsModel {
  type: 'fixed' | 'percentage' | 'standard-deviation' | 'standard-error' | 'custom';
  value?: number;
  plusRange?: RangeRef;
  minusRange?: RangeRef;
  color?: string;
  width?: number;
}

export interface ChartDataLabelsModel {
  visible: boolean;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercentage?: boolean;
  position?: ChartDataLabelPosition;
  numberFormat?: string;
}

export interface ChartSeriesModel {
  name: string;
  range: RangeRef;
  xRange?: RangeRef;
  yRange?: RangeRef;
  color?: string;
  chartType?: ChartSeriesType;
  axis?: 'primary' | 'secondary';
  smooth?: boolean;
  marker?: ChartMarkerModel;
  dataLabels?: ChartDataLabelsModel;
  trendline?: ChartTrendlineModel;
  errorBars?: ChartErrorBarsModel;
}

export type ChartAggregate = 'sum' | 'average' | 'count' | 'min' | 'max' | 'none';
export type ChartBindingArea = 'values' | 'category' | 'details' | 'color' | 'size' | 'tooltip' | 'filter';

export interface ChartFieldBinding {
  fieldId: string;
  area: ChartBindingArea;
  aggregate: ChartAggregate;
  sort?: 'asc' | 'desc';
  format?: string;
}

export type ChartBindings = Record<ChartBindingArea, ChartFieldBinding[]>;

/** The data source is a first-class chart dimension; it never changes drawing identity. */
export type ChartSource =
  | { kind: 'worksheet-ranges'; ranges: RangeRef[] }
  | { kind: 'pivot'; pivotId: string }
  | { kind: 'table'; tableId: string; bindings: ChartBindings }
  | { kind: 'report-range'; range: RangeRef; bindings: ChartBindings };

/** All chart semantics live in this value object; DrawingObject only owns placement. */
export interface ChartElementModel {
  title?: string;
  legend?: {
    visible: boolean;
    position: 'top' | 'bottom' | 'left' | 'right';
  };
  dataLabels?: ChartDataLabelsModel;
  categoryAxis?: ChartAxisModel;
  valueAxis?: ChartAxisModel;
  secondaryCategoryAxis?: ChartAxisModel;
  secondaryValueAxis?: ChartAxisModel;
  plotArea?: ChartAreaStyle;
  chartArea?: ChartAreaStyle;
  hiddenData: 'show' | 'hideRows' | 'hideColumns';
}

export interface ChartDrawingPayload {
  kind: 'chart';
  chartId: string;
  chartType: P1ChartType;
  subtype: ChartSubtype;
  source: ChartSource;
  series?: ChartSeriesModel[];
  categoryRange?: RangeRef;
  stacked?: 'none' | 'stacked' | 'percent';
  elements: ChartElementModel;
}

/** A typed member filter owned by a floating Pivot slicer. */
export interface PivotControlFilter {
  mode: 'all' | 'include' | 'exclude';
  memberKeys: PivotMemberKey[];
}

/** Date bounds owned by a floating Pivot timeline. Values are ISO-like strings. */
export interface PivotTimelinePeriod {
  start?: string;
  end?: string;
}

/** Native Excel Timeline granularity, shared by the model and OOXML codec. */
export type PivotTimelineLevel = 'years' | 'quarters' | 'months' | 'days';

/** Native Timeline cache filter mode. `unknown` represents an unfiltered cache. */
export type PivotTimelineFilterType = 'unknown' | 'dateBetween' | 'dateNotBetween';

/** Persisted visual settings for both native Pivot controls. */
export interface PivotControlStyle {
  theme: 'light' | 'dark' | 'accent';
  fill: string;
  border: string;
  textColor: string;
  accentColor: string;
  selectedFill?: string;
  fontSize?: number;
}

/** Persisted presentation and interaction settings owned by an item Slicer. */
export interface PivotSlicerSettings {
  /** Whether the object header is rendered. */
  showHeader: boolean;
  /** User-authored caption; an empty caption is invalid so the object remains discoverable. */
  caption: string;
  /** Single selection replaces the current member selection; multiple selection toggles members. */
  multiSelect: boolean;
  /** Stable member ordering owned by the document, never inferred from locale. */
  sort: 'ascending' | 'descending';
  /** Hide members that have no rows after all other controls and report filters. */
  showNoDataItems: boolean;
  /** Keep no-data members visible, but place them after members with data. */
  noDataItemsLast: boolean;
  /** Render an explicit unavailable state for visible no-data members. */
  showNoDataStyle: boolean;
  /** Number of item columns in the object viewport. */
  columnCount: number;
  /** Item row height in CSS pixels. */
  itemHeight: number;
}

/** A validated Report Connections edge, resolved against the target Pivot. */
export interface PivotControlConnection {
  pivotId: string;
  /** Canonical full Pivot source/cache identity. */
  sourceKey: string;
  /** Field identity owned by the target Pivot and source dimension. */
  fieldId: string;
}

/** Canonical floating Slicer payload. Its filter is not stored on PivotModel. */
export interface PivotSlicerDrawingPayload {
  kind: 'slicer';
  pivotId: string;
  fieldId: string;
  filter: PivotControlFilter;
  style: PivotControlStyle;
  settings: PivotSlicerSettings;
  connections?: PivotControlConnection[];
}

/** Canonical floating Timeline payload. Its period is not stored on PivotModel. */
export interface PivotTimelineDrawingPayload {
  kind: 'timeline';
  pivotId: string;
  fieldId: string;
  period: PivotTimelinePeriod;
  /** Visible grouping level of the Timeline track. */
  level: PivotTimelineLevel;
  /** Granularity of the selected window, independent from `level`. */
  selectionLevel: PivotTimelineLevel;
  showHeader: boolean;
  showSelectionLabel: boolean;
  showTimeLevel: boolean;
  showHorizontalScrollbar: boolean;
  /** The native scroll window is independent from the selected filter period. */
  scrollPosition?: string;
  /** Native cache bounds. New controls derive these from the typed date domain. */
  bounds: PivotTimelinePeriod;
  filterType: PivotTimelineFilterType;
  /** Native Timeline caption/style identity retained alongside generic colors. */
  caption?: string;
  styleName?: string;
  style: PivotControlStyle;
  connections?: PivotControlConnection[];
}

export type DrawingPayload =
  | ImageDrawingPayload
  | ShapeDrawingPayload
  | ConnectorDrawingPayload
  | TextBoxDrawingPayload
  | ChartDrawingPayload
  | CameraDrawingPayload
  | FormControlDrawingPayload
  | PivotSlicerDrawingPayload
  | PivotTimelineDrawingPayload;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

export function isDrawingConnectionPoint(value: unknown): value is DrawingConnectionPoint {
  return value === 'top' || value === 'right' || value === 'bottom' || value === 'left' || value === 'center';
}

export function isWorksheetSnapSettings(value: unknown): value is WorksheetSnapSettings {
  if (!isRecord(value)) return false;
  return typeof value.enabled === 'boolean'
    && typeof value.snapToGrid === 'boolean'
    && typeof value.snapToShape === 'boolean'
    && typeof value.gridSize === 'number'
    && Number.isFinite(value.gridSize)
    && value.gridSize > 0
    && value.gridSize <= 1024;
}

export function isDrawingGroup(value: unknown): value is DrawingGroup {
  if (!isRecord(value)) return false;
  return typeof value.id === 'string' && value.id.trim().length > 0
    && typeof value.sheetId === 'string' && value.sheetId.trim().length > 0
    && Array.isArray(value.memberDrawingIds)
    && value.memberDrawingIds.length >= 2
    && value.memberDrawingIds.every((id) => typeof id === 'string' && id.trim().length > 0);
}

export function isShapeDrawingPayload(value: unknown): value is ShapeDrawingPayload {
  if (!isRecord(value) || value.kind !== 'shape') return false;
  if (!['rectangle', 'rounded-rectangle', 'ellipse', 'line', 'arrow', 'callout', 'star'].includes(String(value.type))) return false;
  if (typeof value.fill !== 'string' || value.fill.trim().length === 0 || typeof value.stroke !== 'string' || value.stroke.trim().length === 0) return false;
  if (value.strokeWidth !== undefined && (typeof value.strokeWidth !== 'number' || !Number.isFinite(value.strokeWidth) || value.strokeWidth <= 0 || value.strokeWidth > 100)) return false;
  if (value.text !== undefined && typeof value.text !== 'string') return false;
  if (value.textColor !== undefined && typeof value.textColor !== 'string') return false;
  if (value.fontSize !== undefined && (typeof value.fontSize !== 'number' || !Number.isFinite(value.fontSize) || value.fontSize <= 0)) return false;
  if (value.textDirection !== undefined && value.textDirection !== 'horizontal' && value.textDirection !== 'vertical') return false;
  if (value.textAlignment !== undefined && !['left', 'center', 'right'].includes(String(value.textAlignment))) return false;
  if (value.textVerticalAlignment !== undefined && !['top', 'middle', 'bottom'].includes(String(value.textVerticalAlignment))) return false;
  if (value.propertyFormula !== undefined && (typeof value.propertyFormula !== 'string' || value.propertyFormula.trim().length === 0)) return false;
  if (value.effects !== undefined && !isRecord(value.effects)) return false;
  return true;
}

export function isDrawingConnectorPayload(value: unknown): value is ConnectorDrawingPayload {
  if (!isRecord(value) || value.kind !== 'connector') return false;
  const endpoint = (entry: unknown): entry is DrawingConnectionEndpoint => isRecord(entry)
    && typeof entry.drawingId === 'string' && entry.drawingId.trim().length > 0
    && isDrawingConnectionPoint(entry.connectionPoint);
  const route = value.route;
  return (value.connectorType === 'straight' || value.connectorType === 'elbow' || value.connectorType === 'curved')
    && endpoint(value.start) && endpoint(value.end)
    && typeof value.stroke === 'string' && value.stroke.trim().length > 0
    && (value.strokeWidth === undefined || (typeof value.strokeWidth === 'number' && Number.isFinite(value.strokeWidth) && value.strokeWidth > 0 && value.strokeWidth <= 100))
    && ['none', 'triangle', 'stealth', 'diamond', 'oval'].includes(String(value.startArrowhead))
    && ['none', 'triangle', 'stealth', 'diamond', 'oval'].includes(String(value.endArrowhead))
    && isRecord(route)
    && Array.isArray(route.points)
    && route.points.length >= 2
    && route.points.every((point) => isRecord(point) && typeof point.x === 'number' && Number.isFinite(point.x) && typeof point.y === 'number' && Number.isFinite(point.y));
}

function isPivotMemberKey(value: unknown): value is PivotMemberKey {
  if (!isRecord(value) || !['text', 'number', 'boolean', 'blank', 'error'].includes(String(value.type))) return false;
  if (value.type === 'blank') return value.value === null;
  if (value.type === 'text') return typeof value.value === 'string';
  if (value.type === 'number') return typeof value.value === 'number' && Number.isFinite(value.value);
  if (value.type === 'error') return typeof value.value === 'string'
    && ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CALC!', '#BLOCKED!', '#SPILL!', '#PARSE!'].includes(value.value);
  return typeof value.value === 'boolean';
}

export function isPivotControlFilter(value: unknown): value is PivotControlFilter {
  if (!isRecord(value) || !['all', 'include', 'exclude'].includes(String(value.mode))) return false;
  return Array.isArray(value.memberKeys) && value.memberKeys.every(isPivotMemberKey);
}

export function isPivotTimelinePeriod(value: unknown): value is PivotTimelinePeriod {
  if (!isRecord(value)) return false;
  return (value.start === undefined || (typeof value.start === 'string' && value.start.trim().length > 0))
    && (value.end === undefined || (typeof value.end === 'string' && value.end.trim().length > 0));
}

function isPivotTimelineLevel(value: unknown): value is PivotTimelineLevel {
  return value === 'years' || value === 'quarters' || value === 'months' || value === 'days';
}

function isPivotTimelineFilterType(value: unknown): value is PivotTimelineFilterType {
  return value === 'unknown' || value === 'dateBetween' || value === 'dateNotBetween';
}

export function isPivotControlStyle(value: unknown): value is PivotControlStyle {
  if (!isRecord(value) || !['light', 'dark', 'accent'].includes(String(value.theme))) return false;
  return typeof value.fill === 'string'
    && typeof value.border === 'string'
    && typeof value.textColor === 'string'
    && typeof value.accentColor === 'string'
    && (value.selectedFill === undefined || typeof value.selectedFill === 'string')
    && (value.fontSize === undefined || (typeof value.fontSize === 'number' && Number.isFinite(value.fontSize) && value.fontSize > 0));
}

function isPivotControlConnection(value: unknown): value is PivotControlConnection {
  if (!isRecord(value)) return false;
  return typeof value.pivotId === 'string' && value.pivotId.trim().length > 0
    && typeof value.sourceKey === 'string' && value.sourceKey.trim().length > 0
    && typeof value.fieldId === 'string' && value.fieldId.trim().length > 0;
}

function isPivotControlConnections(value: unknown): value is PivotControlConnection[] {
  return value === undefined || (Array.isArray(value) && value.every(isPivotControlConnection));
}

export function isPivotSlicerDrawingPayload(value: unknown): value is PivotSlicerDrawingPayload {
  if (!isRecord(value) || value.kind !== 'slicer') return false;
  return typeof value.pivotId === 'string'
    && value.pivotId.trim().length > 0
    && typeof value.fieldId === 'string'
    && value.fieldId.trim().length > 0
    && isPivotControlFilter(value.filter)
    && isPivotControlStyle(value.style)
    && isPivotSlicerSettings(value.settings)
    && isPivotControlConnections(value.connections);
}

export function isPivotSlicerSettings(value: unknown): value is PivotSlicerSettings {
  if (!isRecord(value)) return false;
  return typeof value.showHeader === 'boolean'
    && typeof value.caption === 'string'
    && value.caption.trim().length > 0
    && typeof value.multiSelect === 'boolean'
    && (value.sort === 'ascending' || value.sort === 'descending')
    && typeof value.showNoDataItems === 'boolean'
    && typeof value.noDataItemsLast === 'boolean'
    && typeof value.showNoDataStyle === 'boolean'
    && typeof value.columnCount === 'number' && Number.isSafeInteger(value.columnCount) && value.columnCount >= 1 && value.columnCount <= 32
    && typeof value.itemHeight === 'number' && Number.isFinite(value.itemHeight) && value.itemHeight >= 16 && value.itemHeight <= 96;
}

export function isPivotTimelineDrawingPayload(value: unknown): value is PivotTimelineDrawingPayload {
  if (!isRecord(value) || value.kind !== 'timeline') return false;
  return typeof value.pivotId === 'string'
    && value.pivotId.trim().length > 0
    && typeof value.fieldId === 'string'
    && value.fieldId.trim().length > 0
    && isPivotTimelinePeriod(value.period)
    && isPivotTimelineLevel(value.level)
    && isPivotTimelineLevel(value.selectionLevel)
    && typeof value.showHeader === 'boolean'
    && typeof value.showSelectionLabel === 'boolean'
    && typeof value.showTimeLevel === 'boolean'
    && typeof value.showHorizontalScrollbar === 'boolean'
    && (value.scrollPosition === undefined || (typeof value.scrollPosition === 'string' && value.scrollPosition.trim().length > 0))
    && isPivotTimelinePeriod(value.bounds)
    && isPivotTimelineFilterType(value.filterType)
    && (value.caption === undefined || typeof value.caption === 'string')
    && (value.styleName === undefined || (typeof value.styleName === 'string' && value.styleName.trim().length > 0))
    && isPivotControlStyle(value.style)
    && isPivotControlConnections(value.connections);
}

/** 浮动对象唯一 bounds/z-order 入口 */
export interface DrawingObject {
  id: string;
  sheetId: SheetId;
  kind: DrawingKind;
  name?: string;
  /** Persisted visibility state; omitted means visible for legacy snapshots. */
  visible?: boolean;
  anchor: DrawingAnchor;
  transform: DrawingTransform;
  zIndex: number;
  payloadId: string;
}

export type HyperlinkTarget =
  | { kind: 'url'; url: string }
  | { kind: 'email'; address: string; subject?: string }
  | { kind: 'sheet'; sheetId: SheetId; address?: string; row?: Row; column?: Column }
  | { kind: 'name'; name: string };

export interface CellHyperlink {
  id: string;
  target: HyperlinkTarget;
  tooltip?: string;
}

export interface OutlineGroup {
  id: string;
  axis: 'row' | 'column';
  start: number;
  end: number;
  level: number;
  collapsed: boolean;
}

export interface OutlineModel {
  groups: OutlineGroup[];
}

export interface SparklineGroup {
  id: string;
  sheetId: SheetId;
  type: 'line' | 'column' | 'win-loss';
  sparklineIds: string[];
  showAxis?: boolean;
  showMarkers?: boolean;
}

/** 单格批注 — 无线程 */
export interface CellNote {
  id: string;
  author: string;
  text: string;
  createdAt: string;
  visible: boolean;
}

/** 线程评论 — 与 Note 分离 */
export interface CommentThread {
  id: string;
  sheetId: SheetId;
  row: Row;
  column: Column;
  author: string;
  text: string;
  createdAt: string;
  mentions?: string[];
  replies: CommentReply[];
  resolved?: boolean;
  resolvedAt?: string;
}

export interface CommentReply {
  id: string;
  author: string;
  text: string;
  createdAt: string;
  mentions?: string[];
}

export type SpillState = 'ok' | 'blocked' | 'spill-error';

/** 动态数组溢出区 — 子格不存公式 */
export interface SpillRange {
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
  range: RangeRef;
  values: FormulaValue[][];
  state: SpillState;
}

export type ProtectionScope = 'workbook' | 'sheet' | 'range';

export interface ProtectionAllow {
  selectLocked?: boolean;
  selectUnlocked?: boolean;
  formatCells?: boolean;
  insertRows?: boolean;
  insertColumns?: boolean;
  deleteRows?: boolean;
  deleteColumns?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  editObjects?: boolean;
}

export interface ProtectionRule {
  id: string;
  scope: ProtectionScope;
  sheetId?: SheetId;
  range?: RangeRef;
  passwordHash?: string;
  locked: boolean;
  allow: ProtectionAllow;
}

export type StructuralOpKind =
  | 'insert-rows'
  | 'delete-rows'
  | 'insert-columns'
  | 'delete-columns'
  | 'cell-shift'
  | 'move-range'
;

export interface CellShiftSpec {
  sheetId: SheetId;
  range: RangeRef;
  operation: 'insert' | 'delete';
  axis: 'row' | 'column';
}

export interface StructuralTransformParams {
  kind: StructuralOpKind;
  sheetId: SheetId;
  at?: Row | Column;
  count?: number;
  sourceRange?: RangeRef;
  targetOrigin?: { row: Row; column: Column };
  operation?: CellShiftSpec['operation'];
  axis?: CellShiftSpec['axis'];
}
