import type { AutoFilterModel, CellData, RangeRef, Row, Column, SheetId, UnitId } from './index';
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
  const formula = input.formula.trim();
  if (!formula) throw new Error(`Defined name ${name} requires a formula`);
  return {
    name,
    formula,
    scope: input.scope,
    ...(input.sheetId ? { sheetId: input.sheetId } : {}),
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
  | '#PARSE!'
  | '#CYCLE!';

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

/** 编辑会话 — Cancel 只恢复 original，不写模型 */
export interface EditSession {
  sheetId: SheetId;
  row: Row;
  column: Column;
  originalValue: CellData | null;
  originalFormula?: string;
  originalSelection: SelectionSnapshot;
  currentDraft: string;
  referenceMode: boolean;
  isDirty: boolean;
  source: 'cell' | 'formula-bar';
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

export type DrawingKind = 'image' | 'shape' | 'chart' | 'data-chart' | 'camera' | 'textbox' | 'form-control' | 'slicer' | 'timeline';

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
  src: string;
  altText?: string;
  name?: string;
  crop?: ImageCrop;
  effects?: ImageEffects;
}

export interface ShapeDrawingPayload {
  kind: 'shape';
  type: 'rectangle' | 'rounded-rectangle' | 'ellipse' | 'line' | 'arrow' | 'callout' | 'star';
  fill: string;
  stroke: string;
  strokeWidth?: number;
  text?: string;
  textColor?: string;
  fontSize?: number;
}

export interface TextBoxDrawingPayload {
  kind: 'textbox';
  text: string;
  textColor?: string;
  fontSize?: number;
}

export type DataChartAggregate = 'sum' | 'average' | 'count' | 'min' | 'max' | 'none';

export type DataChartPlotType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'treemap' | 'funnel';
export type DataChartBindingArea = 'values' | 'category' | 'details' | 'color' | 'size' | 'tooltip' | 'filter';
export type DataChartSource =
  | { kind: 'table'; tableId: string }
  | { kind: 'report-sheet'; range: RangeRef };

export interface DataChartFieldBinding {
  fieldId: string;
  area: DataChartBindingArea;
  aggregate: DataChartAggregate;
  sort?: 'asc' | 'desc';
  format?: string;
}

export interface DataChartInspectorModel {
  title?: string;
  legendPosition: 'top' | 'bottom' | 'left' | 'right' | 'none';
  showDataLabels: boolean;
  showHiddenData: boolean;
  chartArea: { fill: string; border: string; borderWidth: number };
  plotArea: { fill: string; border?: string };
  axis: {
    categoryTitle?: string;
    valueTitle?: string;
    showGridlines: boolean;
  };
}

export interface DataChartDrawingPayload {
  kind: 'data-chart';
  source: DataChartSource;
  plotType: DataChartPlotType;
  bindings: Record<DataChartBindingArea, DataChartFieldBinding[]>;
  inspector: DataChartInspectorModel;
}

export interface CameraDrawingPayload {
  kind: 'camera';
  sourceRange: RangeRef;
  refreshPolicy: 'live';
}

export type FormControlType = 'button' | 'spin-button' | 'list-box' | 'combo-box' | 'checkbox' | 'option-button' | 'group-box' | 'label' | 'scrollbar';

export interface FormControlDrawingPayload {
  kind: 'form-control';
  controlType: FormControlType;
  text?: string;
  cellLink?: { sheetId: SheetId; row: Row; column: Column };
  inputRange?: RangeRef;
  value: string | number | boolean | null;
  enabled: boolean;
  style: {
    fill: string;
    border: string;
    textColor: string;
    fontSize?: number;
  };
}

export type P1ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'combo';

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
  pivotId?: string;
  sourceRanges: RangeRef[];
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

/** Canonical floating Slicer payload. Its filter is not stored on PivotModel. */
export interface PivotSlicerDrawingPayload {
  kind: 'slicer';
  pivotId: string;
  fieldId: string;
  filter: PivotControlFilter;
  style: PivotControlStyle;
  connectedPivotIds?: string[];
}

/** Canonical floating Timeline payload. Its period is not stored on PivotModel. */
export interface PivotTimelineDrawingPayload {
  kind: 'timeline';
  pivotId: string;
  fieldId: string;
  period: PivotTimelinePeriod;
  style: PivotControlStyle;
  connectedPivotIds?: string[];
}

export type DrawingPayload =
  | ImageDrawingPayload
  | ShapeDrawingPayload
  | TextBoxDrawingPayload
  | ChartDrawingPayload
  | DataChartDrawingPayload
  | CameraDrawingPayload
  | FormControlDrawingPayload
  | PivotSlicerDrawingPayload
  | PivotTimelineDrawingPayload;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isPivotMemberKey(value: unknown): value is PivotMemberKey {
  if (!isRecord(value) || !['text', 'number', 'boolean', 'blank'].includes(String(value.type))) return false;
  if (value.type === 'blank') return value.value === null;
  if (value.type === 'text') return typeof value.value === 'string';
  if (value.type === 'number') return typeof value.value === 'number' && Number.isFinite(value.value);
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

export function isPivotControlStyle(value: unknown): value is PivotControlStyle {
  if (!isRecord(value) || !['light', 'dark', 'accent'].includes(String(value.theme))) return false;
  return typeof value.fill === 'string'
    && typeof value.border === 'string'
    && typeof value.textColor === 'string'
    && typeof value.accentColor === 'string'
    && (value.selectedFill === undefined || typeof value.selectedFill === 'string')
    && (value.fontSize === undefined || (typeof value.fontSize === 'number' && Number.isFinite(value.fontSize) && value.fontSize > 0));
}

function isConnectedPivotIds(value: unknown): value is string[] {
  return value === undefined || (Array.isArray(value)
    && value.every((entry) => typeof entry === 'string' && entry.trim().length > 0));
}

export function isPivotSlicerDrawingPayload(value: unknown): value is PivotSlicerDrawingPayload {
  if (!isRecord(value) || value.kind !== 'slicer') return false;
  return typeof value.pivotId === 'string'
    && value.pivotId.trim().length > 0
    && typeof value.fieldId === 'string'
    && value.fieldId.trim().length > 0
    && isPivotControlFilter(value.filter)
    && isPivotControlStyle(value.style)
    && isConnectedPivotIds(value.connectedPivotIds);
}

export function isPivotTimelineDrawingPayload(value: unknown): value is PivotTimelineDrawingPayload {
  if (!isRecord(value) || value.kind !== 'timeline') return false;
  return typeof value.pivotId === 'string'
    && value.pivotId.trim().length > 0
    && typeof value.fieldId === 'string'
    && value.fieldId.trim().length > 0
    && isPivotTimelinePeriod(value.period)
    && isPivotControlStyle(value.style)
    && isConnectedPivotIds(value.connectedPivotIds);
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
  /** 兼容 M10 权限检查的动作白名单 */
  allowedActions?: string[];
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
