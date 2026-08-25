/** XLSX 兼容级别 */
import type { PivotErrorValue, PivotStyleOptions, PivotTimelineFilterType, PivotTimelineLevel, PivotTimelinePeriod } from '@react-sheets/core-model';

export type CompatibilityLevel = 'A' | 'B' | 'C';
export type XlsxCompatibilityMode = 'strict' | 'balanced' | 'best-effort';
export const OOXML_CODEC_REVISION = 4 as const;

export type ExcelDocumentFormat =
  | { family: 'ooxml'; profile: 'transitional' | 'strict'; variant: 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'xlam' }
  | { family: 'xlsb'; variant: 'xlsb' }
  | { family: 'biff'; variant: 'xls' }
  | { family: 'text'; variant: 'csv' | 'txt' }
  | { family: 'ods'; variant: 'ods' };

/** Excel 日期系统 */
export type DateSystem = '1900' | '1904';

export interface XlsxImportOptions {
  compatibilityTarget: CompatibilityLevel;
  compatibilityMode?: XlsxCompatibilityMode;
  dateSystem?: DateSystem;
  preserveMacros?: boolean;
  /** Reject packages above the parser safety limits instead of attempting to inflate them. */
  limits?: Partial<XlsxZipLimits>;
}

export interface XlsxExportOptions {
  compatibilityTarget: CompatibilityLevel;
  dateSystem?: DateSystem;
  includeCachedValues?: boolean;
  preserveMacros?: boolean;
}

/** Limits are deliberately finite so an XLSX cannot be used as a zip bomb. */
export interface XlsxZipLimits {
  maxArchiveBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxUncompressedBytes: number;
  maxCompressionRatio: number;
}

export const DEFAULT_XLSX_ZIP_LIMITS: Readonly<XlsxZipLimits> = {
  maxArchiveBytes: 200 * 1024 * 1024,
  maxEntries: 20_000,
  maxEntryBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 500 * 1024 * 1024,
  maxCompressionRatio: 1_000,
};

export interface XlsxRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * The bytes which were supplied by the user when a workbook was imported.
 *
 * This is deliberately not part of `WorkbookSnapshot` or an operation
 * envelope.  It is a local exchange artifact that allows the package writer
 * to preserve OOXML parts which the editable model does not understand yet
 * (for example VBA, custom XML, and native Pivot parts).
 */
export interface NativePackageState {
  schema: 'NativePackageState';
  format: ExcelDocumentFormat;
  fileName: string;
  sourceBytes: ArrayBuffer;
  checksum: string;
  dateSystem: DateSystem;
  detectedFeatures: string[];
  packageGraph: OpcPackageGraph;
  ownership: FeatureOwnershipResult[];
  codecRevision: number;
  compatibility: CompatibilityReport;
}

export type NativePivotSource =
  | { kind: 'worksheet-range'; sheetName: string; sheetPart?: string; ref: string }
  | { kind: 'table'; tableName: string; sheetName?: string; sheetPart?: string };

export type NativePivotScalar = string | number | boolean | null | PivotErrorValue;

export interface NativePivotFieldRange {
  groupBy?: string;
  start?: NativePivotScalar;
  end?: NativePivotScalar;
  interval?: number;
  autoStart?: boolean;
  autoEnd?: boolean;
}

export interface NativePivotFieldGroup {
  base?: number;
  parent?: number;
  range?: NativePivotFieldRange;
  discreteIndexes?: number[];
  groupItems?: NativePivotScalar[];
}

export interface NativePivotCacheField {
  index: number;
  name: string;
  dataType?: 'string' | 'number' | 'date' | 'boolean' | 'error' | 'mixed';
  sharedItems?: NativePivotScalar[];
  fieldGroup?: NativePivotFieldGroup;
}

export interface NativePivotCacheDefinition {
  cacheId: number;
  part: string;
  recordsPart?: string;
  source: NativePivotSource | { kind: 'unsupported'; reason: string };
  fields: NativePivotCacheField[];
  recordCount?: number;
  refreshOnLoad?: boolean;
  refreshOnSave?: boolean;
  saveData?: boolean;
  enableRefresh?: boolean;
}

export interface NativePivotTableField {
  index: number;
  axis?: 'row' | 'column' | 'page' | 'data';
  compact?: boolean;
  outline?: boolean;
  sortType?: 'manual' | 'ascending' | 'descending';
  nonAutoSortDefault?: boolean;
  autoSortScope?: NativePivotAutoSortScope;
  collapsedItemIndexes?: number[];
  /** Native pivotField/items entries with h="1"; mapped to a canonical manual filter. */
  hiddenItemIndexes?: number[];
  subtotal?: { mode: 'automatic' | 'none' | 'custom'; functions?: string[] };
}

export interface NativePivotAutoSortScope {
  dataOnly?: boolean;
  labelOnly?: boolean;
  outline?: boolean;
  fieldPosition?: number;
  attributes: Record<string, string>;
  references: Array<{
    field: number;
    selected?: boolean;
    itemIndexes?: number[];
  }>;
}

/** A typed projection plus the original attributes for lossless native replay. */
export interface NativePivotFilter {
  field: number;
  type: string;
  measureField?: number;
  secondMeasureField?: number;
  evalOrder?: number;
  id?: number;
  stringValue1?: string;
  stringValue2?: string;
  value1?: NativePivotScalar;
  value2?: NativePivotScalar;
  wholeDay?: boolean;
  top?: boolean;
  percent?: boolean;
  attributes: Record<string, string>;
}

export interface NativePivotDataField {
  field: number;
  name?: string;
  subtotal?: string;
  showDataAs?: string;
  /** Native dataField base coordinate used by Difference/Running Total. */
  baseField?: number;
  /** Native shared-item index, or the explicit previous/next sentinel. */
  baseItem?: number;
  /** Canonical Excel format code represented by OOXML dataField@numFmtId. */
  numberFormat?: string;
}

export interface NativePivotTableDefinition {
  name: string;
  part: string;
  sheetPart: string;
  relationshipId: string;
  cacheId: number;
  locationRef?: string;
  fields: NativePivotTableField[];
  rowFields: number[];
  columnFields: number[];
  pageFields: number[];
  dataFields: NativePivotDataField[];
  pivotFilters?: NativePivotFilter[];
  showRowGrandTotals?: boolean;
  showColumnGrandTotals?: boolean;
  subtotalLocation?: 'top' | 'bottom' | 'off';
  repeatLabels?: boolean;
  compactData?: boolean;
  multipleFieldFilters?: boolean;
  styleName?: string;
  styleOptions?: PivotStyleOptions;
  showButtons?: boolean;
  showFieldHeaders?: boolean;
  fillEmptyCells?: boolean;
  emptyCellText?: string;
  showErrorValues?: boolean;
  errorCellText?: string;
  preserveFormatting?: boolean;
  pivotId?: string;
}

export interface NativePivotControlDefinition {
  kind: 'slicer' | 'timeline';
  id: string;
  name: string;
  sheetPart: string;
  part: string;
  cachePart: string;
  cacheName: string;
  relationshipId: string;
  cacheRelationshipId: string;
  drawingPart?: string;
  drawingRelationshipId?: string;
  drawingAnchor?: { row: number; column: number };
  pivotId?: string;
  fieldId?: string;
  fieldIndex?: number;
  pivotCacheId?: number;
  /** Native OOXML pivotTable references; converted to typed model connections on import. */
  connectionPivotIds?: string[];
  selection?: { start?: string; end?: string };
  level?: PivotTimelineLevel;
  selectionLevel?: PivotTimelineLevel;
  showHeader?: boolean;
  showSelectionLabel?: boolean;
  showTimeLevel?: boolean;
  showHorizontalScrollbar?: boolean;
  scrollPosition?: string;
  bounds?: PivotTimelinePeriod;
  filterType?: PivotTimelineFilterType;
  selectedItemIndexes?: number[];
  styleName?: string;
  caption?: string;
  valid: boolean;
  reason?: string;
}

/**
 * A constrained, validated view of the native Pivot OOXML relationship graph.
 * Unknown XML is kept in the source package, never copied into the canonical
 * workbook model or collaboration payloads.
 */
export interface NativePivotGraph {
  schema: 'NativePivotGraph';
  caches: NativePivotCacheDefinition[];
  tables: NativePivotTableDefinition[];
  controls?: NativePivotControlDefinition[];
}

/**
 * The native package update is deliberately kept outside WorkbookSnapshot.
 * It carries only reachable OOXML parts/relationships and derived display
 * cells needed by Excel; calculation/result trees never cross this boundary.
 */
export interface NativePivotPackageUpdate {
  graph: NativePivotGraph;
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  displayCellsBySheetPart: Record<string, Record<string, Record<string, import('@react-sheets/core-model').CellData>>>;
}

/**
 * The original OOXML package accompanying an imported snapshot.
 *
 * `parts` contains every part, including binary and feature parts this package
 * does not edit.  Export overlays the editable core parts on this package so
 * chart/pivot/VBA/custom XML bytes are not silently discarded.
 */
export interface OpcPackageGraph {
  schema: 'OpcPackageGraph';
  workbookPart: string;
  parts: Record<string, Uint8Array>;
  opaqueParts: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  sheetPartById: Record<string, string>;
  contentTypesXml?: Uint8Array;
  dateSystem: DateSystem;
  format: ExcelDocumentFormat;
  profile: 'transitional' | 'strict';
  nativePivotGraph?: NativePivotGraph;
}

export interface FeatureOwnershipResult {
  feature: string;
  scope: string;
  read: 'full' | 'partial' | 'none';
  edit: 'full' | 'partial' | 'none';
  write: 'full' | 'partial' | 'none';
  preserve: 'full' | 'partial' | 'none';
  ownership: 'editable-owned' | 'preserved-owned' | 'mixed-owned';
}

export type CompatibilityIssueSeverity = 'error' | 'warning' | 'info';

export interface CompatibilityIssue {
  level: CompatibilityLevel;
  severity: CompatibilityIssueSeverity;
  feature: string;
  location?: string;
  message: string;
  preserved: boolean;
  status: 'editable' | 'preserved-only' | 'unsupported';
  projection?: 'native' | 'projected' | 'preserved' | 'unsupported';
  reason: string;
}

export interface CompatibilityReport {
  schema: 'CompatibilityReport';
  fileName: string;
  importLevel: CompatibilityLevel;
  exportLevel: CompatibilityLevel;
  dateSystem: DateSystem;
  issues: CompatibilityIssue[];
  summary: {
    editableFeatures: number;
    preservedOnly: number;
    unsupported: number;
  };
}

export interface XlsxWorkbookPayload {
  name: string;
  sheetCount: number;
  dateSystem: DateSystem;
  compatibilityLevel: CompatibilityLevel;
}

export interface XlsxImportResult {
  payload: XlsxWorkbookPayload;
  report: CompatibilityReport;
  snapshot: import('@react-sheets/core-model').WorkbookSnapshot;
  nativePackage: NativePackageState;
  taskId?: string;
}

export interface XlsxExportResult {
  taskId: string;
  report: CompatibilityReport;
  buffer: ArrayBuffer;
  fileName: string;
  nativePackage?: NativePackageState;
}
