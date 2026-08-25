/** XLSX 兼容级别 */
import type { PivotStyleOptions } from '@react-sheets/core-model';

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

export type NativePivotScalar = string | number | boolean | null;

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
  sharedItems?: Array<string | number | boolean | null>;
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
  collapsedItemIndexes?: number[];
}

export interface NativePivotDataField {
  field: number;
  name?: string;
  subtotal?: string;
  showDataAs?: string;
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
  showRowGrandTotals?: boolean;
  showColumnGrandTotals?: boolean;
  showSubtotals?: boolean;
  repeatLabels?: boolean;
  compactData?: boolean;
  styleName?: string;
  styleOptions?: PivotStyleOptions;
  showButtons?: boolean;
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
  connectedPivotIds?: string[];
  selection?: { start?: string; end?: string };
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
