/** XLSX 兼容级别 */
import type { PivotErrorValue, PivotStyleOptions, PivotTimelineFilterType, PivotTimelineLevel, PivotTimelinePeriod } from '@react-sheets/core-model';
import type { XmlNode } from './xml';

export type CompatibilityLevel = 'A' | 'B' | 'C';
export type NativeCompatibilityMode = 'strict' | 'balanced' | 'best-effort';
export const NATIVE_DOCUMENT_CODEC_REVISION = 1 as const;

export type NativeDocumentFormat =
  | { family: 'ooxml'; profile: 'transitional' | 'strict'; variant: 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'xlam' }
  | { family: 'xlsb'; variant: 'xlsb' }
  | { family: 'biff'; variant: 'xls' | 'xlt' | 'xla' | 'biff5' | 'xlw' }
  | { family: 'xmlss'; variant: 'xml' }
  | { family: 'text'; variant: 'csv' | 'txt' | 'prn' | 'dif' | 'sylk' }
  | { family: 'ods'; variant: 'ods' }
  | { family: 'sjs'; variant: 'sjs' }
  | { family: 'ssjson'; variant: 'ssjson' }
  | { family: 'dbf'; variant: 'dbf' }
  | { family: 'works'; variant: 'xlr' }
  | { family: 'web'; variant: 'html' | 'mht' }
  | { family: 'presentation'; variant: 'pdf' | 'xps' };

/** Excel 日期系统 */
export type DateSystem = '1900' | '1904';

export interface NativeDocumentImportOptions {
  compatibilityTarget: CompatibilityLevel;
  compatibilityMode?: NativeCompatibilityMode;
  dateSystem?: DateSystem;
  preserveMacros?: boolean;
  /** Reject documents above the parser safety limits instead of attempting to materialize them. */
  limits?: Partial<NativeDocumentResourceLimits>;
}

export interface NativeDocumentExportOptions {
  compatibilityTarget: CompatibilityLevel;
  dateSystem?: DateSystem;
  includeCachedValues?: boolean;
  preserveMacros?: boolean;
  /** Same finite limits apply while serializing a native document. */
  limits?: Partial<NativeDocumentResourceLimits>;
  /** Authoritative asset bytes resolved before the synchronous OOXML writer runs. */
  assetBytes?: Record<string, Uint8Array>;
}

/**
 * Finite limits shared by every native codec.  ZIP/OPC limits are the first
 * five fields; CFB/BIFF, XML and worksheet limits prevent a hostile native
 * document from moving the work to an unbounded allocation.
 */
export interface NativeDocumentResourceLimits {
  maxArchiveBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxUncompressedBytes: number;
  maxCompressionRatio: number;
  maxCfbStreams: number;
  maxStreamBytes: number;
  maxRecordCount: number;
  maxXmlDepth: number;
  maxXmlBytes: number;
  maxCells: number;
}

export const DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS: Readonly<NativeDocumentResourceLimits> = {
  maxArchiveBytes: 200 * 1024 * 1024,
  maxEntries: 20_000,
  maxEntryBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 500 * 1024 * 1024,
  maxCompressionRatio: 1_000,
  maxCfbStreams: 10_000,
  maxStreamBytes: 500 * 1024 * 1024,
  maxRecordCount: 2_000_000,
  maxXmlDepth: 256,
  maxXmlBytes: 100 * 1024 * 1024,
  maxCells: 10_000_000,
};

export interface NativeRelationship {
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
export interface NativeDocumentArtifact {
  schema: 'NativeDocumentArtifact';
  format: NativeDocumentFormat;
  fileName: string;
  sourceBytes: ArrayBuffer;
  checksum: string;
  securityEnvelope?: NativeSecurityEnvelope;
  dateSystem: DateSystem;
  detectedFeatures: string[];
  nativeGraph: NativeGraph;
  /** Stable projection identity used to prove an untouched Save can return source bytes. */
  sourceSnapshotHash?: string;
  ownership: FeatureOwnershipResult[];
  codecRevision: number;
  compatibility: CompatibilityReport;
}

export interface NativeSecurityEnvelope {
  kind: 'none' | 'office-standard' | 'office-agile' | 'cfb-opaque' | 'unsupported';
  encrypted: boolean;
  passwordRequired: boolean;
  signatureInvalidated?: boolean;
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

/** Native SpreadsheetML dataField base-item identity.
 *
 * Excel stores the Previous/Next choices as reserved unsigned integers, not
 * as cache shared-item indexes.  Keeping them typed at the exchange boundary
 * prevents a serializer from accidentally treating either sentinel as a
 * real member index.
 */
export type NativePivotBaseItem = number | 'previous' | 'next';

export interface NativePivotDataField {
  field: number;
  name?: string;
  subtotal?: string;
  showDataAs?: string;
  /** Native cache-field index used by custom Show Values As calculations. */
  baseField?: number;
  /** Shared-item index or Excel's Previous/Next sentinel. */
  baseItem?: NativePivotBaseItem;
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
  relationships: Record<string, NativeRelationship[]>;
  displayCellsBySheetPart: Record<string, Record<string, Record<string, import('@react-sheets/core-model').CellData>>>;
}

/**
 * The original OOXML package accompanying an imported snapshot.
 *
 * `parts` contains every part, including binary and feature parts this package
 * does not edit.  Export overlays the editable core parts on this package so
 * chart/pivot/VBA/custom XML bytes are not silently discarded.
 */
export type NativeGraph =
  | { kind: 'opc'; package: OpcPackageGraph }
  | { kind: 'text'; dialect: TextDialectGraph }
  | { kind: 'xml'; root: XmlDocumentGraph }
  | { kind: 'ods'; package: OdsPackageGraph }
  | { kind: 'sjs'; package: SpreadSjsGraph }
  | { kind: 'ssjson'; document: SpreadSsjsonGraph }
  | { kind: 'biff'; container: BinaryRecordGraph }
  | { kind: 'xlsb'; container: BinaryRecordGraph }
  | { kind: 'dbf'; table: DbfTableGraph };

export interface TextDialectGraph {
  encoding: 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252';
  bom: boolean;
  delimiter: ',' | '\t' | ';' | ' ';
  rowDelimiter: '\r\n' | '\n' | '\r';
  quote: 'double' | 'none';
  variant: 'csv' | 'txt' | 'prn' | 'dif' | 'sylk';
}

export interface XmlDocumentGraph {
  namespace: string;
  root: string;
  encoding?: string;
  parsed?: XmlNode;
}

export interface OdsPackageGraph {
  parts: Record<string, Uint8Array>;
  mimetype: string;
  contentPart: string;
  contentTree?: XmlNode;
}

export interface SpreadSjsGraph {
  parts: Record<string, Uint8Array>;
  workbookPart: string;
  unknownParts: Record<string, Uint8Array>;
  unknownFields: Record<string, unknown>;
}

export interface SpreadSsjsonGraph {
  unknownFields: Record<string, unknown>;
}

export interface BinaryRecordGraph {
  container: 'cfb' | 'biff12';
  records: Array<{ type: number; offset: number; bytes: Uint8Array }>;
  opaque: Uint8Array;
}

export interface DbfTableGraph {
  version: number;
  headerLength: number;
  recordLength: number;
  fields: Array<{ name: string; type: string; length: number; decimals: number }>;
  recordCount: number;
  headerBytes: Uint8Array;
}

export interface OpcPackageGraph {
  schema: 'OpcPackageGraph';
  workbookPart: string;
  parts: Record<string, Uint8Array>;
  opaqueParts: Record<string, Uint8Array>;
  relationships: Record<string, NativeRelationship[]>;
  sheetPartById: Record<string, string>;
  contentTypesXml?: Uint8Array;
  dateSystem: DateSystem;
  format: Extract<NativeDocumentFormat, { family: 'ooxml' }>;
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

export interface NativeWorkbookPayload {
  name: string;
  sheetCount: number;
  dateSystem: DateSystem;
  compatibilityLevel: CompatibilityLevel;
}

export interface NativeDocumentImportResult {
  payload: NativeWorkbookPayload;
  report: CompatibilityReport;
  snapshot: import('@react-sheets/core-model').WorkbookSnapshot;
  artifact: NativeDocumentArtifact;
  taskId?: string;
}

export interface NativeDocumentExportResult {
  taskId: string;
  report: CompatibilityReport;
  buffer: ArrayBuffer;
  fileName: string;
  artifact: NativeDocumentArtifact;
}
