/** XLSX 兼容级别 */
export type CompatibilityLevel = 'A' | 'B' | 'C';

/** Excel 日期系统 */
export type DateSystem = '1900' | '1904';

export interface XlsxImportOptions {
  compatibilityTarget: CompatibilityLevel;
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
 * The original OOXML package accompanying an imported snapshot.
 *
 * `parts` contains every part, including binary and feature parts this package
 * does not edit.  Export overlays the editable core parts on this package so
 * chart/pivot/VBA/custom XML bytes are not silently discarded.
 */
export interface XlsxPackage {
  schema: 'XlsxPackage';
  parts: Record<string, Uint8Array>;
  opaqueParts: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  sheetPartById: Record<string, string>;
  contentTypesXml?: Uint8Array;
  dateSystem: DateSystem;
}

export type CompatibilityIssueSeverity = 'error' | 'warning' | 'info';

export interface CompatibilityIssue {
  level: CompatibilityLevel;
  severity: CompatibilityIssueSeverity;
  feature: string;
  location?: string;
  message: string;
  preserved: boolean;
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
  /** Pass this package to exportXlsx to preserve unsupported OOXML features. */
  package: XlsxPackage;
  taskId?: string;
}

export interface XlsxExportResult {
  taskId: string;
  report: CompatibilityReport;
  base64: string;
  fileName: string;
}
