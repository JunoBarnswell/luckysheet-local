/** XLSX 兼容级别 */
export type CompatibilityLevel = 'A' | 'B' | 'C';

/** Excel 日期系统 */
export type DateSystem = '1900' | '1904';

export interface XlsxImportOptions {
  compatibilityTarget: CompatibilityLevel;
  dateSystem?: DateSystem;
  preserveMacros?: boolean;
}

export interface XlsxExportOptions {
  compatibilityTarget: CompatibilityLevel;
  dateSystem?: DateSystem;
  includeCachedValues?: boolean;
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
  schema: 'CompatibilityReportV1';
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
  /** server-first: 实际 snapshot 由 exchange 任务返回 */
  taskId?: string;
}

export interface XlsxExportResult {
  taskId: string;
  report: CompatibilityReport;
}

export * from './date-system';
export * from './compatibility-report';
export * from './import';
export * from './export';
