import type { KernelReplicaManifest, KernelReplicaPagePayload } from '@react-sheets/core-model';

/** Native formats are identified by the server/native host, never by a browser codec. */
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

export type CompatibilityLevel = 'A' | 'B' | 'C';
export type NativeCompatibilityMode = 'strict' | 'balanced' | 'best-effort';
export type DateSystem = '1900' | '1904';
export const NATIVE_DOCUMENT_CODEC_REVISION = 1 as const;

export interface NativeDocumentImportOptions { compatibilityTarget: CompatibilityLevel; compatibilityMode?: NativeCompatibilityMode; preserveMacros?: boolean; dateSystem?: DateSystem; }
export interface NativeDocumentExportOptions { compatibilityTarget: CompatibilityLevel; includeCachedValues?: boolean; preserveMacros?: boolean; dateSystem?: DateSystem; }

export interface CompatibilityIssue {
  level: CompatibilityLevel;
  severity: 'error' | 'warning' | 'info';
  feature: string;
  location?: string;
  message: string;
  preserved: boolean;
  status?: 'editable' | 'preserved-only' | 'unsupported';
  projection?: 'native' | 'projected' | 'preserved' | 'unsupported';
  reason?: string;
}
export interface CompatibilityReport {
  schema: 'CompatibilityReport';
  fileName: string;
  importLevel: CompatibilityLevel;
  exportLevel: CompatibilityLevel;
  dateSystem: DateSystem;
  issues: CompatibilityIssue[];
  summary: { editableFeatures: number; preservedOnly: number; unsupported: number };
}

/** Native bytes, package parts, and parser graphs remain on the server/native host. */
export interface NativeDocumentArtifact {
  schema: 'NativeDocumentArtifact';
  unitId: string;
  fileName: string;
  format: NativeDocumentFormat;
  checksum: string;
  byteLength: number;
  sourceRevision: number;
  sourceManifestHash?: string;
  dateSystem: DateSystem;
  detectedFeatures: string[];
  codecRevision: number;
  compatibility: CompatibilityReport;
}

export interface NativeDocumentImportRequest { fileName: string; content: Blob | ArrayBuffer; options: NativeDocumentImportOptions; formatHint?: string; }
export interface NativeDocumentImportResult { unitId: string; manifest: KernelReplicaManifest; pages?: KernelReplicaPagePayload[]; report: CompatibilityReport; artifact: NativeDocumentArtifact; }
export interface NativeDocumentExportRequest { unitId: string; revision: number; fileName: string; options: NativeDocumentExportOptions; }
export interface NativeDocumentExportResult { unitId: string; revision: number; content: ArrayBuffer; fileName: string; report: CompatibilityReport; artifact: NativeDocumentArtifact; }

/** Server/native host is the only implementation permitted in production. */
export interface NativeDocumentTransport {
  import(request: NativeDocumentImportRequest): Promise<NativeDocumentImportResult>;
  export(request: NativeDocumentExportRequest): Promise<NativeDocumentExportResult>;
}

export interface NativeDocumentCapability { family: NativeDocumentFormat['family']; variants: readonly string[]; import: 'server'; export: 'server'; preserveUnknownParts: true; }
