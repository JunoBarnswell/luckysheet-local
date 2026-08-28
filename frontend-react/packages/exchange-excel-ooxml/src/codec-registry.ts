import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { strFromU8, unzipSync } from 'fflate';
import { exportOoxmlDocument } from './export';
import { importOoxmlDocument } from './import';
import { exportNativeDocumentWithWorker, importNativeDocumentWithWorker } from './worker-port';
import type {
  CompatibilityReport,
  NativeDocumentArtifact,
  NativeDocumentExportOptions,
  NativeDocumentExportResult,
  NativeDocumentImportOptions,
  NativeDocumentImportResult,
} from './types';
import { DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS } from './types';
import type { NativeDocumentWorkerPort } from './worker-port';
import { binaryCodec, dbfCodec, odsCodec, presentationCodec, sjsCodec, ssjsonCodec, textCodec, webCodec, worksCodec, xmlssCodec, xlsbCodec } from './native-codecs';
import { asNativeDocumentError } from './native-document-error';
import { loadOpcPackageGraph } from './ooxml';

export type NativeDocumentFamily = 'ooxml' | 'xlsb' | 'biff' | 'xmlss' | 'text' | 'ods' | 'sjs' | 'ssjson' | 'dbf' | 'works' | 'web' | 'presentation';

export interface NativeDocumentImportTransaction {
  fileName: string;
  buffer: ArrayBuffer;
  options: NativeDocumentImportOptions;
  workerPort?: NativeDocumentWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface NativeDocumentExportTransaction {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: NativeDocumentExportOptions;
  artifact?: NativeDocumentArtifact;
  mode?: 'save' | 'save-as' | 'export';
  workerPort?: NativeDocumentWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface NativeDocumentCodec<ImportRequest = unknown, ExportRequest = unknown> {
  family: NativeDocumentFamily;
  canRead(fileName: string, buffer: ArrayBuffer): boolean;
  import(request: ImportRequest): Promise<NativeDocumentImportResult>;
  export(request: ExportRequest): Promise<NativeDocumentExportResult>;
}

/** The only protocol detector. Suffixes are used only for empty Save targets. */
export class NativeFormatDetector {
  constructor(private readonly codecs: NativeDocumentCodec[]) {}

  detectCodec(fileName: string, buffer: ArrayBuffer): NativeDocumentCodec {
    const codec = this.codecs.find((entry) => entry.canRead(fileName, buffer));
    if (!codec) throw new Error(`NATIVE_DOCUMENT_DETECTION_FAILED: No native document codec can read ${fileName}`);
    return codec;
  }

  detectFormat(fileName: string, buffer: ArrayBuffer): import('./types').NativeDocumentFormat {
    const codec = this.detectCodec(fileName, buffer);
    if (codec.family === 'ooxml') return loadOpcPackageGraph(buffer, {}, fileName).packageGraph.format;
    if (codec.family === 'text') return textFormat(fileName, buffer);
    if (codec.family === 'xmlss') return { family: 'xmlss', variant: 'xml' };
    if (codec.family === 'ods') return { family: 'ods', variant: 'ods' };
    if (codec.family === 'sjs') return { family: 'sjs', variant: 'sjs' };
    if (codec.family === 'ssjson') return { family: 'ssjson', variant: 'ssjson' };
    if (codec.family === 'dbf') return { family: 'dbf', variant: 'dbf' };
    if (codec.family === 'works') return { family: 'works', variant: 'xlr' };
    if (codec.family === 'web') return { family: 'web', variant: fileName.toLowerCase().endsWith('.mht') || fileName.toLowerCase().endsWith('.mhtml') ? 'mht' : 'html' };
    if (codec.family === 'presentation') return { family: 'presentation', variant: fileName.toLowerCase().endsWith('.xps') ? 'xps' : 'pdf' };
    return binaryFormat(fileName, buffer, codec.family);
  }
}

/** The single format boundary. Each codec owns its native parse/write semantics. */
export class NativeDocumentCodecRegistry {
  private readonly codecs: NativeDocumentCodec[];
  private readonly detector: NativeFormatDetector;

  constructor(codecs: NativeDocumentCodec[] = [ooxmlCodec, textCodec, xmlssCodec, odsCodec, sjsCodec, ssjsonCodec, xlsbCodec, binaryCodec, dbfCodec, worksCodec, webCodec, presentationCodec]) {
    this.codecs = [...codecs];
    this.detector = new NativeFormatDetector(this.codecs);
  }

  register(codec: NativeDocumentCodec): void {
    if (this.codecs.some((entry) => entry.family === codec.family)) throw new Error(`Native document codec already registered: ${codec.family}`);
    this.codecs.push(codec);
  }

  detect(fileName: string, buffer: ArrayBuffer): NativeDocumentCodec {
    return this.detector.detectCodec(fileName, buffer);
  }

  list(): readonly NativeDocumentFamily[] {
    return this.codecs.map((codec) => codec.family);
  }

  async import(request: NativeDocumentImportTransaction): Promise<NativeDocumentImportResult> {
    if (request.execution !== 'inline-test') {
      return importNativeDocumentWithWorker({ fileName: request.fileName, buffer: request.buffer, options: request.options }, request.workerPort, request.revision ?? 0);
    }
    try {
      const codec = this.detect(request.fileName, request.buffer);
      return await codec.import(request);
    } catch (error) {
      throw asNativeDocumentError(error, { fileName: request.fileName });
    }
  }

  async export(request: NativeDocumentExportTransaction): Promise<NativeDocumentExportResult> {
    if (request.execution !== 'inline-test') {
      return exportNativeDocumentWithWorker({ snapshot: request.snapshot, fileName: request.fileName, options: request.options, artifact: request.artifact, mode: request.mode }, request.workerPort, request.revision ?? 0);
    }
    try {
      const codec = request.artifact && (request.mode === 'save' || formatMatchesFileName(request.artifact.format, request.fileName))
        ? this.codecs.find((entry) => entry.family === request.artifact!.format.family)
        : this.detect(request.fileName, new ArrayBuffer(0));
      if (!codec) throw new Error('NATIVE_DOCUMENT_DETECTION_FAILED: No codec is registered for the workbook format');
      return await codec.export(request);
    } catch (error) {
      throw asNativeDocumentError(error, { fileName: request.fileName, format: request.artifact?.format });
    }
  }

  /** Resolve the concrete native format without exposing codec ownership. */
  detectFormat(fileName: string, buffer: ArrayBuffer): import('./types').NativeDocumentFormat {
    return this.detector.detectFormat(fileName, buffer);
  }
}

const ooxmlCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'ooxml',
  canRead: (fileName, buffer) => looksLikeOoxml(fileName, buffer),
  import: (request) => importOoxmlDocument({ fileName: request.fileName, buffer: request.buffer, options: request.options }),
  export: (request) => exportOoxmlDocument({ snapshot: request.snapshot, fileName: request.fileName, options: request.options, artifact: request.artifact }),
};

export const nativeDocumentCodecRegistry = new NativeDocumentCodecRegistry();
export type { CompatibilityReport };

/** Public detector entry; callers must not infer protocol identity from a suffix. */
export function detectNativeDocumentFormat(fileName: string, buffer: ArrayBuffer): import('./types').NativeDocumentFormat {
  return nativeDocumentCodecRegistry.detectFormat(fileName, buffer);
}

function looksLikeOoxml(fileName: string, buffer: ArrayBuffer): boolean {
  const bytes = new Uint8Array(buffer);
  if (bytes.byteLength === 0) return /\.(xlsx|xlsm|xltx|xltm|xlam)$/i.test(fileName);
  if (bytes.byteLength > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxArchiveBytes) return false;
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false;
  try {
    let entries = 0;
    let total = 0;
    const parts = unzipSync(bytes, { filter(file) {
      entries += 1;
      total += file.originalSize;
      if (entries > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxEntries || file.originalSize > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxEntryBytes || total > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxUncompressedBytes) throw new Error('detection budget exceeded');
      return true;
    } });
    const contentTypes = parts['[Content_Types].xml'];
    return Boolean(contentTypes && /spreadsheetml|macroEnabled|template.main/i.test(strFromU8(contentTypes)));
  } catch {
    return false;
  }
}

function formatMatchesFileName(format: import('./types').NativeDocumentFormat, fileName: string): boolean {
  const extension = fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  if (!extension) return true;
  return format.family === 'ooxml' ? format.variant === extension
    : format.family === 'text' ? format.variant === extension || (format.variant === 'sylk' && extension === 'slk')
      : format.family === 'xmlss' ? extension === 'xml'
        : format.variant === extension;
}

function textFormat(fileName: string, buffer: ArrayBuffer): import('./types').NativeDocumentFormat {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.dif')) return { family: 'text', variant: 'dif' };
  if (lower.endsWith('.slk')) return { family: 'text', variant: 'sylk' };
  if (lower.endsWith('.prn')) return { family: 'text', variant: 'prn' };
  if (lower.endsWith('.txt')) return { family: 'text', variant: 'txt' };
  void buffer;
  return { family: 'text', variant: 'csv' };
}

function binaryFormat(fileName: string, buffer: ArrayBuffer, family: NativeDocumentFamily): import('./types').NativeDocumentFormat {
  const lower = fileName.toLowerCase();
  if (family === 'xlsb' || lower.endsWith('.xlsb')) return { family: 'xlsb', variant: 'xlsb' };
  if (lower.endsWith('.xlt')) return { family: 'biff', variant: 'xlt' };
  if (lower.endsWith('.xla')) return { family: 'biff', variant: 'xla' };
  if (lower.endsWith('.xlw')) return { family: 'biff', variant: 'xlw' };
  const bytes = new Uint8Array(buffer);
  return { family: 'biff', variant: bytes[0] === 0xd0 && bytes[1] === 0xcf ? 'xls' : 'biff5' };
}
