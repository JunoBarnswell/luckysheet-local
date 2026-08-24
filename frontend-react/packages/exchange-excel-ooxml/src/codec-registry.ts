import type { WorkbookSnapshot } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { exportXlsxWithWorker, importXlsxWithWorker } from './worker-port';
import type {
  CompatibilityReport,
  NativePackageState,
  XlsxExportOptions,
  XlsxExportResult,
  XlsxImportOptions,
  XlsxImportResult,
} from './types';
import type { XlsxWorkerPort } from './worker-port';

export type ExcelFormatFamily = 'ooxml' | 'xlsb' | 'biff' | 'text' | 'ods';

export interface ExcelImportTransaction {
  fileName: string;
  buffer: ArrayBuffer;
  options: XlsxImportOptions;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface ExcelExportTransaction {
  snapshot: WorkbookSnapshot;
  fileName: string;
  options: XlsxExportOptions;
  nativePackage?: NativePackageState;
  workerPort?: XlsxWorkerPort;
  execution?: 'worker' | 'inline-test';
  revision?: number;
}

export interface ExcelCodec<ImportRequest = unknown, ExportRequest = unknown> {
  family: ExcelFormatFamily;
  canRead(fileName: string, buffer: ArrayBuffer): boolean;
  import(request: ImportRequest): Promise<XlsxImportResult>;
  export(request: ExportRequest): Promise<XlsxExportResult>;
}

/** The single format boundary. Each future codec owns its native semantics. */
export class ExcelCodecRegistry {
  private readonly codecs: ExcelCodec[];

  constructor(codecs: ExcelCodec[] = [ooxmlCodec]) {
    this.codecs = [...codecs];
  }

  register(codec: ExcelCodec): void {
    if (this.codecs.some((entry) => entry.family === codec.family)) throw new Error(`Excel codec already registered: ${codec.family}`);
    this.codecs.push(codec);
  }

  detect(fileName: string, buffer: ArrayBuffer): ExcelCodec {
    const codec = this.codecs.find((entry) => entry.canRead(fileName, buffer));
    if (!codec) throw new Error(`No Excel codec can read ${fileName}`);
    return codec;
  }

  import(request: ExcelImportTransaction): Promise<XlsxImportResult> {
    return this.detect(request.fileName, request.buffer).import(request);
  }

  export(request: ExcelExportTransaction): Promise<XlsxExportResult> {
    const codec = request.nativePackage?.format.family === 'ooxml'
      ? this.codecs.find((entry) => entry.family === 'ooxml')
      : this.detect(request.fileName, new ArrayBuffer(0));
    if (!codec) throw new Error('No codec is registered for the workbook format');
    return codec.export(request);
  }
}

const ooxmlCodec: ExcelCodec<ExcelImportTransaction, ExcelExportTransaction> = {
  family: 'ooxml',
  canRead: (fileName, buffer) => /\.(xlsx|xlsm|xltx|xltm|xlam)$/i.test(fileName) || new Uint8Array(buffer).slice(0, 2)[0] === 0x50,
  import: (request) => request.execution === 'inline-test'
    ? importXlsx({ fileName: request.fileName, buffer: request.buffer, options: request.options })
    : importXlsxWithWorker({ fileName: request.fileName, buffer: request.buffer, options: request.options }, request.workerPort, request.revision ?? 0),
  export: (request) => request.execution === 'inline-test'
    ? exportXlsx({ snapshot: request.snapshot, fileName: request.fileName, options: request.options, nativePackage: request.nativePackage })
    : exportXlsxWithWorker({ snapshot: request.snapshot, fileName: request.fileName, options: request.options, nativePackage: request.nativePackage }, request.workerPort, request.revision ?? 0),
};

export const excelCodecRegistry = new ExcelCodecRegistry();
export type { CompatibilityReport };
