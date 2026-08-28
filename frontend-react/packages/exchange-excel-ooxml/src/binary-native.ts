import { strFromU8, unzipSync, zipSync } from 'fflate';
import { WorkbookModel, type CellValue, type WorkbookSnapshot } from '@react-sheets/core-model';
import { children, parseXml } from './xml';
import { createCompatibilityReport } from './compatibility-report';
import { createNativeDocumentArtifact, nativeSnapshotHash } from './native-document-artifact';
import { NativeDocumentError } from './native-document-error';
import {
  DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS,
  type BinaryCellGraph,
  type BinaryPackageGraph,
  type BinaryRecordGraph,
  type BinarySheetGraph,
  type CfbDirectoryEntryGraph,
  type CfbPackageGraph,
  type DateSystem,
  type NativeDocumentArtifact,
  type NativeDocumentExportResult,
  type NativeDocumentFormat,
  type NativeDocumentImportResult,
  type NativeDocumentResourceLimits,
  type NativeGraph,
} from './types';
import type { NativeDocumentCodec, NativeDocumentExportTransaction, NativeDocumentImportTransaction } from './codec-registry';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const CFB_FREE = -1;
const CFB_END = -2;
const CFB_FAT = -3;

const BIFF = {
  BOF: 0x0809,
  EOF: 0x000a,
  BOUNDSHEET8: 0x0085,
  SST: 0x00fc,
  CONTINUE: 0x003c,
  DIMENSIONS: 0x0200,
  BLANK: 0x0201,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  FORMULA: 0x0006,
  STRING: 0x0207,
  LABELSST: 0x00fd,
  RK: 0x027e,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  ARRAY: 0x0221,
  TABLE: 0x0236,
  SHRFMLA: 0x04bc,
  DATE1904: 0x0022,
} as const;

const XLSB = {
  CELLBLANK: 1,
  CELLRK: 2,
  CELLERROR: 3,
  CELLBOOL: 4,
  CELLREAL: 5,
  CELLST: 6,
  CELLISST: 7,
  FMLASTRING: 8,
  FMLANUM: 9,
  FMLABOOL: 10,
  FMLAERROR: 11,
  SHORTBLANK: 12,
  SHORTRK: 13,
  SHORTERROR: 14,
  SHORTBOOL: 15,
  SHORTREAL: 16,
  SHORTST: 17,
  SHORTISST: 18,
  SSTITEM: 19,
  BEGINSST: 231,
  ENDSST: 232,
  BEGINSHEET: 129,
  BUNDLESHEET: 156,
  ENDSHEET: 132,
  BEGINSHEETDATA: 145,
  ENDSHEETDATA: 146,
  ROWHDR: 0,
} as const;

const BIFF_CELL_TYPES = new Set<number>([BIFF.LABELSST, BIFF.NUMBER, BIFF.RK, BIFF.BOOLERR, BIFF.LABEL, BIFF.FORMULA, BIFF.BLANK]);
const BIFF_MULTI_CELL_TYPES = new Set<number>([BIFF.MULRK, BIFF.MULBLANK]);
const BIFF_UNSUPPORTED_CELL_TYPES = new Set<number>([BIFF.ARRAY, BIFF.TABLE, BIFF.SHRFMLA]);
const XLSB_LONG_CELL_TYPES = new Set<number>([XLSB.CELLBLANK, XLSB.CELLRK, XLSB.CELLERROR, XLSB.CELLBOOL, XLSB.CELLREAL, XLSB.CELLST, XLSB.CELLISST, XLSB.FMLASTRING, XLSB.FMLANUM, XLSB.FMLABOOL, XLSB.FMLAERROR]);
const XLSB_SHORT_CELL_TYPES = new Set<number>([XLSB.SHORTBLANK, XLSB.SHORTRK, XLSB.SHORTERROR, XLSB.SHORTBOOL, XLSB.SHORTREAL, XLSB.SHORTST, XLSB.SHORTISST]);
const XLSB_FORMULA_TYPES = new Set<number>([XLSB.FMLASTRING, XLSB.FMLANUM, XLSB.FMLABOOL, XLSB.FMLAERROR]);

interface CfbParsed {
  graph: CfbPackageGraph;
  streams: Record<string, Uint8Array>;
}

interface BinaryRecord {
  type: number;
  offset: number;
  bytes: Uint8Array;
  payload: Uint8Array;
}

interface BiffSheetDescriptor {
  name: string;
  hidden: boolean;
  type: BinarySheetGraph['type'];
  startOffset: number;
  boundRecordIndex: number;
  startRecordIndex: number;
  endRecordIndex: number;
}

interface BinaryParseResult {
  snapshot: WorkbookSnapshot;
  graph: NativeGraph;
  format: NativeDocumentFormat;
  dateSystem: DateSystem;
  features: string[];
}

interface NativeDocumentImportOptionsLike {
  compatibilityTarget: 'A' | 'B' | 'C';
  limits?: Partial<NativeDocumentResourceLimits>;
}

function limitsFor(options: { limits?: Partial<NativeDocumentResourceLimits> }): NativeDocumentResourceLimits {
  return { ...DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS, ...(options.limits ?? {}) };
}

function invalid(message: string): never {
  throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_INVALID', message });
}

function unsupported(message: string, recovery: string): never {
  throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_UNSUPPORTED', message, recovery });
}

function resource(message: string): never {
  throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message });
}

function assertInput(bytes: Uint8Array, limits: NativeDocumentResourceLimits, format: string): void {
  if (bytes.byteLength > limits.maxArchiveBytes) resource(`${format} input exceeds ${limits.maxArchiveBytes} bytes`);
}

function isCfb(bytes: Uint8Array): boolean {
  return bytes.length >= CFB_MAGIC.length && CFB_MAGIC.every((value, index) => bytes[index] === value);
}

function dataView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function readI32(view: DataView, offset: number): number {
  return view.getInt32(offset, true);
}

function readU32(view: DataView, offset: number): number {
  return view.getUint32(offset, true);
}

function writeI32(view: DataView, offset: number, value: number): void {
  view.setInt32(offset, value, true);
}

function writeU32(view: DataView, offset: number, value: number): void {
  view.setUint32(offset, value >>> 0, true);
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function toBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function readSector(bytes: Uint8Array, sector: number, sectorSize: number): Uint8Array {
  if (!Number.isSafeInteger(sector) || sector < 0) invalid(`CFB sector id is invalid: ${sector}`);
  const start = (sector + 1) * sectorSize;
  if (start + sectorSize > bytes.byteLength) invalid(`CFB sector ${sector} is outside the source bytes`);
  return bytes.slice(start, start + sectorSize);
}

function readChain(bytes: Uint8Array, start: number, table: readonly number[], sectorSize: number, maxSectors: number): Uint8Array {
  if (start === CFB_END || start === CFB_FREE) return new Uint8Array();
  const parts: Uint8Array[] = [];
  const seen = new Set<number>();
  let current = start;
  while (current !== CFB_END) {
    if (current < 0 || current >= table.length || seen.has(current)) invalid(`CFB sector chain is invalid at ${current}`);
    if (parts.length >= maxSectors) resource(`CFB stream exceeds ${maxSectors * sectorSize} bytes`);
    seen.add(current);
    parts.push(readSector(bytes, current, sectorSize));
    current = table[current]!;
  }
  return concatBytes(parts);
}

function readMiniChain(stream: Uint8Array, start: number, table: readonly number[], limits: NativeDocumentResourceLimits): Uint8Array {
  if (start < 0) invalid('CFB mini stream start sector is invalid');
  const parts: Uint8Array[] = [];
  const seen = new Set<number>();
  let current = start;
  while (current !== CFB_END) {
    if (current < 0 || current >= table.length || seen.has(current)) invalid(`CFB mini sector chain is invalid at ${current}`);
    if ((parts.length + 1) * 64 > limits.maxStreamBytes) resource('CFB mini stream exceeds the byte limit');
    const offset = current * 64;
    if (offset + 64 > stream.length) invalid('CFB mini sector is outside the root mini stream');
    seen.add(current);
    parts.push(stream.slice(offset, offset + 64));
    current = table[current]!;
  }
  return concatBytes(parts);
}

function readUtf16Name(bytes: Uint8Array, length: number): string {
  if (length === 0) return '';
  if (length < 2 || length > 64 || (length & 1) !== 0) invalid('CFB directory name length is invalid');
  return new TextDecoder('utf-16le').decode(bytes.slice(0, length - 2));
}

function parseCfb(bytes: Uint8Array, limits: NativeDocumentResourceLimits): CfbParsed {
  assertInput(bytes, limits, 'CFB document');
  if (!isCfb(bytes) || bytes.byteLength < 512) invalid('CFB header is missing or truncated');
  const header = dataView(bytes);
  const major = header.getUint16(0x1a, true);
  const sectorShift = header.getUint16(0x1e, true);
  const miniShift = header.getUint16(0x20, true);
  const sectorSize = sectorShift === 9 ? 512 : sectorShift === 12 ? 4096 : 0;
  if ((major !== 3 && major !== 4) || sectorSize === 0 || miniShift !== 6 || header.getUint16(0x1c, true) !== 0xfffe) invalid('CFB byte order or sector version is unsupported');

  const numFat = readU32(header, 0x2c);
  const firstDirectory = readI32(header, 0x30);
  const firstMiniFat = readI32(header, 0x3c);
  const numMiniFat = readU32(header, 0x40);
  const firstDifat = readI32(header, 0x44);
  const numDifat = readU32(header, 0x48);
  const difat: number[] = [];
  for (let index = 0; index < 109; index += 1) {
    const value = readI32(header, 0x4c + index * 4);
    if (value >= 0) difat.push(value);
  }
  const seenDifat = new Set<number>();
  let difatSector = firstDifat;
  for (let index = 0; index < numDifat; index += 1) {
    if (difatSector < 0 || seenDifat.has(difatSector)) invalid('CFB DIFAT chain is invalid');
    seenDifat.add(difatSector);
    const sector = readSector(bytes, difatSector, sectorSize);
    const view = dataView(sector);
    for (let entry = 0; entry < sectorSize / 4 - 1; entry += 1) {
      const value = readI32(view, entry * 4);
      if (value >= 0) difat.push(value);
    }
    difatSector = readI32(view, sectorSize - 4);
  }
  if (numFat > difat.length) invalid('CFB FAT sector list is incomplete');

  const fat: number[] = [];
  for (let index = 0; index < numFat; index += 1) {
    const sector = readSector(bytes, difat[index]!, sectorSize);
    const view = dataView(sector);
    for (let entry = 0; entry < sectorSize / 4; entry += 1) fat.push(readI32(view, entry * 4));
  }
  const maxSectors = Math.max(1, Math.min(fat.length + 1, Math.ceil(limits.maxStreamBytes / sectorSize) + 1));
  const directoryBytes = readChain(bytes, firstDirectory, fat, sectorSize, maxSectors);
  const entries: CfbDirectoryEntryGraph[] = [];
  for (let offset = 0; offset + 128 <= directoryBytes.byteLength; offset += 128) {
    const raw = directoryBytes.slice(offset, offset + 128);
    const view = dataView(raw);
    const nameLength = view.getUint16(0x40, true);
    const type = view.getUint8(0x42) as CfbDirectoryEntryGraph['type'];
    if (![0, 1, 2, 5].includes(type)) invalid(`CFB directory entry type is invalid: ${type}`);
    if (type === 0 && nameLength !== 0) invalid('CFB empty directory entry has a name');
    const name = type === 0 ? '' : readUtf16Name(raw.slice(0, 64), nameLength);
    const size = major === 4 ? readU32(view, 0x78) + readU32(view, 0x7c) * 0x100000000 : readU32(view, 0x78);
    if (!Number.isSafeInteger(size) || size > limits.maxStreamBytes) resource(`CFB stream ${name} exceeds ${limits.maxStreamBytes} bytes`);
    entries.push({ name, type, left: readI32(view, 0x44), right: readI32(view, 0x48), child: readI32(view, 0x4c), startSector: readI32(view, 0x74), size, raw });
    if (entries.length > limits.maxCfbStreams) resource(`CFB contains more than ${limits.maxCfbStreams} directory entries`);
  }
  const root = entries.find((entry) => entry.type === 5);
  if (!root) invalid('CFB root storage is missing');
  for (const entry of entries) {
    for (const reference of [entry.left, entry.right, entry.child]) if (reference < -1 || reference >= entries.length) invalid(`CFB directory reference is outside the directory: ${reference}`);
  }

  const miniFatBytes = numMiniFat && firstMiniFat >= 0 ? readChain(bytes, firstMiniFat, fat, sectorSize, maxSectors) : new Uint8Array();
  const miniFat: number[] = [];
  for (let offset = 0; offset + 4 <= miniFatBytes.length; offset += 4) miniFat.push(dataView(miniFatBytes.slice(offset, offset + 4)).getInt32(0, true));
  const rootMiniStream = root.size > 0 && root.startSector >= 0 ? readChain(bytes, root.startSector, fat, sectorSize, maxSectors).slice(0, root.size) : new Uint8Array();
  const streams: Record<string, Uint8Array> = {};
  const streamNames = new Set<string>();
  for (const entry of entries) {
    if (entry.type !== 2) continue;
    if (!entry.name || streamNames.has(entry.name)) invalid(`CFB stream name is duplicated or empty: ${entry.name || '<empty>'}`);
    streamNames.add(entry.name);
    let value: Uint8Array;
    if (entry.size === 0) value = new Uint8Array();
    else if (entry.size < 4096 && miniFat.length && rootMiniStream.length) value = readMiniChain(rootMiniStream, entry.startSector, miniFat, limits).slice(0, entry.size);
    else value = readChain(bytes, entry.startSector, fat, sectorSize, maxSectors).slice(0, entry.size);
    if (value.byteLength !== entry.size) invalid(`CFB stream ${entry.name} is truncated`);
    streams[entry.name] = value;
  }
  return { graph: { sectorSize: sectorSize as 512 | 4096, miniSectorSize: 64, majorVersion: major as 3 | 4, entries }, streams };
}

function writeName(raw: Uint8Array, name: string): void {
  raw.fill(0, 0, 64);
  const chars = name.slice(0, 31);
  for (let index = 0; index < chars.length; index += 1) dataView(raw).setUint16(index * 2, chars[index]!.charCodeAt(0), true);
  dataView(raw).setUint16(0x40, chars.length * 2 + 2, true);
}

/** Rebuilds the CFB sector tables while retaining every directory entry and stream payload. */
export function writeCfbPackage(
  graph: CfbPackageGraph,
  originalStreams: Record<string, Uint8Array>,
  updates: Record<string, Uint8Array>,
  limits: NativeDocumentResourceLimits = DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS,
): ArrayBuffer {
  const sectorSize = graph.sectorSize;
  if (graph.entries.length === 0 || graph.entries.length > limits.maxCfbStreams) resource('CFB directory entry count exceeds the native limit');
  const entries = graph.entries.map((entry) => ({ ...entry, raw: entry.raw.length === 128 ? entry.raw.slice() : new Uint8Array(128) }));
  const streams: Array<{ entryIndex: number; bytes: Uint8Array }> = [];
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index]!;
    if (entry.type !== 2) continue;
    const bytes = (updates[entry.name] ?? originalStreams[entry.name] ?? new Uint8Array()).slice();
    if (bytes.length > limits.maxStreamBytes) resource(`CFB stream ${entry.name} exceeds ${limits.maxStreamBytes} bytes`);
    streams.push({ entryIndex: index, bytes });
  }
  const entriesPerSector = sectorSize / 128;
  const directorySectors = Math.max(1, Math.ceil(entries.length / entriesPerSector));
  const dataSectors = streams.reduce((total, stream) => total + Math.ceil(stream.bytes.length / sectorSize), 0);
  let fatSectors = 1;
  for (;;) {
    const totalSectors = dataSectors + directorySectors + fatSectors;
    const next = Math.max(1, Math.ceil(totalSectors / (sectorSize / 4)));
    if (next === fatSectors) break;
    fatSectors = next;
  }
  if (fatSectors > 109) unsupported('CFB writer cannot represent this document within the header DIFAT capacity', 'Reduce the document size before saving the binary document.');

  const firstDirectory = dataSectors;
  const firstFat = firstDirectory + directorySectors;
  const totalSectors = firstFat + fatSectors;
  const fat = new Int32Array(totalSectors);
  fat.fill(CFB_FREE);
  const link = (start: number, count: number, terminal = CFB_END): void => {
    for (let index = 0; index < count; index += 1) fat[start + index] = index + 1 < count ? start + index + 1 : terminal;
  };
  let nextSector = 0;
  for (const stream of streams) {
    const entry = entries[stream.entryIndex]!;
    const count = Math.ceil(stream.bytes.length / sectorSize);
    entry.startSector = count ? nextSector : CFB_END;
    entry.size = stream.bytes.length;
    if (count) link(nextSector, count);
    nextSector += count;
  }
  link(firstDirectory, directorySectors);
  for (let index = 0; index < fatSectors; index += 1) fat[firstFat + index] = CFB_FAT;

  const directory = new Uint8Array(directorySectors * sectorSize);
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index]!;
    const raw = entry.raw.slice();
    const view = dataView(raw);
    raw[0x42] = entry.type;
    if (entry.type === 0) {
      raw.fill(0, 0, 0x44);
    } else {
      writeName(raw, entry.name);
      writeI32(view, 0x44, entry.left);
      writeI32(view, 0x48, entry.right);
      writeI32(view, 0x4c, entry.child);
    }
    if (entry.type === 5) {
      writeI32(view, 0x74, CFB_END);
      writeU32(view, 0x78, 0);
      writeU32(view, 0x7c, 0);
    } else if (entry.type === 2) {
      writeI32(view, 0x74, entry.startSector);
      writeU32(view, 0x78, entry.size);
      writeU32(view, 0x7c, 0);
    }
    directory.set(raw, index * 128);
  }

  const result = new Uint8Array((totalSectors + 1) * sectorSize);
  for (const stream of streams) {
    const entry = entries[stream.entryIndex]!;
    if (entry.size > 0) result.set(stream.bytes, (entry.startSector + 1) * sectorSize);
  }
  result.set(directory, (firstDirectory + 1) * sectorSize);
  for (let index = 0; index < fatSectors; index += 1) {
    const view = new DataView(result.buffer, (firstFat + index + 1) * sectorSize, sectorSize);
    for (let offset = 0; offset < sectorSize / 4; offset += 1) view.setInt32(offset * 4, fat[index * (sectorSize / 4) + offset] ?? CFB_FREE, true);
  }
  result.set(Uint8Array.from(CFB_MAGIC), 0);
  const header = dataView(result.slice(0, sectorSize));
  header.setUint16(0x18, 0x003e, true);
  header.setUint16(0x1a, graph.majorVersion, true);
  header.setUint16(0x1c, 0xfffe, true);
  header.setUint16(0x1e, sectorSize === 512 ? 9 : 12, true);
  header.setUint16(0x20, 6, true);
  header.setUint32(0x28, graph.majorVersion === 4 ? directorySectors : 0, true);
  header.setUint32(0x2c, fatSectors, true);
  writeI32(header, 0x30, firstDirectory);
  writeI32(header, 0x3c, CFB_END);
  header.setUint32(0x40, 0, true);
  writeI32(header, 0x44, CFB_END);
  header.setUint32(0x48, 0, true);
  header.setUint32(0x38, 4096, true);
  for (let index = 0; index < 109; index += 1) writeI32(header, 0x4c + index * 4, index < fatSectors ? firstFat + index : CFB_FREE);
  result.set(new Uint8Array(header.buffer, header.byteOffset, header.byteLength), 0);
  return result.buffer.slice(result.byteOffset, result.byteOffset + result.byteLength) as ArrayBuffer;
}

function parseBinaryRecords(bytes: Uint8Array, limits: NativeDocumentResourceLimits): BinaryRecord[] {
  const records: BinaryRecord[] = [];
  let offset = 0;
  while (offset < bytes.length) {
    if (offset + 4 > bytes.length) invalid(`BIFF record header is truncated at ${offset}`);
    const view = dataView(bytes.slice(offset));
    const type = view.getUint16(0, true);
    const size = view.getUint16(2, true);
    const end = offset + 4 + size;
    if (end > bytes.length) invalid(`BIFF record payload is truncated at ${offset}`);
    records.push({ type, offset, bytes: bytes.slice(offset, end), payload: bytes.slice(offset + 4, end) });
    if (records.length > limits.maxRecordCount) resource(`BIFF contains more than ${limits.maxRecordCount} records`);
    offset = end;
  }
  return records;
}

function readShortString(payload: Uint8Array, offset: number, requireFlags = true): { value: string; next: number } {
  if (offset + 2 > payload.length) invalid('BIFF string header is truncated');
  const view = dataView(payload);
  const count = view.getUint16(offset, true);
  let cursor = offset + 2;
  if (requireFlags) {
    if (cursor >= payload.length) invalid('BIFF string flags are missing');
    const flags = payload[cursor]!;
    cursor += 1;
    const unicode = (flags & 1) !== 0;
    const byteLength = count * (unicode ? 2 : 1);
    if (cursor + byteLength > payload.length) invalid('BIFF string payload is truncated');
    return { value: unicode ? new TextDecoder('utf-16le').decode(payload.slice(cursor, cursor + byteLength)) : new TextDecoder('windows-1252').decode(payload.slice(cursor, cursor + byteLength)), next: cursor + byteLength };
  }
  if (cursor + count > payload.length) invalid('BIFF string payload is truncated');
  return { value: new TextDecoder('windows-1252').decode(payload.slice(cursor, cursor + count)), next: cursor + count };
}

function readBiffLabel(payload: Uint8Array): string {
  if (payload.length < 8) invalid('BIFF Label record is truncated');
  const count = dataView(payload).getUint16(6, true);
  if (payload.length >= 9 + count * 2 && (payload[8]! & 0xfe) === 0) return readShortString(payload, 6, true).value;
  return readShortString(payload, 6, false).value;
}

function readFormulaString(payload: Uint8Array): string {
  if (payload.length < 2) invalid('BIFF Formula string record is truncated');
  const count = dataView(payload).getUint16(0, true);
  return payload.length === 2 + count ? readShortString(payload, 0, false).value : readShortString(payload, 0, true).value;
}

function readBoundSheet(record: BinaryRecord, index: number): BiffSheetDescriptor {
  if (record.payload.length < 7) invalid(`BoundSheet8 record ${index} is truncated`);
  const view = dataView(record.payload);
  const startOffset = view.getUint32(0, true);
  const state = record.payload[4]! & 0x03;
  const sheetType = record.payload[5]!;
  const count = record.payload[6]!;
  const legacy = record.payload.length === 7 + count;
  const unicode = !legacy && (record.payload[7]! & 1) !== 0;
  const nameOffset = legacy ? 7 : 8;
  const byteLength = count * (unicode ? 2 : 1);
  if (nameOffset + byteLength > record.payload.length) invalid(`BoundSheet8 record ${index} has a truncated name`);
  const name = unicode ? new TextDecoder('utf-16le').decode(record.payload.slice(nameOffset, nameOffset + byteLength)) : new TextDecoder('windows-1252').decode(record.payload.slice(nameOffset, nameOffset + byteLength));
  return {
    name: name || `Sheet${index + 1}`,
    hidden: state !== 0,
    type: sheetType === 1 ? 'macro' : sheetType === 2 ? 'chart' : sheetType === 6 ? 'module' : sheetType === 0 ? 'worksheet' : 'unknown',
    startOffset,
    boundRecordIndex: record.offset,
    startRecordIndex: -1,
    endRecordIndex: -1,
  };
}

function parseSst(records: readonly BinaryRecord[], limits: NativeDocumentResourceLimits): string[] {
  const start = records.findIndex((record) => record.type === BIFF.SST);
  if (start < 0) return [];
  if (records[start + 1]?.type === BIFF.CONTINUE) unsupported('BIFF shared strings split across CONTINUE records are not editable in the local codec', 'Use a workbook with native string records that fit one SST record or save as XLSX.');
  const chunks: Uint8Array[] = [records[start]!.payload];
  const bytes = concatBytes(chunks);
  if (bytes.length < 8) invalid('BIFF SST header is truncated');
  const view = dataView(bytes);
  const unique = view.getUint32(4, true);
  if (unique > limits.maxCells) resource('BIFF shared string table exceeds the cell limit');
  const result: string[] = [];
  let offset = 8;
  for (let index = 0; index < unique; index += 1) {
    if (offset + 3 > bytes.length) invalid('BIFF SST item header is truncated');
    const count = view.getUint16(offset, true);
    const flags = bytes[offset + 2]!;
    offset += 3;
    const richRuns = (flags & 0x08) !== 0 ? view.getUint16(offset, true) : 0;
    if (flags & 0x08) offset += 2;
    const extSize = (flags & 0x04) !== 0 ? view.getUint32(offset, true) : 0;
    if (flags & 0x04) offset += 4;
    const byteLength = count * ((flags & 1) !== 0 ? 2 : 1);
    if (offset + byteLength + richRuns * 4 + extSize > bytes.length) invalid('BIFF SST item is truncated');
    result.push((flags & 1) !== 0 ? new TextDecoder('utf-16le').decode(bytes.slice(offset, offset + byteLength)) : new TextDecoder('windows-1252').decode(bytes.slice(offset, offset + byteLength)));
    offset += byteLength + richRuns * 4 + extSize;
  }
  return result;
}

function decodeRk(raw: number): number {
  if (raw & 2) {
    const value = raw >> 2;
    return raw & 1 ? value / 100 : value;
  }
  const bytes = new Uint8Array(8);
  dataView(bytes).setUint32(4, raw & 0xfffffffc, true);
  return dataView(bytes).getFloat64(0, true);
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function parseBiffCell(record: BinaryRecord, recordIndex: number, records: readonly BinaryRecord[], sharedStrings: readonly string[]): BinaryCellGraph | undefined {
  if (!BIFF_CELL_TYPES.has(record.type)) return undefined;
  if (record.payload.length < 6) invalid(`BIFF cell record ${record.type} is truncated`);
  const view = dataView(record.payload);
  const row = view.getUint16(0, true);
  const column = view.getUint16(2, true);
  const styleIndex = view.getUint16(4, true);
  let value: CellValue = null;
  let auxiliaryRecordIndex: number | undefined;
  if (record.type === BIFF.LABELSST) {
    if (record.payload.length < 10) invalid('BIFF LabelSst record is truncated');
    const index = view.getUint32(6, true);
    if (index >= sharedStrings.length) invalid(`BIFF shared string index is out of range: ${index}`);
    value = sharedStrings[index] ?? '';
  } else if (record.type === BIFF.NUMBER) {
    if (record.payload.length < 14) invalid('BIFF Number record is truncated');
    value = view.getFloat64(6, true);
  } else if (record.type === BIFF.RK) {
    if (record.payload.length < 10) invalid('BIFF RK record is truncated');
    value = decodeRk(view.getUint32(6, true));
  } else if (record.type === BIFF.BOOLERR) {
    if (record.payload.length < 8) invalid('BIFF BoolErr record is truncated');
    value = record.payload[6] === 0 ? Boolean(record.payload[7]) : `#ERR${record.payload[7]}`;
  } else if (record.type === BIFF.LABEL) {
    value = readBiffLabel(record.payload);
  } else if (record.type === BIFF.FORMULA) {
    if (record.payload.length < 14) invalid('BIFF Formula record is truncated');
    const resultBytes = record.payload.slice(6, 14);
    const special = resultBytes.slice(0, 6).every((entry) => entry === 0xff);
    if (special && records[recordIndex + 1]?.type === BIFF.STRING) {
      value = readFormulaString(records[recordIndex + 1]!.payload);
      auxiliaryRecordIndex = recordIndex + 1;
    } else value = dataView(resultBytes).getFloat64(0, true);
  }
  return { row, column, recordIndex, recordType: record.type, value, styleIndex, ...(auxiliaryRecordIndex === undefined ? {} : { auxiliaryRecordIndex }) };
}

function parseBiffMultiCells(record: BinaryRecord, recordIndex: number): BinaryCellGraph[] {
  if (!BIFF_MULTI_CELL_TYPES.has(record.type)) return [];
  if (record.payload.length < (record.type === BIFF.MULRK ? 12 : 8)) invalid(`BIFF multi-cell record ${record.type} is truncated`);
  const view = dataView(record.payload);
  const row = view.getUint16(0, true);
  const firstColumn = view.getUint16(2, true);
  const lastColumn = view.getUint16(record.payload.length - 2, true);
  if (lastColumn < firstColumn) invalid(`BIFF multi-cell record ${record.type} has an invalid column range`);
  const count = lastColumn - firstColumn + 1;
  const itemBytes = record.type === BIFF.MULRK ? 6 : 2;
  const expected = 6 + count * itemBytes;
  if (record.payload.length < expected) invalid(`BIFF multi-cell record ${record.type} has a truncated item list`);
  const cells: BinaryCellGraph[] = [];
  for (let index = 0; index < count; index += 1) {
    const offset = 4 + index * itemBytes;
    const styleIndex = view.getUint16(offset, true);
    const value = record.type === BIFF.MULRK ? decodeRk(view.getUint32(offset + 2, true)) : null;
    assertBiffCoordinate(row, firstColumn + index);
    cells.push({ row, column: firstColumn + index, recordIndex, recordType: record.type, value, styleIndex });
  }
  return cells;
}

function snapshotFromBiff(name: string, descriptors: readonly BiffSheetDescriptor[], sheetCells: readonly Record<string, BinaryCellGraph>[]): WorkbookSnapshot {
  const workbook = new WorkbookModel(`imported-biff-${Date.now().toString(36)}`, name);
  const first = workbook.getSheet(workbook.primarySheetId);
  if (descriptors.length === 0) first.name = 'Sheet1';
  descriptors.forEach((descriptor, index) => {
    const sheet = index === 0 ? first : workbook.addSheet(`sheet-${index + 1}`, descriptor.name);
    sheet.name = descriptor.name;
    sheet.hidden = descriptor.hidden;
    for (const cell of Object.values(sheetCells[index] ?? {})) if (cell.value !== null) sheet.cells.set(cell.row, cell.column, { value: cell.value });
  });
  const snapshot = workbook.snapshot();
  snapshot.name = name;
  return snapshot;
}

function parseBiffDocument(bytes: Uint8Array, fileName: string, limits: NativeDocumentResourceLimits): BinaryParseResult {
  const cfb = parseCfb(bytes, limits);
  const workbookStreamName = Object.keys(cfb.streams).find((name) => /^workbook$/i.test(name) || /^book$/i.test(name));
  if (!workbookStreamName) {
    if (Object.keys(cfb.streams).some((name) => /^(EncryptedPackage|EncryptionInfo)$/i.test(name))) unsupported('Encrypted Office binary content cannot be projected by the local BIFF codec', 'Open the document with an approved decryption provider before importing it.');
    invalid('CFB Workbook stream is missing');
  }
  const stream = cfb.streams[workbookStreamName!]!;
  const records = parseBinaryRecords(stream, limits);
  const bof = records.find((record) => record.type === BIFF.BOF);
  if (!bof || bof.payload.length < 2) invalid('BIFF workbook BOF is missing');
  const version = dataView(bof.payload).getUint16(0, true);
  const lower = fileName.toLowerCase();
  const variant: Extract<NativeDocumentFormat, { family: 'biff' }>['variant'] = lower.endsWith('.xlt') ? 'xlt' : lower.endsWith('.xla') ? 'xla' : lower.endsWith('.xlw') ? 'xlw' : version === 0x0600 ? 'xls' : 'biff5';
  const boundRecords = records.map((record, index) => ({ record, index })).filter(({ record }) => record.type === BIFF.BOUNDSHEET8);
  const descriptors = boundRecords.map(({ record, index }, item) => ({ ...readBoundSheet(record, item), boundRecordIndex: index }));
  if (!descriptors.length) invalid('BIFF workbook has no BoundSheet8 records');
  const ordered = descriptors.map((descriptor) => descriptor).sort((left, right) => left.startOffset - right.startOffset);
  for (let index = 0; index < ordered.length; index += 1) {
    const descriptor = ordered[index]!;
    if (index > 0 && descriptor.startOffset <= ordered[index - 1]!.startOffset) invalid('BIFF BoundSheet8 offsets are not strictly ordered');
    const start = records.findIndex((record) => record.offset === descriptor.startOffset);
    if (start < 0) invalid(`BoundSheet8 points to missing sheet offset: ${descriptor.startOffset}`);
    descriptor.startRecordIndex = start;
    const nextOffset = ordered[index + 1]?.startOffset ?? Number.MAX_SAFE_INTEGER;
    const end = records.findIndex((record, recordIndex) => recordIndex >= start && record.type === BIFF.EOF && record.offset < nextOffset);
    descriptor.endRecordIndex = end >= 0 ? end : records.length - 1;
  }
  const sharedStrings = parseSst(records, limits);
  const sheetCells: Record<string, BinaryCellGraph>[] = [];
  const features = new Set<string>(['cells', 'biff']);
  if (sharedStrings.length) features.add('sharedStrings');
  let dateSystem: DateSystem = '1900';
  if (records.some((record) => record.type === BIFF.DATE1904 && record.payload[0])) dateSystem = '1904';
  for (const descriptor of descriptors) {
    const cells: Record<string, BinaryCellGraph> = {};
    for (let index = descriptor.startRecordIndex; index <= descriptor.endRecordIndex; index += 1) {
      const record = records[index]!;
      if (BIFF_UNSUPPORTED_CELL_TYPES.has(record.type)) unsupported(`BIFF cell structure ${record.type} is not editable in ${descriptor.name}`, 'Leave the native formula structure unchanged or save through a format with an owned formula writer.');
      const multi = parseBiffMultiCells(record, index);
      if (multi.length) {
        for (const cell of multi) cells[cellKey(cell.row, cell.column)] = cell;
        continue;
      }
      const cell = parseBiffCell(record, index, records, sharedStrings);
      if (!cell) continue;
      assertBiffCoordinate(cell.row, cell.column);
      cells[cellKey(cell.row, cell.column)] = cell;
      if (cell.recordType === BIFF.FORMULA) features.add('formulas');
    }
    sheetCells.push(cells);
  }
  const graph: BinaryRecordGraph = {
    container: 'cfb',
    records: records.map((record) => ({ type: record.type, offset: record.offset, bytes: record.bytes.slice() })),
    opaque: new Uint8Array(),
    streamName: workbookStreamName,
    streams: Object.fromEntries(Object.entries(cfb.streams).map(([name, value]) => [name, value.slice()])),
    cfb: cfb.graph,
    sheets: descriptors.map((descriptor, index) => ({ name: descriptor.name, hidden: descriptor.hidden, type: descriptor.type, boundRecordIndex: descriptor.boundRecordIndex, startRecord: descriptor.startRecordIndex, endRecord: descriptor.endRecordIndex, cells: sheetCells[index]! })),
    sharedStrings,
    dateSystem,
  };
  const snapshot = snapshotFromBiff(fileName.replace(/\.[^.]+$/, ''), descriptors, sheetCells);
  return { snapshot, graph: { kind: 'biff', container: graph }, format: { family: 'biff', variant }, dateSystem, features: [...features] };
}

function biffRecord(type: number, payload: Uint8Array): Uint8Array {
  if (payload.length > 0xffff) invalid(`BIFF record ${type} payload exceeds 65535 bytes`);
  const bytes = new Uint8Array(payload.length + 4);
  const view = dataView(bytes);
  view.setUint16(0, type, true);
  view.setUint16(2, payload.length, true);
  bytes.set(payload, 4);
  return bytes;
}

function assertBiffCoordinate(row: number, column: number): void {
  if (!Number.isInteger(row) || row < 0 || row > 0xffff || !Number.isInteger(column) || column < 0 || column > 0xff) invalid(`BIFF cell coordinate is outside the native range: ${row},${column}`);
}

function biffCellHeader(row: number, column: number, style: number): Uint8Array {
  assertBiffCoordinate(row, column);
  const payload = new Uint8Array(6);
  const view = dataView(payload);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, style & 0xffff, true);
  return payload;
}

function biffNumber(row: number, column: number, style: number, value: number): Uint8Array {
  const payload = new Uint8Array(14);
  payload.set(biffCellHeader(row, column, style));
  dataView(payload).setFloat64(6, value, true);
  return biffRecord(BIFF.NUMBER, payload);
}

function biffBlank(row: number, column: number, style: number): Uint8Array {
  return biffRecord(BIFF.BLANK, biffCellHeader(row, column, style));
}

function biffBool(row: number, column: number, style: number, value: boolean): Uint8Array {
  const payload = new Uint8Array(8);
  payload.set(biffCellHeader(row, column, style));
  payload[6] = value ? 1 : 0;
  payload[7] = 0;
  return biffRecord(BIFF.BOOLERR, payload);
}

function biffLabel(row: number, column: number, style: number, value: string, unicode = true): Uint8Array {
  const chars = value.slice(0, 32767);
  if (!unicode) {
    if ([...chars].some((character) => character.charCodeAt(0) > 0xff)) unsupported('BIFF5 cannot encode a non-single-byte cell string', 'Save the workbook as BIFF8 or another Unicode-capable format.');
    const payload = new Uint8Array(8 + chars.length);
    payload.set(biffCellHeader(row, column, style));
    dataView(payload).setUint16(6, chars.length, true);
    for (let index = 0; index < chars.length; index += 1) payload[8 + index] = chars.charCodeAt(index) & 0xff;
    return biffRecord(BIFF.LABEL, payload);
  }
  const payload = new Uint8Array(9 + chars.length * 2);
  payload.set(biffCellHeader(row, column, style));
  const view = dataView(payload);
  view.setUint16(6, chars.length, true);
  payload[8] = 1;
  for (let index = 0; index < chars.length; index += 1) view.setUint16(9 + index * 2, chars.charCodeAt(index), true);
  return biffRecord(BIFF.LABEL, payload);
}

function sameCellValue(left: CellValue, right: CellValue): boolean {
  return left === right || (typeof left === 'number' && typeof right === 'number' && Number.isNaN(left) && Number.isNaN(right));
}

function snapshotCellValue(snapshot: WorkbookSnapshot['sheets'][number], row: number, column: number): CellValue {
  return snapshot.cells[String(row)]?.[String(column)]?.value ?? null;
}

function snapshotCellEntries(snapshot: WorkbookSnapshot['sheets'][number]): Array<{ row: number; column: number; value: CellValue }> {
  const entries: Array<{ row: number; column: number; value: CellValue }> = [];
  for (const [rowKey, columns] of Object.entries(snapshot.cells)) {
    const row = Number(rowKey);
    if (!Number.isSafeInteger(row)) invalid(`Workbook snapshot row is invalid: ${rowKey}`);
    for (const [columnKey, cell] of Object.entries(columns)) {
      const column = Number(columnKey);
      if (!Number.isSafeInteger(column)) invalid(`Workbook snapshot column is invalid: ${columnKey}`);
      entries.push({ row, column, value: cell.value });
    }
  }
  return entries;
}

function encodeBiffCell(row: number, column: number, style: number, value: CellValue, unicode = true): Uint8Array {
  if (value === null) return biffBlank(row, column, style);
  if (typeof value === 'number') return biffNumber(row, column, style, value);
  if (typeof value === 'boolean') return biffBool(row, column, style, value);
  return biffLabel(row, column, style, String(value), unicode);
}

function updateBiffDimensions(bytes: Uint8Array, sheet: WorkbookSnapshot['sheets'][number]): void {
  const values = snapshotCellEntries(sheet).filter((entry) => entry.value !== null);
  if (!values.length) return;
  const minRow = Math.min(...values.map((entry) => entry.row));
  const maxRow = Math.max(...values.map((entry) => entry.row));
  const minColumn = Math.min(...values.map((entry) => entry.column));
  const maxColumn = Math.max(...values.map((entry) => entry.column));
  assertBiffCoordinate(maxRow, maxColumn);
  const payloadLength = bytes.length - 4;
  const view = dataView(bytes);
  if (payloadLength >= 14) {
    view.setUint32(4, minRow, true);
    view.setUint32(8, maxRow + 1, true);
    view.setUint16(12, minColumn, true);
    view.setUint16(14, maxColumn + 1, true);
  } else if (payloadLength >= 10 && maxRow <= 0xffff) {
    view.setUint16(4, minRow, true);
    view.setUint16(6, maxRow + 1, true);
    view.setUint16(8, minColumn, true);
    view.setUint16(10, maxColumn + 1, true);
  }
}

function rewriteBiffStream(graph: BinaryRecordGraph, snapshot: WorkbookSnapshot, limits: NativeDocumentResourceLimits, variant: Extract<NativeDocumentFormat, { family: 'biff' }>['variant']): Uint8Array {
  if (graph.container !== 'cfb' || !graph.cfb || !graph.streams || !graph.streamName || !graph.sheets) invalid('BIFF artifact does not contain a writable CFB graph');
  if (snapshot.sheets.length < graph.sheets.length) unsupported('The workbook lost a native sheet during BIFF Save', 'Keep every imported sheet in the workbook before saving.');
  const original = parseBinaryRecords(graph.streams[graph.streamName]!, limits);
  const replacements = new Map<number, Uint8Array>();
  const inserts = new Map<number, Uint8Array[]>();

  for (let sheetIndex = 0; sheetIndex < graph.sheets.length; sheetIndex += 1) {
    const nativeSheet = graph.sheets[sheetIndex]!;
    const sheet = snapshot.sheets[sheetIndex];
    if (!sheet) invalid(`BIFF snapshot is missing sheet ${nativeSheet.name}`);
    const seen = new Set<string>();
    const multiBindings = new Map<number, BinaryCellGraph[]>();
    for (const binding of Object.values(nativeSheet.cells)) {
      const key = cellKey(binding.row, binding.column);
      seen.add(key);
      const current = snapshotCellValue(sheet!, binding.row, binding.column);
      if (BIFF_MULTI_CELL_TYPES.has(binding.recordType)) {
        const group = multiBindings.get(binding.recordIndex) ?? [];
        group.push(binding);
        multiBindings.set(binding.recordIndex, group);
        continue;
      }
      if (sameCellValue(current, binding.value)) continue;
      if (binding.recordType === BIFF.FORMULA) unsupported(`Formula cell ${nativeSheet.name}!${binding.row}:${binding.column} cannot be rewritten from a value-only snapshot`, 'Preserve the formula expression or leave the formula cell unchanged.');
      replacements.set(binding.recordIndex, encodeBiffCell(binding.row, binding.column, binding.styleIndex ?? 0, current, variant !== 'biff5'));
    }
    for (const [recordIndex, bindings] of multiBindings) {
      let changed = false;
      const replacement = bindings
        .sort((left, right) => left.column - right.column)
        .map((binding) => {
          const current = snapshotCellValue(sheet!, binding.row, binding.column);
          changed ||= !sameCellValue(current, binding.value);
          return encodeBiffCell(binding.row, binding.column, binding.styleIndex ?? 0, current, variant !== 'biff5');
        });
      if (changed) replacements.set(recordIndex, concatBytes(replacement));
    }
    const additional: Uint8Array[] = [];
    for (const entry of snapshotCellEntries(sheet!)) {
      const key = cellKey(entry.row, entry.column);
      if (seen.has(key) || entry.value === null) continue;
      additional.push(encodeBiffCell(entry.row, entry.column, 0, entry.value, variant !== 'biff5'));
    }
    if (additional.length) inserts.set(nativeSheet.endRecord, [...(inserts.get(nativeSheet.endRecord) ?? []), ...additional]);
    const dimensionIndex = original.findIndex((record, index) => index >= nativeSheet.startRecord && index <= nativeSheet.endRecord && record.type === BIFF.DIMENSIONS);
    if (dimensionIndex >= 0) {
      const dimension = (replacements.get(dimensionIndex) ?? original[dimensionIndex]!.bytes).slice();
      updateBiffDimensions(dimension, sheet!);
      replacements.set(dimensionIndex, dimension);
    }
  }
  for (let index = graph.sheets.length; index < snapshot.sheets.length; index += 1) {
    if (snapshotCellEntries(snapshot.sheets[index]!).some((entry) => entry.value !== null)) unsupported('New BIFF sheets are not representable by the imported native graph', 'Edit existing native sheets or create a new workbook in a format with sheet creation support.');
  }

  const output: Uint8Array[] = [];
  const outputIndex = new Map<number, number>();
  const outputOffsets = new Map<number, number>();
  let offset = 0;
  for (let index = 0; index < original.length; index += 1) {
    for (const inserted of inserts.get(index) ?? []) {
      output.push(inserted);
      offset += inserted.length;
    }
    outputIndex.set(index, output.length);
    outputOffsets.set(index, offset);
    const bytes = replacements.get(index) ?? original[index]!.bytes;
    output.push(bytes);
    offset += bytes.length;
  }
  for (const sheet of graph.sheets) {
    if (sheet.boundRecordIndex === undefined) continue;
    const boundOutputIndex = outputIndex.get(sheet.boundRecordIndex);
    const startOffset = outputOffsets.get(sheet.startRecord);
    if (boundOutputIndex === undefined || startOffset === undefined) invalid(`BIFF sheet offset mapping is incomplete for ${sheet.name}`);
    const bytes = output[boundOutputIndex]!.slice();
    if (bytes.length < 8) invalid(`BoundSheet8 record is truncated for ${sheet.name}`);
    writeU32(dataView(bytes), 4, startOffset);
    output[boundOutputIndex] = bytes;
  }
  return concatBytes(output);
}

function readVarUint(bytes: Uint8Array, offset: number): { value: number; next: number } {
  let value = 0;
  let shift = 0;
  let cursor = offset;
  while (cursor < bytes.length && shift <= 28) {
    const byte = bytes[cursor++]!;
    value |= (byte & 0x7f) << shift;
    if ((byte & 0x80) === 0) return { value: value >>> 0, next: cursor };
    shift += 7;
  }
  invalid(`BIFF12 variable integer is invalid at ${offset}`);
}

function varUint(value: number): Uint8Array {
  const bytes: number[] = [];
  let current = value >>> 0;
  do {
    let byte = current & 0x7f;
    current >>>= 7;
    if (current) byte |= 0x80;
    bytes.push(byte);
  } while (current);
  return Uint8Array.from(bytes);
}

function parseBiff12Records(bytes: Uint8Array, limits: NativeDocumentResourceLimits): BinaryRecord[] {
  const records: BinaryRecord[] = [];
  let offset = 0;
  while (offset < bytes.length) {
    const typeHeader = readVarUint(bytes, offset);
    const sizeHeader = readVarUint(bytes, typeHeader.next);
    const end = sizeHeader.next + sizeHeader.value;
    if (end > bytes.length) invalid(`BIFF12 record payload is truncated at ${offset}`);
    records.push({ type: typeHeader.value, offset, bytes: bytes.slice(offset, end), payload: bytes.slice(sizeHeader.next, end) });
    if (records.length > limits.maxRecordCount) resource(`BIFF12 contains more than ${limits.maxRecordCount} records`);
    offset = end;
  }
  return records;
}

function biff12Record(type: number, payload: Uint8Array): Uint8Array {
  return concatBytes([varUint(type), varUint(payload.length), payload]);
}

function readWideString(bytes: Uint8Array, offset: number): { value: string; next: number } {
  if (offset + 4 > bytes.length) invalid('BIFF12 wide string header is truncated');
  const count = dataView(bytes).getUint32(offset, true);
  const end = offset + 4 + count * 2;
  if (end > bytes.length) invalid('BIFF12 wide string payload is truncated');
  return { value: new TextDecoder('utf-16le').decode(bytes.slice(offset + 4, end)), next: end };
}

function readBiff12Cell(record: BinaryRecord, row: number, recordIndex: number, sharedStrings: readonly string[], lastColumn: number): { cell?: BinaryCellGraph; column: number } {
  const isLong = XLSB_LONG_CELL_TYPES.has(record.type);
  const isShort = XLSB_SHORT_CELL_TYPES.has(record.type);
  if (!isLong && !isShort) return { column: lastColumn };
  const base = isShort ? 0 : 4;
  if (record.payload.length < base + 4) invalid(`BIFF12 cell record ${record.type} is truncated`);
  const column = isShort ? lastColumn + 1 : dataView(record.payload).getUint32(0, true);
  const style = record.payload[base]! | (record.payload[base + 1]! << 8) | (record.payload[base + 2]! << 16);
  const data = record.payload.slice(base + 4);
  let value: CellValue = null;
  if (record.type === XLSB.CELLREAL || record.type === XLSB.SHORTREAL) {
    if (data.length < 8) invalid('BIFF12 real cell is truncated');
    value = dataView(data).getFloat64(0, true);
  } else if (record.type === XLSB.CELLRK || record.type === XLSB.SHORTRK) {
    if (data.length < 4) invalid('BIFF12 RK cell is truncated');
    value = decodeRk(dataView(data).getUint32(0, true));
  } else if (record.type === XLSB.CELLBOOL || record.type === XLSB.SHORTBOOL) value = Boolean(data[0]);
  else if (record.type === XLSB.CELLERROR || record.type === XLSB.SHORTERROR) value = `#ERR${data[0] ?? 0}`;
  else if (record.type === XLSB.CELLST || record.type === XLSB.SHORTST) value = readWideString(data, 0).value;
  else if (record.type === XLSB.CELLISST || record.type === XLSB.SHORTISST) {
    if (data.length < 4) invalid('BIFF12 shared string cell is truncated');
    const index = dataView(data).getUint32(0, true);
    if (index >= sharedStrings.length) invalid(`BIFF12 shared string index is out of range: ${index}`);
    value = sharedStrings[index] ?? '';
  } else if (record.type === XLSB.FMLANUM) {
    if (data.length < 8) invalid('BIFF12 numeric formula result is truncated');
    value = dataView(data).getFloat64(0, true);
  } else if (record.type === XLSB.FMLABOOL) value = Boolean(data[0]);
  else if (record.type === XLSB.FMLAERROR) value = `#ERR${data[0] ?? 0}`;
  else if (record.type === XLSB.FMLASTRING) value = readWideString(data, 0).value;
  return { cell: { row, column, recordIndex, recordType: record.type, value, styleIndex: style }, column };
}

function readBinaryPackage(bytes: Uint8Array, limits: NativeDocumentResourceLimits): Record<string, Uint8Array> {
  assertInput(bytes, limits, 'XLSB package');
  let count = 0;
  let total = 0;
  const files = unzipSync(bytes, {
    filter(file) {
      count += 1;
      total += file.originalSize;
      if (count > limits.maxEntries || file.originalSize > limits.maxEntryBytes || total > limits.maxUncompressedBytes || file.originalSize > limits.maxCompressionRatio * Math.max(file.size, 1)) resource('XLSB package exceeds native resource limits');
      if (file.name.endsWith('/') || file.name.includes('\0') || file.name.split('/').includes('..')) invalid(`XLSB part name is unsafe: ${file.name}`);
      return true;
    },
  });
  const result: Record<string, Uint8Array> = {};
  for (const [name, value] of Object.entries(files)) result[name] = new Uint8Array(value);
  return result;
}

function packageRelationships(parts: Record<string, Uint8Array>): Record<string, import('./types').NativeRelationship[]> {
  const result: Record<string, import('./types').NativeRelationship[]> = {};
  for (const [name, bytes] of Object.entries(parts)) {
    if (!name.endsWith('.rels')) continue;
    const marker = name.indexOf('/_rels/');
    const source = marker >= 0 ? `${name.slice(0, marker)}/${name.slice(marker + 7, -5)}` : '';
    const root = parseXml(strFromU8(bytes));
    result[source] = children(root.children[0], 'Relationship').map((node) => ({ id: node.attrs.Id ?? '', type: node.attrs.Type ?? '', target: node.attrs.Target ?? '', ...(node.attrs.TargetMode ? { targetMode: node.attrs.TargetMode } : {}) }));
  }
  return result;
}

function normalizePackagePath(path: string): string {
  const result: string[] = [];
  for (const part of path.replaceAll('\\', '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      if (!result.length) invalid(`Package path escapes root: ${path}`);
      result.pop();
    } else result.push(part);
  }
  return result.join('/');
}

function parseXlsbDocument(bytes: Uint8Array, fileName: string, limits: NativeDocumentResourceLimits): BinaryParseResult {
  const parts = readBinaryPackage(bytes, limits);
  const workbookPart = Object.keys(parts).find((name) => /(^|\/)workbook\.bin$/i.test(name));
  if (!workbookPart) invalid('XLSB workbook.bin part is missing');
  const workbookRecords = parseBiff12Records(parts[workbookPart!]!, limits);
  const relationships = packageRelationships(parts);
  const bundles = workbookRecords.map((record, index) => ({ record, index })).filter(({ record }) => record.type === XLSB.BUNDLESHEET);
  const descriptors = bundles.map(({ record }, index) => {
    if (record.payload.length < 8) invalid('XLSB BrtBundleSh record is truncated');
    const hidden = dataView(record.payload).getUint32(0, true) !== 0;
    const relationId = readWideString(record.payload, 8);
    const name = readWideString(record.payload, relationId.next).value;
    return { name: name || `Sheet${index + 1}`, hidden, type: 'worksheet' as const, relId: relationId.value };
  });
  if (!descriptors.length) invalid('XLSB workbook has no BrtBundleSh records');
  const sharedPart = Object.keys(parts).find((name) => /sharedStrings\.bin$/i.test(name));
  const sharedStrings = sharedPart ? parseBiff12Records(parts[sharedPart]!, limits).filter((record) => record.type === XLSB.SSTITEM).map((record) => readWideString(record.payload, 0).value) : [];
  const sheetGraphs: BinarySheetGraph[] = [];
  const features = new Set<string>(['cells', 'xlsb']);
  if (sharedStrings.length) features.add('sharedStrings');
  descriptors.forEach((descriptor, sheetIndex) => {
    const relation = (relationships[workbookPart!] ?? []).find((entry) => entry.id === descriptor.relId);
    if (!relation) invalid(`XLSB sheet relationship is missing: ${descriptor.relId}`);
    const base = workbookPart!.includes('/') ? workbookPart!.slice(0, workbookPart!.lastIndexOf('/') + 1) : '';
    const part = normalizePackagePath(`${base}${relation!.target}`);
    if (!parts[part]) invalid(`XLSB worksheet part is missing: ${part}`);
    const records = parseBiff12Records(parts[part]!, limits);
    let row = 0;
    let lastColumn = -1;
    const cells: Record<string, BinaryCellGraph> = {};
    for (let index = 0; index < records.length; index += 1) {
      const record = records[index]!;
      if (record.type === XLSB.ROWHDR) {
        if (record.payload.length < 4) invalid(`XLSB BrtRowHdr record is truncated in ${part}`);
        row = dataView(record.payload).getUint32(0, true);
        lastColumn = -1;
      }
      const parsed = readBiff12Cell(record, row, index, sharedStrings, lastColumn);
      if (parsed.cell) {
        cells[cellKey(parsed.cell.row, parsed.cell.column)] = parsed.cell;
        lastColumn = parsed.column;
        if (XLSB_FORMULA_TYPES.has(parsed.cell.recordType)) features.add('formulas');
      }
    }
    sheetGraphs.push({ name: descriptor.name, hidden: descriptor.hidden, type: descriptor.type, part, startRecord: 0, endRecord: records.length - 1, cells });
    void sheetIndex;
  });
  const workbook = new WorkbookModel(`imported-xlsb-${Date.now().toString(36)}`, fileName.replace(/\.[^.]+$/, ''));
  const first = workbook.getSheet(workbook.primarySheetId);
  sheetGraphs.forEach((descriptor, index) => {
    const sheet = index === 0 ? first : workbook.addSheet(`sheet-${index + 1}`, descriptor.name);
    sheet.name = descriptor.name;
    sheet.hidden = descriptor.hidden;
    for (const cell of Object.values(descriptor.cells)) if (cell.value !== null) sheet.cells.set(cell.row, cell.column, { value: cell.value });
  });
  const snapshot = workbook.snapshot();
  snapshot.name = fileName.replace(/\.[^.]+$/, '');
  const graph: BinaryRecordGraph = {
    container: 'biff12',
    records: workbookRecords.map((record) => ({ type: record.type, offset: record.offset, bytes: record.bytes.slice() })),
    opaque: new Uint8Array(),
    streamName: workbookPart,
    streams: Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, value.slice()])),
    package: {
      parts: Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, value.slice()])),
      workbookPart: workbookPart!,
      worksheetParts: Object.fromEntries(sheetGraphs.map((sheet) => [sheet.name, sheet.part!])),
      relationships,
      contentTypesXml: parts['[Content_Types].xml']?.slice(),
    },
    sheets: sheetGraphs,
    sharedStrings,
    dateSystem: '1900',
  };
  return { snapshot, graph: { kind: 'xlsb', container: graph }, format: { family: 'xlsb', variant: 'xlsb' }, dateSystem: '1900', features: [...features] };
}

function untouched(request: NativeDocumentExportTransaction): NativeDocumentExportResult | undefined {
  if (!request.artifact || request.artifact.fileName !== request.fileName || request.artifact.sourceSnapshotHash !== nativeSnapshotHash(request.snapshot)) return undefined;
  return { taskId: `export-${Date.now().toString(36)}`, report: structuredClone(request.artifact.compatibility), buffer: request.artifact.sourceBytes.slice(0), fileName: request.fileName, artifact: request.artifact };
}

function nativeImportResult(fileName: string, bytes: Uint8Array, parsed: BinaryParseResult, options: NativeDocumentImportOptionsLike): Promise<NativeDocumentImportResult> {
  const report = createCompatibilityReport({ fileName, importLevel: options.compatibilityTarget, exportLevel: options.compatibilityTarget, dateSystem: parsed.dateSystem, detectedFeatures: parsed.features, editableFeatures: ['cells'], preservedFeatures: [parsed.format.family], projectedFeatures: [] });
  return createNativeDocumentArtifact({ fileName, buffer: toBuffer(bytes), dateSystem: parsed.dateSystem, format: parsed.format, nativeGraph: parsed.graph, snapshot: parsed.snapshot, detectedFeatures: parsed.features, compatibility: report }).then((artifact) => ({ payload: { name: parsed.snapshot.name, sheetCount: parsed.snapshot.sheets.length, dateSystem: parsed.dateSystem, compatibilityLevel: options.compatibilityTarget }, report, snapshot: parsed.snapshot, artifact, taskId: `import-${Date.now().toString(36)}` }));
}

function nativeExportResult(fileName: string, bytes: Uint8Array, snapshot: WorkbookSnapshot, parsed: BinaryParseResult, options: { compatibilityTarget: 'A' | 'B' | 'C' }): Promise<NativeDocumentExportResult> {
  const report = createCompatibilityReport({ fileName, importLevel: options.compatibilityTarget, exportLevel: options.compatibilityTarget, dateSystem: parsed.dateSystem, detectedFeatures: parsed.features, editableFeatures: ['cells'], preservedFeatures: [parsed.format.family], projectedFeatures: [] });
  return createNativeDocumentArtifact({ fileName, buffer: toBuffer(bytes), dateSystem: parsed.dateSystem, format: parsed.format, nativeGraph: parsed.graph, snapshot, detectedFeatures: parsed.features, compatibility: report }).then((artifact) => ({ taskId: `export-${Date.now().toString(36)}`, report, buffer: toBuffer(bytes), fileName, artifact }));
}

function cellHeader12(column: number, style: number): Uint8Array {
  if (!Number.isInteger(column) || column < 0 || column > 0xffffffff) invalid(`BIFF12 cell column is invalid: ${column}`);
  const payload = new Uint8Array(8);
  dataView(payload).setUint32(0, column, true);
  payload[4] = style & 0xff;
  payload[5] = (style >>> 8) & 0xff;
  payload[6] = (style >>> 16) & 0xff;
  payload[7] = 0;
  return payload;
}

function floatBytes(value: number): Uint8Array {
  const bytes = new Uint8Array(8);
  dataView(bytes).setFloat64(0, value, true);
  return bytes;
}

function wideString(value: string): Uint8Array {
  const chars = value;
  const bytes = new Uint8Array(4 + chars.length * 2);
  dataView(bytes).setUint32(0, chars.length, true);
  for (let index = 0; index < chars.length; index += 1) dataView(bytes).setUint16(4 + index * 2, chars[index]!.charCodeAt(0), true);
  return bytes;
}

function encodeBiff12Cell(column: number, style: number, value: CellValue): Uint8Array {
  const header = cellHeader12(column, style);
  if (value === null) return biff12Record(XLSB.CELLBLANK, header);
  if (typeof value === 'number') return biff12Record(XLSB.CELLREAL, concatBytes([header, floatBytes(value)]));
  if (typeof value === 'boolean') return biff12Record(XLSB.CELLBOOL, concatBytes([header, Uint8Array.of(value ? 1 : 0)]));
  return biff12Record(XLSB.CELLST, concatBytes([header, wideString(String(value))]));
}

function findXlsbRowInsertionIndex(records: readonly BinaryRecord[], row: number, start: number, end: number): number {
  let rowHeader = -1;
  for (let index = start; index <= end; index += 1) {
    if (records[index]!.type !== XLSB.ROWHDR) continue;
    if (records[index]!.payload.length < 4) invalid('XLSB row header is truncated');
    const current = dataView(records[index]!.payload).getUint32(0, true);
    if (current === row) rowHeader = index;
    else if (rowHeader >= 0) return index;
  }
  return rowHeader >= 0 ? Math.min(end, rowHeader + 1) : -1;
}

function rewriteXlsbParts(snapshot: WorkbookSnapshot, graph: BinaryRecordGraph, limits: NativeDocumentResourceLimits): Record<string, Uint8Array> {
  if (graph.container !== 'biff12' || !graph.package || !graph.sheets) invalid('XLSB artifact does not contain a writable BIFF12 graph');
  if (snapshot.sheets.length < graph.sheets.length) unsupported('The workbook lost a native XLSB sheet during Save', 'Keep every imported sheet in the workbook before saving.');
  const parts = Object.fromEntries(Object.entries(graph.package.parts).map(([name, bytes]) => [name, bytes.slice()]));
  for (let sheetIndex = 0; sheetIndex < graph.sheets.length; sheetIndex += 1) {
    const nativeSheet = graph.sheets[sheetIndex]!;
    const sheet = snapshot.sheets[sheetIndex];
    if (!sheet || !nativeSheet.part || !parts[nativeSheet.part]) invalid(`XLSB worksheet part is missing: ${nativeSheet.part ?? nativeSheet.name}`);
    const records = parseBiff12Records(parts[nativeSheet.part!]!, limits);
    const replacements = new Map<number, Uint8Array>();
    const inserts = new Map<number, Uint8Array[]>();
    const seen = new Set<string>();
    for (const binding of Object.values(nativeSheet.cells)) {
      const key = cellKey(binding.row, binding.column);
      seen.add(key);
      const current = snapshotCellValue(sheet!, binding.row, binding.column);
      if (sameCellValue(current, binding.value)) continue;
      if (XLSB_FORMULA_TYPES.has(binding.recordType)) unsupported(`Formula cell ${nativeSheet.name}!${binding.row}:${binding.column} cannot be rewritten from a value-only snapshot`, 'Preserve the formula expression or leave the formula cell unchanged.');
      replacements.set(binding.recordIndex, encodeBiff12Cell(binding.column, binding.styleIndex ?? 0, current));
    }
    for (const entry of snapshotCellEntries(sheet!)) {
      if (seen.has(cellKey(entry.row, entry.column)) || entry.value === null) continue;
      const insertionIndex = findXlsbRowInsertionIndex(records, entry.row, nativeSheet.startRecord, nativeSheet.endRecord);
      if (insertionIndex < 0) unsupported(`XLSB row ${entry.row} is not present in ${nativeSheet.name}`, 'Create the row in Excel before editing a binary worksheet.');
      const list = inserts.get(insertionIndex) ?? [];
      list.push(encodeBiff12Cell(entry.column, 0, entry.value));
      inserts.set(insertionIndex, list);
    }
    const output: Uint8Array[] = [];
    for (let index = 0; index < records.length; index += 1) {
      for (const inserted of inserts.get(index) ?? []) output.push(inserted);
      output.push(replacements.get(index) ?? records[index]!.bytes);
    }
    parts[nativeSheet.part!] = new Uint8Array(concatBytes(output));
  }
  for (let index = graph.sheets.length; index < snapshot.sheets.length; index += 1) {
    if (snapshotCellEntries(snapshot.sheets[index]!).some((entry) => entry.value !== null)) unsupported('New XLSB sheets are not representable by the imported native graph', 'Edit existing native sheets or create a new workbook in a format with sheet creation support.');
  }
  return parts;
}

function binaryFormatForName(fileName: string, bytes: Uint8Array): Extract<NativeDocumentFormat, { family: 'biff' | 'xlsb' }> {
  const lower = fileName.toLowerCase();
  if (!isCfb(bytes) && (lower.endsWith('.xlsb') || (bytes[0] === 0x50 && bytes[1] === 0x4b))) return { family: 'xlsb', variant: 'xlsb' };
  if (lower.endsWith('.xlt')) return { family: 'biff', variant: 'xlt' };
  if (lower.endsWith('.xla')) return { family: 'biff', variant: 'xla' };
  if (lower.endsWith('.xlw')) return { family: 'biff', variant: 'xlw' };
  return { family: 'biff', variant: isCfb(bytes) ? 'xls' : 'biff5' };
}

export function detectBinaryDocumentFormat(fileName: string, bytes: ArrayBuffer): Extract<NativeDocumentFormat, { family: 'biff' | 'xlsb' }> {
  return binaryFormatForName(fileName, new Uint8Array(bytes));
}

export const biffCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'biff',
  canRead: (fileName, buffer) => {
    const bytes = new Uint8Array(buffer);
    return bytes.length === 0 ? /\.(?:xls|xlt|xla|xlw)$/i.test(fileName) : isCfb(bytes);
  },
  import: async (request) => {
    const bytes = new Uint8Array(request.buffer);
    return nativeImportResult(request.fileName, bytes, parseBiffDocument(bytes, request.fileName, limitsFor(request.options)), request.options);
  },
  export: async (request) => {
    const stable = untouched(request);
    if (stable) return stable;
    const artifact = request.artifact;
    if (!artifact || artifact.nativeGraph.kind !== 'biff') unsupported('BIFF Save requires the original CFB artifact', 'Open the original .xls/.xlt/.xla document before saving.');
    const graph = artifact!.nativeGraph.container;
    const limits = limitsFor(request.options);
    const variant = artifact!.format.family === 'biff' ? artifact!.format.variant : 'xls';
    const stream = rewriteBiffStream(graph, request.snapshot, limits, variant);
    const streamName = graph.streamName!;
    const output = new Uint8Array(writeCfbPackage(graph.cfb!, graph.streams!, { ...graph.streams, [streamName]: stream }, limits));
    const parsed = parseBiffDocument(output, request.fileName, limits);
    return nativeExportResult(request.fileName, output, request.snapshot, parsed, request.options);
  },
};

export const xlsbCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'xlsb',
  canRead: (fileName, buffer) => {
    const bytes = new Uint8Array(buffer);
    if (bytes.length === 0) return /\.xlsb$/i.test(fileName);
    if (bytes.length < 2 || bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes.length > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxArchiveBytes) return false;
    try {
      const parts = readBinaryPackage(bytes, DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS);
      const types = parts['[Content_Types].xml'];
      return Boolean(types && /binary\.macroEnabled\.main|sheet\.binary/i.test(strFromU8(types)));
    } catch {
      return false;
    }
  },
  import: async (request) => {
    const bytes = new Uint8Array(request.buffer);
    return nativeImportResult(request.fileName, bytes, parseXlsbDocument(bytes, request.fileName, limitsFor(request.options)), request.options);
  },
  export: async (request) => {
    const stable = untouched(request);
    if (stable) return stable;
    const artifact = request.artifact;
    if (!artifact || artifact.nativeGraph.kind !== 'xlsb') unsupported('XLSB Save requires the original BIFF12 package artifact', 'Open the original .xlsb document before saving.');
    const limits = limitsFor(request.options);
    const parts = rewriteXlsbParts(request.snapshot, artifact!.nativeGraph.container, limits);
    const output = new Uint8Array(zipSync(parts, { level: 6 }));
    const parsed = parseXlsbDocument(output, request.fileName, limits);
    return nativeExportResult(request.fileName, output, request.snapshot, parsed, request.options);
  },
};
