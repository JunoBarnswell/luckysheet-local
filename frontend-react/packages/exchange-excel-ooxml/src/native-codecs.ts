import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { WorkbookModel, type WorkbookSnapshot, type CellValue } from '@react-sheets/core-model';
import { child, children, descendants, localName, parseXml, serializeXml, textContent, type XmlNode } from './xml';
import { createCompatibilityReport } from './compatibility-report';
import { createNativeDocumentArtifact, nativeSnapshotHash } from './native-document-artifact';
import { NativeDocumentError } from './native-document-error';
import type { NativeDocumentCodec, NativeDocumentExportTransaction, NativeDocumentImportTransaction } from './codec-registry';
import { DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS, type NativeDocumentExportResult, type NativeDocumentImportResult, type NativeDocumentFormat, type NativeGraph, type TextDialectGraph, type NativeDocumentResourceLimits } from './types';

const XMLSS_NAMESPACE = 'urn:schemas-microsoft-com:office:spreadsheet';
const ODS_MIMETYPE = 'application/vnd.oasis.opendocument.spreadsheet';
const TEXT_FEATURES = ['cells'];

function limitsFor(options: { limits?: Partial<NativeDocumentResourceLimits> }): NativeDocumentResourceLimits {
  return { ...DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS, ...(options.limits ?? {}) };
}

function assertInputBudget(bytes: Uint8Array, limits: NativeDocumentResourceLimits, format: string): void {
  if (bytes.byteLength > limits.maxArchiveBytes) {
    throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} input exceeds ${limits.maxArchiveBytes} bytes`, recovery: 'Reduce the document size or raise the explicit resource limit.' });
  }
}

function assertCellBudget(rows: ReadonlyArray<ReadonlyArray<unknown>>, limits: NativeDocumentResourceLimits, format: string): void {
  const count = rows.reduce((total, row) => total + row.length, 0);
  if (count > limits.maxCells) {
    throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} materializes ${count} cells, above the ${limits.maxCells} cell limit`, recovery: 'Reduce the worksheet materialization range or raise the explicit resource limit.' });
  }
}

function assertXmlTreeBudget(node: XmlNode, limits: NativeDocumentResourceLimits, format: string, depth = 0): void {
  if (depth > limits.maxXmlDepth) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} exceeds XML depth ${limits.maxXmlDepth}` });
  for (const childNode of node.children) assertXmlTreeBudget(childNode, limits, format, depth + 1);
}

function invalidNativeDocument(message: string): never {
  throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_INVALID', message });
}

function unzipNativePackage(bytes: Uint8Array, limits: NativeDocumentResourceLimits, format: string): Record<string, Uint8Array> {
  assertInputBudget(bytes, limits, format);
  let entries = 0;
  let total = 0;
  const names = new Set<string>();
  const files = unzipSync(bytes, {
    filter(file) {
      entries += 1;
      if (!file.name || file.name.endsWith('/') || file.name.includes('\0') || file.name.split('/').some((part) => part === '..')) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_INVALID', message: `${format} contains an unsafe part name` });
      if (names.has(file.name)) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_INVALID', message: `${format} contains duplicate part ${file.name}` });
      names.add(file.name);
      if (entries > limits.maxEntries) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} contains more than ${limits.maxEntries} parts` });
      if (file.originalSize > limits.maxEntryBytes) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} part ${file.name} exceeds ${limits.maxEntryBytes} bytes` });
      total += file.originalSize;
      if (total > limits.maxUncompressedBytes) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} exceeds ${limits.maxUncompressedBytes} uncompressed bytes` });
      if (file.originalSize > limits.maxCompressionRatio * Math.max(file.size, 1)) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `${format} part ${file.name} exceeds the compression ratio limit` });
      return true;
    },
  });
  return files as Record<string, Uint8Array>;
}

function detectZipParts(buffer: ArrayBuffer): Record<string, Uint8Array> | undefined {
  const bytes = new Uint8Array(buffer);
  if (bytes.byteLength === 0 || bytes.byteLength > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxArchiveBytes) return undefined;
  try {
    let entries = 0;
    let total = 0;
    return unzipSync(bytes, {
      filter(file) {
        entries += 1;
        total += file.originalSize;
        if (entries > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxEntries || file.originalSize > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxEntryBytes || total > DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxUncompressedBytes) throw new Error('detection budget exceeded');
        return true;
      },
    }) as Record<string, Uint8Array>;
  } catch {
    return undefined;
  }
}

function taskId(prefix: string): string {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function scalar(value: string): CellValue {
  const guarded = /^'[=+\-@]/.test(value) ? value.slice(1) : value;
  const normalized = guarded.trim();
  if (normalized === '') return null;
  if (normalized === 'TRUE') return true;
  if (normalized === 'FALSE') return false;
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(normalized)) {
    const number = Number(normalized);
    if (Number.isFinite(number)) return number;
  }
  return guarded;
}

function guardFormulaInjection(value: string): string {
  return /^[=+\-@]/.test(value) ? `'${value}` : value;
}

function workbookFromRows(name: string, rows: string[][], sheetName = 'Sheet1'): WorkbookSnapshot {
  const workbook = new WorkbookModel(`imported-${taskId('native')}`, name);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.name = sheetName || 'Sheet1';
  rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    const parsed = scalar(value);
    if (parsed !== null) sheet.cells.set(rowIndex, columnIndex, { value: parsed });
  }));
  const snapshot = workbook.snapshot();
  snapshot.name = name;
  return snapshot;
}

function rowsFromSnapshot(snapshot: WorkbookSnapshot): string[][] {
  const sheet = snapshot.sheets[0];
  if (!sheet) return invalidNativeDocument('NATIVE_TEXT_EXPORT_EMPTY: workbook has no active worksheet');
  const entries = Object.entries(sheet.cells);
  if (!entries.length) return [];
  const maxRow = entries.length ? Math.max(...entries.map(([row]) => Number(row))) : 0;
  const maxColumn = entries.length ? Math.max(...entries.flatMap(([, columns]) => Object.keys(columns).map(Number))) : 0;
  return Array.from({ length: maxRow + 1 }, (_, row) => Array.from({ length: maxColumn + 1 }, (_, column) => {
    const value = sheet.cells[String(row)]?.[String(column)]?.value;
    return value === null || value === undefined ? '' : String(value);
  }));
}

function decodeText(bytes: Uint8Array): { text: string; encoding: TextDialectGraph['encoding']; bom: boolean } {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(bytes.slice(2)), encoding: 'utf-16le', bom: true };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(bytes.slice(2)), encoding: 'utf-16be', bom: true };
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(bytes.slice(3)), encoding: 'utf-8', bom: true };
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'utf-8', bom: false };
  } catch {
    return { text: new TextDecoder('windows-1252').decode(bytes), encoding: 'windows-1252', bom: false };
  }
}

function textDialect(fileName: string, text: string, encoding: TextDialectGraph['encoding'], bom: boolean): TextDialectGraph {
  const lower = fileName.toLowerCase();
  const variant = textVariantForFileName(fileName);
  const rowDelimiter: TextDialectGraph['rowDelimiter'] = text.includes('\r\n') ? '\r\n' : text.includes('\r') ? '\r' : '\n';
  const delimiter: TextDialectGraph['delimiter'] = variant === 'txt' ? '\t' : variant === 'prn' ? ' ' : text.includes('\t') && !text.includes(',') ? '\t' : text.includes(';') && !text.includes(',') ? ';' : ',';
  return { encoding, bom, delimiter, rowDelimiter, quote: variant === 'prn' ? 'none' : 'double', variant };
}

function textVariantForFileName(fileName: string): TextDialectGraph['variant'] {
  const lower = fileName.toLowerCase();
  return lower.endsWith('.dif') ? 'dif' : lower.endsWith('.slk') ? 'sylk' : lower.endsWith('.prn') ? 'prn' : lower.endsWith('.txt') ? 'txt' : 'csv';
}

function parseDelimited(text: string, dialect: TextDialectGraph): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  const pushField = () => { row.push(field); field = ''; };
  const pushRow = () => { pushField(); rows.push(row); row = []; };
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]!;
    if (dialect.quote === 'double' && character === '"') {
      if (quoted && text[index + 1] === '"') { field += '"'; index += 1; }
      else quoted = !quoted;
      continue;
    }
    if (!quoted && character === dialect.delimiter) { pushField(); continue; }
    if (!quoted && ((dialect.rowDelimiter === '\r\n' && character === '\r' && text[index + 1] === '\n') || character === dialect.rowDelimiter)) {
      if (dialect.rowDelimiter === '\r\n') index += 1;
      pushRow();
      continue;
    }
    field += character;
  }
  if (quoted) invalidNativeDocument('NATIVE_TEXT_INVALID: unterminated quoted field');
  if (field.length > 0 || row.length > 0 || text.endsWith(dialect.delimiter)) pushRow();
  return rows;
}

function serializeDelimited(rows: string[][], dialect: TextDialectGraph): Uint8Array {
  const delimiter = dialect.delimiter;
  const text = rows.map((row) => row.map((value) => {
    const safeValue = guardFormulaInjection(value);
    if (dialect.quote === 'none') return safeValue;
    return safeValue.includes(delimiter) || safeValue.includes('\n') || safeValue.includes('\r') || safeValue.includes('"') ? `"${safeValue.replaceAll('"', '""')}"` : safeValue;
  }).join(delimiter)).join(dialect.rowDelimiter);
  if (dialect.encoding === 'utf-8') {
    const encoded = new TextEncoder().encode(text);
    if (!dialect.bom) return encoded;
    const result = new Uint8Array(encoded.length + 3); result.set([0xef, 0xbb, 0xbf]); result.set(encoded, 3); return result;
  }
  if (dialect.encoding === 'utf-16le' || dialect.encoding === 'utf-16be') {
    const result = new Uint8Array(text.length * 2 + (dialect.bom ? 2 : 0));
    if (dialect.bom) result.set(dialect.encoding === 'utf-16le' ? [0xff, 0xfe] : [0xfe, 0xff]);
    const view = new DataView(result.buffer);
    const offset = dialect.bom ? 2 : 0;
    for (let index = 0; index < text.length; index += 1) view.setUint16(offset + index * 2, text.charCodeAt(index), dialect.encoding === 'utf-16le');
    return result;
  }
  const encoded = new Uint8Array(text.length);
  for (let index = 0; index < text.length; index += 1) encoded[index] = windows1252Byte(text.charCodeAt(index));
  return encoded;
}

function windows1252Byte(codePoint: number): number {
  if (codePoint <= 0xff) return codePoint;
  const substitutions: Record<number, number> = { 0x20ac: 0x80, 0x201a: 0x82, 0x192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x2c6: 0x88, 0x2030: 0x89, 0x160: 0x8a, 0x2039: 0x8b, 0x152: 0x8c, 0x17d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x2dc: 0x98, 0x2122: 0x99, 0x161: 0x9a, 0x203a: 0x9b, 0x153: 0x9c, 0x17e: 0x9e, 0x178: 0x9f };
  return substitutions[codePoint] ?? 0x3f;
}

function report(fileName: string, features: string[], compatibilityLevel: 'A' | 'B' | 'C' = 'B', editable = features.filter((feature) => feature !== 'dbf')): ReturnType<typeof createCompatibilityReport> {
  return createCompatibilityReport({ fileName, importLevel: compatibilityLevel, exportLevel: compatibilityLevel, dateSystem: '1900', detectedFeatures: features, editableFeatures: editable, preservedFeatures: features.includes('dbf') ? ['dbf'] : [], projectedFeatures: features.filter((feature) => feature !== 'dbf') });
}

async function importedResult(fileName: string, bytes: Uint8Array, format: NativeDocumentFormat, snapshot: WorkbookSnapshot, nativeGraph: NativeGraph, features: string[], compatibilityLevel: 'A' | 'B' | 'C' = 'B'): Promise<NativeDocumentImportResult> {
  const compatibility = report(fileName, features, compatibilityLevel);
  const artifact = await createNativeDocumentArtifact({ fileName, buffer: asArrayBuffer(bytes), dateSystem: '1900', format, nativeGraph, snapshot, detectedFeatures: features, compatibility });
  return { payload: { name: snapshot.name, sheetCount: snapshot.sheets.length, dateSystem: '1900', compatibilityLevel }, report: compatibility, snapshot, artifact, taskId: taskId('import') };
}

async function exportedResult(fileName: string, bytes: Uint8Array, format: NativeDocumentFormat, snapshot: WorkbookSnapshot, nativeGraph: NativeGraph, features: string[], compatibilityLevel: 'A' | 'B' | 'C' = 'B'): Promise<NativeDocumentExportResult> {
  const compatibility = report(fileName, features, compatibilityLevel);
  const artifact = await createNativeDocumentArtifact({ fileName, buffer: asArrayBuffer(bytes), dateSystem: '1900', format, nativeGraph, snapshot, detectedFeatures: features, compatibility });
  return { taskId: taskId('export'), report: compatibility, buffer: asArrayBuffer(bytes), fileName, artifact };
}

function untouchedExport(request: NativeDocumentExportTransaction): NativeDocumentExportResult | undefined {
  const artifact = request.artifact;
  if (!artifact || artifact.fileName !== request.fileName || artifact.sourceSnapshotHash !== nativeSnapshotHash(request.snapshot)) return undefined;
  return {
    taskId: taskId('export'),
    report: structuredClone(artifact.compatibility),
    buffer: artifact.sourceBytes.slice(0),
    fileName: request.fileName,
    artifact,
  };
}

function formatForText(dialect: TextDialectGraph): NativeDocumentFormat {
  return { family: 'text', variant: dialect.variant };
}

export const textCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'text',
  canRead: (fileName) => /\.(csv|txt|prn|dif|slk)$/i.test(fileName),
  import: async (request) => {
    const bytes = new Uint8Array(request.buffer);
    assertInputBudget(bytes, limitsFor(request.options), 'Text document');
    const decoded = decodeText(bytes);
    const dialect = textDialect(request.fileName, decoded.text, decoded.encoding, decoded.bom);
    let rows: string[][];
    if (dialect.variant === 'dif') rows = parseDif(decoded.text);
    else if (dialect.variant === 'sylk') rows = parseSylk(decoded.text);
    else rows = parseDelimited(decoded.text, dialect);
    assertCellBudget(rows, limitsFor(request.options), 'Text document');
    return importedResult(request.fileName, bytes, formatForText(dialect), workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), rows), { kind: 'text', dialect }, TEXT_FEATURES, request.options.compatibilityTarget);
  },
  export: async (request) => {
    const untouched = untouchedExport(request);
    if (untouched) return untouched;
    const sourceDialect = request.artifact?.nativeGraph.kind === 'text' ? request.artifact.nativeGraph.dialect : undefined;
    const dialect: TextDialectGraph = sourceDialect && sourceDialect.variant === textVariantForFileName(request.fileName) ? sourceDialect : textDialect(request.fileName, '', 'utf-8', false);
    const rows = rowsFromSnapshot(request.snapshot);
    assertCellBudget(rows, limitsFor(request.options), 'Text document');
    const bytes = dialect.variant === 'dif' ? serializeDif(rows) : dialect.variant === 'sylk' ? serializeSylk(rows) : serializeDelimited(rows, dialect);
    return exportedResult(request.fileName, bytes, formatForText(dialect), request.snapshot, { kind: 'text', dialect }, TEXT_FEATURES, request.options.compatibilityTarget);
  },
};

function parseDif(text: string): string[][] {
  const lines = text.split(/\r?\n/);
  const dataStart = lines.findIndex((line) => line.trim().toUpperCase() === 'DATA');
  if (dataStart < 0) invalidNativeDocument('NATIVE_DIF_INVALID: DATA section is missing');
  const vectorsIndex = lines.findIndex((line) => line.trim().toUpperCase() === 'VECTORS');
  const width = vectorsIndex >= 0 ? Number((lines[vectorsIndex + 1] ?? '').split(',')[1]) : 0;
  if (!Number.isSafeInteger(width) || width < 0) invalidNativeDocument('NATIVE_DIF_INVALID: VECTORS width is invalid');
  const values: string[] = [];
  for (let index = dataStart + 1; index < lines.length; index += 1) {
    const typeLine = lines[index]?.trim() ?? '';
    if (!typeLine) continue;
    if (typeLine.toUpperCase() === 'EOD') break;
    const [typeText] = typeLine.split(',', 1);
    const type = Number(typeText);
    if (!Number.isFinite(type)) invalidNativeDocument(`NATIVE_DIF_INVALID: invalid data type at line ${index + 1}`);
    if (type === -1) break;
    const valueLine = lines[++index];
    if (valueLine === undefined) invalidNativeDocument('NATIVE_DIF_INVALID: missing value line');
    const value = valueLine.replace(/^"|"$/g, '').replaceAll('""', '"');
    values.push(type === 2 ? (value === '1' || value.toUpperCase() === 'TRUE' ? 'TRUE' : 'FALSE') : value);
  }
  if (width === 0) return [];
  return Array.from({ length: Math.ceil(values.length / width) }, (_, rowIndex) => values.slice(rowIndex * width, (rowIndex + 1) * width));
}

function serializeDif(rows: string[][]): Uint8Array {
  const width = Math.max(0, ...rows.map((row) => row.length));
  const lines = [`TABLE\r\n0,1\r\n"React Sheets"\r\nVECTORS\r\n0,${width}\r\n""\r\nTUPLES\r\n0,${rows.length}\r\n""\r\nDATA\r\n`];
  rows.forEach((row) => row.forEach((value) => {
    const numeric = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value.trim());
    const safeValue = guardFormulaInjection(value);
    lines.push(`${numeric ? '0,0' : '1,0'}\r\n"${safeValue.replaceAll('"', '""')}"\r\n`);
  }));
  lines.push('EOD\r\n');
  return strToU8(lines.join(''));
}

function parseSylk(text: string): string[][] {
  const rows: string[][] = [];
  for (const line of text.split(/\r?\n/)) {
    if (!line.startsWith('C;')) continue;
    const row = Number(line.match(/Y(\d+)/)?.[1] ?? 1) - 1;
    const column = Number(line.match(/X(\d+)/)?.[1] ?? 1) - 1;
    const value = line.match(/;K(.*)$/)?.[1] ?? '';
    rows[row] ??= [];
    rows[row]![column] = value.replace(/^"|"$/g, '');
  }
  return rows.map((row) => row.map((value) => value ?? ''));
}

function serializeSylk(rows: string[][]): Uint8Array {
  const lines = ['ID;PReact Sheets\r\n'];
  rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => lines.push(`C;Y${rowIndex + 1};X${columnIndex + 1};K${guardFormulaInjection(value)}\r\n`)));
  lines.push('E\r\n');
  return strToU8(lines.join(''));
}

function xmlValue(node: XmlNode): string {
  const data = child(node, 'data');
  return data ? textContent(data) : textContent(node);
}

function parseXmlss(bytes: Uint8Array, limits: NativeDocumentResourceLimits): { rows: string[][]; graph: NativeGraph } {
  if (bytes.byteLength > limits.maxXmlBytes) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `XML Spreadsheet exceeds ${limits.maxXmlBytes} XML bytes` });
  const root = parseXml(strFromU8(bytes));
  assertXmlTreeBudget(root, limits, 'XML Spreadsheet');
  const workbook = descendants(root, 'Workbook').find((node) => node.attrs.xmlns === XMLSS_NAMESPACE || node.name.includes('Workbook'));
  if (!workbook) invalidNativeDocument('NATIVE_XMLSS_INVALID: Workbook root is missing');
  const worksheet = descendants(workbook, 'Worksheet')[0];
  if (!worksheet) invalidNativeDocument('NATIVE_XMLSS_INVALID: Worksheet is missing');
  const rows: string[][] = [];
  for (const rowNode of descendants(worksheet, 'Row')) {
    const row: string[] = [];
    let column = 0;
    for (const cellNode of children(rowNode, 'Cell')) {
      const index = Number(cellNode.attrs['ss:Index'] ?? cellNode.attrs.Index ?? column + 1) - 1;
      while (column < index) row.push(''), column += 1;
      row.push(xmlValue(cellNode)); column += 1;
    }
    rows.push(row);
  }
  assertCellBudget(rows, limits, 'XML Spreadsheet');
  return { rows, graph: { kind: 'xml', root: { namespace: XMLSS_NAMESPACE, root: 'Workbook', encoding: 'utf-8', parsed: root } } };
}

function serializeXmlss(snapshot: WorkbookSnapshot): Uint8Array {
  const rows = rowsFromSnapshot(snapshot);
  const rowXml = rows.map((row) => `<ss:Row>${row.map((value) => `<ss:Cell><ss:Data ss:Type="${typeof scalar(value) === 'number' ? 'Number' : 'String'}">${escapeXml(value)}</ss:Data></ss:Cell>`).join('')}</ss:Row>`).join('');
  return strToU8(`<?xml version="1.0" encoding="UTF-8"?><ss:Workbook xmlns:ss="${XMLSS_NAMESPACE}"><ss:Worksheet ss:Name="${escapeXml(snapshot.sheets[0]?.name ?? 'Sheet1')}"><ss:Table>${rowXml}</ss:Table></ss:Worksheet></ss:Workbook>`);
}

function updateXmlssGraph(snapshot: WorkbookSnapshot, graph: Extract<NativeGraph, { kind: 'xml' }>): Uint8Array | undefined {
  if (!graph.root.parsed) return undefined;
  const root = structuredClone(graph.root.parsed);
  const worksheet = descendants(root, 'Worksheet')[0];
  const rowNodes = worksheet ? descendants(worksheet, 'Row') : [];
  const values = rowsFromSnapshot(snapshot);
  for (let rowIndex = 0; rowIndex < rowNodes.length; rowIndex += 1) {
    const cells = children(rowNodes[rowIndex], 'Cell');
    let column = 0;
    for (const cellNode of cells) {
      const index = Number(cellNode.attrs['ss:Index'] ?? cellNode.attrs.Index ?? column + 1) - 1;
      column = index;
      const data = child(cellNode, 'Data');
      if (data && values[rowIndex]?.[column] !== undefined) data.text = String(values[rowIndex]![column]!);
      column += 1;
    }
  }
  return strToU8(`<?xml version="1.0" encoding="UTF-8"?>${serializeXml(root)}`);
}

function escapeXml(value: string): string { return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;'); }

export const xmlssCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'xmlss',
  canRead: (fileName, buffer) => /\.xml$/i.test(fileName) && (buffer.byteLength === 0 || strFromU8(new Uint8Array(buffer).slice(0, 1024)).includes(XMLSS_NAMESPACE)),
  import: async (request) => { const bytes = new Uint8Array(request.buffer); const parsed = parseXmlss(bytes, limitsFor(request.options)); return importedResult(request.fileName, bytes, { family: 'xmlss', variant: 'xml' }, workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), parsed.rows), parsed.graph, TEXT_FEATURES, request.options.compatibilityTarget); },
  export: async (request) => { const untouched = untouchedExport(request); if (untouched) return untouched; const rows = rowsFromSnapshot(request.snapshot); assertCellBudget(rows, limitsFor(request.options), 'XML Spreadsheet'); const graph = request.artifact?.nativeGraph.kind === 'xml' ? request.artifact.nativeGraph : undefined; const bytes = graph ? updateXmlssGraph(request.snapshot, graph) ?? serializeXmlss(request.snapshot) : serializeXmlss(request.snapshot); return exportedResult(request.fileName.replace(/\.[^.]+$/i, '.xml'), bytes, { family: 'xmlss', variant: 'xml' }, request.snapshot, { kind: 'xml', root: { namespace: XMLSS_NAMESPACE, root: 'Workbook', encoding: 'utf-8', parsed: graph?.root.parsed } }, TEXT_FEATURES, request.options.compatibilityTarget); },
};

function odfRows(content: string): string[][] {
  return odfRowsFromRoot(parseXml(content));
}

function odfRowsFromRoot(root: XmlNode): string[][] {
  const table = descendants(root, 'table').find((node) => localName(node.name) === 'table');
  if (!table) invalidNativeDocument('NATIVE_ODS_INVALID: table:table is missing');
  const rows: string[][] = [];
  for (const rowNode of table.children.filter((node) => localName(node.name) === 'table-row')) {
    const row: string[] = [];
    for (const cellNode of rowNode.children.filter((node) => localName(node.name) === 'table-cell')) {
      const repeated = Number(cellNode.attrs['table:number-columns-repeated'] ?? 1);
      const value = cellNode.attrs['office:value'] ?? textContent(cellNode);
      for (let index = 0; index < Math.max(1, repeated); index += 1) row.push(value);
    }
    rows.push(row);
  }
  return rows;
}

function serializeOds(snapshot: WorkbookSnapshot, existing?: Record<string, Uint8Array>): Uint8Array {
  const parts: Record<string, Uint8Array> = Object.fromEntries(Object.entries(existing ?? {}).map(([name, bytes]) => [name, bytes.slice()]));
  const rows = rowsFromSnapshot(snapshot);
  const rowXml = rows.map((row) => `<table:table-row>${row.map((value) => odsCellXml(value)).join('')}</table:table-row>`).join('');
  parts.mimetype = strToU8(ODS_MIMETYPE);
  parts['content.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="${escapeXml(snapshot.sheets[0]?.name ?? 'Sheet1')}">${rowXml}</table:table></office:spreadsheet></office:body></office:document-content>`);
  parts['META-INF/manifest.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/"/><manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/></manifest:manifest>`);
  return zipSync(parts, { level: 6 });
}

function odsCellXml(value: string): string {
  const parsed = scalar(value);
  const valueType = typeof parsed === 'number' ? 'float' : typeof parsed === 'boolean' ? 'boolean' : 'string';
  const officeValue = parsed === null ? '' : ` office:value="${escapeXml(String(parsed))}"`;
  return `<table:table-cell office:value-type="${valueType}"${officeValue}><text:p>${escapeXml(value)}</text:p></table:table-cell>`;
}

function serializeOdsWithContent(snapshot: WorkbookSnapshot, existing: Record<string, Uint8Array>, contentPart: string, content: Uint8Array): Uint8Array {
  const parts: Record<string, Uint8Array> = Object.fromEntries(Object.entries(existing).map(([name, bytes]) => [name, bytes.slice()]));
  parts.mimetype = strToU8(ODS_MIMETYPE);
  parts[contentPart] = content.slice();
  parts['META-INF/manifest.xml'] ??= strToU8(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/"/><manifest:file-entry manifest:media-type="text/xml" manifest:full-path="${escapeXml(contentPart)}"/></manifest:manifest>`);
  return zipSync(parts, { level: 6 });
}

function updateOdsContent(snapshot: WorkbookSnapshot, graph: Extract<NativeGraph, { kind: 'ods' }>): Uint8Array | undefined {
  if (!graph.package.contentTree) return undefined;
  const root = structuredClone(graph.package.contentTree);
  const table = descendants(root, 'table').find((node) => localName(node.name) === 'table');
  if (!table) return undefined;
  const values = rowsFromSnapshot(snapshot);
  const rowNodes = table.children.filter((node) => localName(node.name) === 'table-row');
  for (let rowIndex = 0; rowIndex < rowNodes.length; rowIndex += 1) {
    const cells = rowNodes[rowIndex]!.children.filter((node) => localName(node.name) === 'table-cell');
    let column = 0;
    for (const cell of cells) {
      const repeated = Math.max(1, Number(cell.attrs['table:number-columns-repeated'] ?? 1));
      for (let offset = 0; offset < repeated; offset += 1) {
        const value = values[rowIndex]?.[column + offset];
        if (value !== undefined) {
          const parsed = scalar(value);
          cell.attrs['office:value-type'] = typeof parsed === 'number' ? 'float' : typeof parsed === 'boolean' ? 'boolean' : 'string';
          if (parsed === null) delete cell.attrs['office:value'];
          else cell.attrs['office:value'] = String(parsed);
          const paragraph = descendants(cell, 'p')[0];
          if (paragraph) paragraph.text = value;
        }
      }
      column += repeated;
    }
  }
  return strToU8(`<?xml version="1.0" encoding="UTF-8"?>${serializeXml(root)}`);
}

export const odsCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'ods',
  canRead: (fileName, buffer) => { const parts = detectZipParts(buffer); return /\.ods$/i.test(fileName) || Boolean(parts?.mimetype && strFromU8(parts.mimetype).includes('opendocument')); },
  import: async (request) => { const bytes = new Uint8Array(request.buffer); const limits = limitsFor(request.options); const parts = unzipNativePackage(bytes, limits, 'ODS document'); const contentPart = parts['content.xml'] ? 'content.xml' : Object.keys(parts).find((name) => name.endsWith('/content.xml')) ?? ''; if (!contentPart) invalidNativeDocument('NATIVE_ODS_INVALID: content.xml is missing'); if (parts[contentPart]!.byteLength > limits.maxXmlBytes) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `ODS content exceeds ${limits.maxXmlBytes} XML bytes` }); const contentTree = parseXml(strFromU8(parts[contentPart]!)); assertXmlTreeBudget(contentTree, limits, 'ODS document'); const rows = odfRowsFromRoot(contentTree); assertCellBudget(rows, limits, 'ODS document'); const graph: NativeGraph = { kind: 'ods', package: { parts: Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, data.slice()])), mimetype: strFromU8(parts.mimetype ?? strToU8(ODS_MIMETYPE)), contentPart, contentTree } }; return importedResult(request.fileName, bytes, { family: 'ods', variant: 'ods' }, workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), rows), graph, TEXT_FEATURES, request.options.compatibilityTarget); },
  export: async (request) => { const untouched = untouchedExport(request); if (untouched) return untouched; const existingGraph = request.artifact?.nativeGraph.kind === 'ods' ? request.artifact.nativeGraph : undefined; const existing = existingGraph?.package.parts; const rows = rowsFromSnapshot(request.snapshot); assertCellBudget(rows, limitsFor(request.options), 'ODS document'); const updatedContent = existingGraph ? updateOdsContent(request.snapshot, existingGraph) : undefined; const bytes = updatedContent && existing ? serializeOdsWithContent(request.snapshot, existing, existingGraph.package.contentPart, updatedContent) : serializeOds(request.snapshot, existing); return exportedResult(request.fileName.replace(/\.[^.]+$/i, '.ods'), bytes, { family: 'ods', variant: 'ods' }, request.snapshot, { kind: 'ods', package: { parts: unzipNativePackage(bytes, limitsFor(request.options), 'ODS document'), mimetype: ODS_MIMETYPE, contentPart: existingGraph?.package.contentPart ?? 'content.xml', contentTree: updatedContent ? parseXml(strFromU8(updatedContent)) : undefined } }, TEXT_FEATURES, request.options.compatibilityTarget); },
};

function jsonRows(value: unknown): { rows: string[][]; sheetName: string; unknownFields: Record<string, unknown> } {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalidNativeDocument('NATIVE_JSON_INVALID: document root must be an object');
  const root = value as Record<string, unknown>;
  if (root.schema === 'WorkbookSnapshot') return { rows: rowsFromSnapshot(root as unknown as WorkbookSnapshot), sheetName: 'Sheet1', unknownFields: {} };
  const sheets = Array.isArray(root.sheets) ? root.sheets : [];
  const first = sheets[0] && typeof sheets[0] === 'object' ? sheets[0] as Record<string, unknown> : root;
  const data = jsonDataRows(first.data ?? first.dataTable);
  if (!data) invalidNativeDocument('NATIVE_JSON_INVALID: sheets[0].data must be a two-dimensional array or dataTable map');
  const unknownFields = Object.fromEntries(Object.entries(root).filter(([key]) => !['version', 'sheets', 'schema', 'name'].includes(key)));
  return { rows: data.map((row) => (row as unknown[]).map((cell) => cell === null || cell === undefined ? '' : String(cell))), sheetName: typeof first.name === 'string' ? first.name : 'Sheet1', unknownFields };
}

function jsonDataRows(value: unknown): string[][] | undefined {
  if (Array.isArray(value) && value.every((row) => Array.isArray(row))) {
    return value.map((row) => (row as unknown[]).map((cell) => jsonCellValue(cell)));
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const object = value as Record<string, unknown>;
  if (object.dataTable !== undefined) return jsonDataRows(object.dataTable);
  const rowEntries = Object.entries(object).filter(([key]) => /^\d+$/.test(key)).sort(([left], [right]) => Number(left) - Number(right));
  if (!rowEntries.length) return undefined;
  const rows: string[][] = [];
  for (const [rowKey, rawRow] of rowEntries) {
    const rowIndex = Number(rowKey);
    if (!Number.isSafeInteger(rowIndex) || rowIndex < 0 || !rawRow || typeof rawRow !== 'object' || Array.isArray(rawRow)) return undefined;
    const columns = Object.entries(rawRow as Record<string, unknown>).filter(([key]) => /^\d+$/.test(key)).sort(([left], [right]) => Number(left) - Number(right));
    const row: string[] = [];
    for (const [columnKey, cell] of columns) row[Number(columnKey)] = jsonCellValue(cell);
    rows[rowIndex] = row.map((cell) => cell ?? '');
  }
  return rows.map((row) => row ?? []);
}

function jsonCellValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && !Array.isArray(value)) {
    const cell = value as Record<string, unknown>;
    if ('value' in cell) return jsonCellValue(cell.value);
    if ('text' in cell) return jsonCellValue(cell.text);
  }
  return String(value);
}

function parseNativeJson(bytes: Uint8Array): unknown {
  try {
    return JSON.parse(strFromU8(bytes)) as unknown;
  } catch (error) {
    return invalidNativeDocument(`NATIVE_JSON_INVALID: invalid JSON (${error instanceof Error ? error.message : 'parse failed'})`);
  }
}

function serializeSpreadJson(snapshot: WorkbookSnapshot, unknownFields: Record<string, unknown>, sjs = false): Uint8Array {
  const root = { ...structuredClone(unknownFields), schema: sjs ? undefined : 'SSJSON', version: 1, name: snapshot.name, sheets: snapshot.sheets.map((sheet) => ({ name: sheet.name, data: rowsFromSnapshot({ ...snapshot, sheets: [sheet] }) })) };
  if (root.schema === undefined) delete (root as { schema?: string }).schema;
  return strToU8(JSON.stringify(root));
}

export const ssjsonCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'ssjson',
  canRead: (fileName, buffer) => /\.ssjson$/i.test(fileName) || (/\.json$/i.test(fileName) && strFromU8(new Uint8Array(buffer).slice(0, 512)).includes('"sheets"')),
  import: async (request) => { const bytes = new Uint8Array(request.buffer); assertInputBudget(bytes, limitsFor(request.options), 'SSJSON document'); const parsed = jsonRows(parseNativeJson(bytes)); assertCellBudget(parsed.rows, limitsFor(request.options), 'SSJSON document'); const snapshot = workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), parsed.rows, parsed.sheetName); return importedResult(request.fileName, bytes, { family: 'ssjson', variant: 'ssjson' }, snapshot, { kind: 'ssjson', document: { unknownFields: parsed.unknownFields } }, TEXT_FEATURES, request.options.compatibilityTarget); },
  export: async (request) => { const untouched = untouchedExport(request); if (untouched) return untouched; const rows = rowsFromSnapshot(request.snapshot); assertCellBudget(rows, limitsFor(request.options), 'SSJSON document'); const bytes = serializeSpreadJson(request.snapshot, request.artifact?.nativeGraph.kind === 'ssjson' ? request.artifact.nativeGraph.document.unknownFields : {}); return exportedResult(request.fileName.replace(/\.[^.]+$/i, '.ssjson'), bytes, { family: 'ssjson', variant: 'ssjson' }, request.snapshot, { kind: 'ssjson', document: { unknownFields: {} } }, TEXT_FEATURES, request.options.compatibilityTarget); },
};

export const sjsCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'sjs',
  canRead: (fileName, buffer) => /\.sjs$/i.test(fileName) || Object.keys(detectZipParts(buffer) ?? {}).some((name) => name.toLowerCase().endsWith('.json')),
  import: async (request) => { const bytes = new Uint8Array(request.buffer); const limits = limitsFor(request.options); const parts = unzipNativePackage(bytes, limits, 'SJS document'); const workbookPart = Object.keys(parts).find((name) => /workbook.*\.json$/i.test(name)) ?? Object.keys(parts).find((name) => name.endsWith('.json')); if (!workbookPart) invalidNativeDocument('NATIVE_SJS_INVALID: workbook JSON part is missing'); const parsed = jsonRows(parseNativeJson(parts[workbookPart]!)); assertCellBudget(parsed.rows, limits, 'SJS document'); const snapshot = workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), parsed.rows, parsed.sheetName); return importedResult(request.fileName, bytes, { family: 'sjs', variant: 'sjs' }, snapshot, { kind: 'sjs', package: { parts: Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, data.slice()])), workbookPart, unknownParts: Object.fromEntries(Object.entries(parts).filter(([name]) => name !== workbookPart).map(([name, data]) => [name, data.slice()])) , unknownFields: parsed.unknownFields } }, TEXT_FEATURES, request.options.compatibilityTarget); },
  export: async (request) => { const untouched = untouchedExport(request); if (untouched) return untouched; const existing = request.artifact?.nativeGraph.kind === 'sjs' ? request.artifact.nativeGraph.package.parts : {}; const parts: Record<string, Uint8Array> = Object.fromEntries(Object.entries(existing).map(([name, bytes]) => [name, bytes.slice()])); const workbookPart = request.artifact?.nativeGraph.kind === 'sjs' ? request.artifact.nativeGraph.package.workbookPart : 'workbook.json'; const unknownFields = request.artifact?.nativeGraph.kind === 'sjs' ? request.artifact.nativeGraph.package.unknownFields : {}; parts[workbookPart] = serializeSpreadJson(request.snapshot, unknownFields, true) as Uint8Array; const bytes = zipSync(parts, { level: 6 }); const rows = rowsFromSnapshot(request.snapshot); assertCellBudget(rows, limitsFor(request.options), 'SJS document'); return exportedResult(request.fileName.replace(/\.[^.]+$/i, '.sjs'), bytes, { family: 'sjs', variant: 'sjs' }, request.snapshot, { kind: 'sjs', package: { parts, workbookPart, unknownParts: Object.fromEntries(Object.entries(parts).filter(([name]) => name !== workbookPart)), unknownFields } }, TEXT_FEATURES, request.options.compatibilityTarget); },
};

function parseDbf(bytes: Uint8Array, limits: NativeDocumentResourceLimits): { rows: string[][]; graph: NativeGraph } {
  if (bytes.byteLength < 32) return invalidNativeDocument('NATIVE_DBF_INVALID: header is truncated');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const version = view.getUint8(0);
  const recordCount = view.getUint32(4, true);
  const headerLength = view.getUint16(8, true);
  const recordLength = view.getUint16(10, true);
  if (!Number.isSafeInteger(recordCount) || headerLength < 33 || recordLength < 1 || headerLength > bytes.byteLength) return invalidNativeDocument('NATIVE_DBF_INVALID: header lengths are invalid');
  if (recordCount > limits.maxRecordCount) throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message: `DBF contains ${recordCount} records, above the ${limits.maxRecordCount} record limit` });
  const fields: Array<{ name: string; type: string; length: number; decimals: number }> = [];
  for (let offset = 32; offset + 32 <= headerLength && bytes[offset] !== 0x0d; offset += 32) {
    const name = decodeDbfText(bytes.slice(offset, offset + 11));
    const type = String.fromCharCode(bytes[offset + 11]!);
    const length = bytes[offset + 16]!;
    const decimals = bytes[offset + 17]!;
    if (!name || !length || fields.length >= 1024) return invalidNativeDocument('NATIVE_DBF_INVALID: field descriptor is invalid');
    fields.push({ name, type, length, decimals });
  }
  if (headerLength + recordCount * recordLength > bytes.byteLength + 1) return invalidNativeDocument('NATIVE_DBF_INVALID: records exceed source bytes');
  const rows: string[][] = [fields.map((field) => field.name)];
  for (let recordIndex = 0; recordIndex < recordCount; recordIndex += 1) {
    const start = headerLength + recordIndex * recordLength;
    const marker = bytes[start];
    if (marker === 0x1a) break;
    if (marker === 0x2a) continue;
    const row: string[] = [];
    let offset = start + 1;
    for (const field of fields) {
      const raw = bytes.slice(offset, offset + field.length);
      const value = decodeDbfText(raw).trim();
      row.push(field.type === 'L' ? (value.toUpperCase() === 'Y' || value.toUpperCase() === 'T' ? 'TRUE' : value.toUpperCase() === 'N' || value.toUpperCase() === 'F' ? 'FALSE' : '') : value);
      offset += field.length;
    }
    rows.push(row);
  }
  assertCellBudget(rows, limits, 'DBF document');
  return { rows, graph: { kind: 'dbf', table: { version, headerLength, recordLength, fields, recordCount, headerBytes: bytes.slice(0, headerLength) } } };
}

function decodeDbfText(bytes: Uint8Array): string {
  return new TextDecoder('windows-1252').decode(bytes).replaceAll('\0', '');
}

function blockedCodec(family: 'works' | 'web' | 'presentation', extensions: RegExp, description: string): NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> {
  return {
    family,
    canRead: (fileName) => extensions.test(fileName),
    import: async (request) => {
      throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_UNSUPPORTED', message: `${description} is detected but has no local record-native reader`, recovery: 'Use a codec with an explicit implementation for this protocol; the source bytes were not converted or discarded.' });
    },
    export: async () => {
      throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_UNSUPPORTED', message: `${description} has no local native writer`, recovery: 'Choose an explicitly supported target format.' });
    },
  };
}

export const dbfCodec: NativeDocumentCodec<NativeDocumentImportTransaction, NativeDocumentExportTransaction> = {
  family: 'dbf',
  canRead: (fileName) => /\.dbf$/i.test(fileName),
  import: async (request) => {
    const bytes = new Uint8Array(request.buffer);
    const limits = limitsFor(request.options);
    assertInputBudget(bytes, limits, 'DBF document');
    const parsed = parseDbf(bytes, limits);
    const snapshot = workbookFromRows(request.fileName.replace(/\.[^.]+$/, ''), parsed.rows);
    return importedResult(request.fileName, bytes, { family: 'dbf', variant: 'dbf' }, snapshot, parsed.graph, ['cells', 'dbf'], request.options.compatibilityTarget);
  },
  export: async () => {
    throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_UNSUPPORTED', message: 'DBF is open-only in the Excel file format matrix; no Save writer is exposed', recovery: 'Use Save As with an explicitly writable native document format.' });
  },
};
export const worksCodec = blockedCodec('works', /\.xlr$/i, 'Works spreadsheet');
export const webCodec = blockedCodec('web', /\.(?:htm|html|mht|mhtml)$/i, 'Office web document');
export const presentationCodec = blockedCodec('presentation', /\.(?:pdf|xps)$/i, 'presentation document');
