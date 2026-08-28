import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { WorkbookModel } from '@react-sheets/core-model';
import { nativeDocumentCodecRegistry, odsCodec, sjsCodec, writeCfbPackage } from './index';
import { NativeDocumentError } from './native-document-error';
import { consumeNativeDocumentWorkerTask } from './worker-entry';
import { createNativeDocumentImportRequest } from './worker-protocol';

const options = { compatibilityTarget: 'B' as const };

function sourceSnapshot(name = 'Native test'): ReturnType<WorkbookModel['snapshot']> {
  const workbook = new WorkbookModel(`native-${name.toLowerCase().replace(/\s+/g, '-')}`, name);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Region' });
  sheet.cells.set(0, 1, { value: 'Amount' });
  sheet.cells.set(1, 0, { value: 'West' });
  sheet.cells.set(1, 1, { value: 42 });
  return workbook.snapshot();
}

function bytesOf(value: ArrayBuffer): Uint8Array {
  return new Uint8Array(value);
}

function biffRecord(type: number, payload: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(4 + payload.length);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, type, true);
  view.setUint16(2, payload.length, true);
  bytes.set(payload, 4);
  return bytes;
}

function biff12Record(type: number, payload: Uint8Array): Uint8Array {
  const encode = (value: number): Uint8Array => {
    const result: number[] = [];
    let current = value >>> 0;
    do {
      let byte = current & 0x7f;
      current >>>= 7;
      if (current) byte |= 0x80;
      result.push(byte);
    } while (current);
    return Uint8Array.from(result);
  };
  return Uint8Array.from([...encode(type), ...encode(payload.length), ...payload]);
}

function wideString(value: string): Uint8Array {
  const bytes = new Uint8Array(4 + value.length * 2);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, value.length, true);
  for (let index = 0; index < value.length; index += 1) view.setUint16(4 + index * 2, value.charCodeAt(index), true);
  return bytes;
}

function binaryCellHeader(column: number): Uint8Array {
  const bytes = new Uint8Array(8);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, column, true);
  return bytes;
}

function binaryReal(value: number): Uint8Array {
  const bytes = new Uint8Array(8);
  new DataView(bytes.buffer).setFloat64(0, value, true);
  return bytes;
}

function biffFixture(formula = false, version = 0x0600, sharedStrings = false): ArrayBuffer {
  const bof = new Uint8Array(16);
  new DataView(bof.buffer).setUint16(0, version, true);
  const globalBof = biffRecord(0x0809, bof);
  const sheetName = 'Sheet1';
  const boundPayload = new Uint8Array(version === 0x0500 ? 7 + sheetName.length : 8 + sheetName.length * 2);
  const boundView = new DataView(boundPayload.buffer);
  boundView.setUint32(0, 0, true);
  boundPayload[4] = 0;
  boundPayload[5] = 0;
  boundPayload[6] = sheetName.length;
  if (version === 0x0500) {
    for (let index = 0; index < sheetName.length; index += 1) boundPayload[7 + index] = sheetName.charCodeAt(index);
  } else {
    boundPayload[7] = 1;
    for (let index = 0; index < sheetName.length; index += 1) boundView.setUint16(8 + index * 2, sheetName.charCodeAt(index), true);
  }
  const bound = biffRecord(0x0085, boundPayload);
  const globalEof = biffRecord(0x000a, new Uint8Array());
  const sheetBof = biffRecord(0x0809, bof);
  const dimensions = biffRecord(0x0200, new Uint8Array(14));
  const numberPayload = new Uint8Array(14);
  const numberView = new DataView(numberPayload.buffer);
  numberView.setUint16(0, 0, true);
  numberView.setUint16(2, 0, true);
  numberView.setFloat64(6, 42, true);
  const number = biffRecord(formula ? 0x0006 : 0x0203, numberPayload);
  const labelPayload = new Uint8Array(version === 0x0500 ? 8 + 4 : 9 + 4 * 2);
  const labelView = new DataView(labelPayload.buffer);
  labelView.setUint16(2, 1, true);
  labelView.setUint16(6, 4, true);
  if (version === 0x0500) {
    for (let index = 0; index < 4; index += 1) labelPayload[8 + index] = 'West'.charCodeAt(index);
  } else {
    labelPayload[8] = 1;
    for (let index = 0; index < 4; index += 1) labelView.setUint16(9 + index * 2, 'West'.charCodeAt(index), true);
  }
  const label = sharedStrings
    ? biffRecord(0x00fd, Uint8Array.from([...labelPayload.slice(0, 6), 0, 0, 0, 0]))
    : biffRecord(0x0204, labelPayload);
  const sharedPayload = new Uint8Array(8 + 3 + 8);
  const sharedView = new DataView(sharedPayload.buffer);
  sharedView.setUint32(0, 1, true);
  sharedView.setUint32(4, 1, true);
  sharedView.setUint16(8, 4, true);
  sharedPayload[10] = 1;
  for (let index = 0; index < 4; index += 1) sharedView.setUint16(11 + index * 2, 'West'.charCodeAt(index), true);
  const shared = biffRecord(0x00fc, sharedPayload);
  const multiPayload = new Uint8Array(6 + 2 * 6);
  const multiView = new DataView(multiPayload.buffer);
  multiView.setUint16(0, 0, true);
  multiView.setUint16(2, 2, true);
  multiView.setUint16(4, 0, true);
  multiView.setUint32(6, (7 << 2) | 2, true);
  multiView.setUint16(10, 0, true);
  multiView.setUint32(12, (8 << 2) | 2, true);
  multiView.setUint16(16, 3, true);
  const multi = biffRecord(0x00bd, multiPayload);
  const sheetEof = biffRecord(0x000a, new Uint8Array());
  const globalBytes = Uint8Array.from([...globalBof, ...(sharedStrings ? [...shared] : []), ...bound, ...globalEof]);
  const boundOffset = globalBof.length + (sharedStrings ? shared.length : 0);
  new DataView(globalBytes.buffer).setUint32(boundOffset + 4, globalBytes.length, true);
  const stream = Uint8Array.from([...globalBytes, ...sheetBof, ...dimensions, ...number, ...label, ...multi, ...sheetEof]);
  const graph = {
    sectorSize: 512 as const,
    miniSectorSize: 64 as const,
    majorVersion: 3 as const,
    entries: [
      { name: 'Root Entry', type: 5 as const, left: -1, right: -1, child: 1, startSector: -2, size: 0, raw: new Uint8Array(128) },
      { name: 'Workbook', type: 2 as const, left: -1, right: 2, child: -1, startSector: -2, size: stream.length, raw: new Uint8Array(128) },
      { name: 'FutureStream', type: 2 as const, left: -1, right: -1, child: -1, startSector: -2, size: 3, raw: new Uint8Array(128) },
    ],
  };
  return writeCfbPackage(graph, { Workbook: stream, FutureStream: Uint8Array.from([9, 8, 7]) }, {}, {
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
  });
}

function xlsbFixture(): ArrayBuffer {
  const bundleName = wideString('Sheet1');
  const relationId = wideString('rId1');
  const bundlePayload = new Uint8Array(8 + relationId.length + bundleName.length);
  const bundleView = new DataView(bundlePayload.buffer);
  bundleView.setUint32(4, 1, true);
  bundlePayload.set(relationId, 8);
  bundlePayload.set(bundleName, 8 + relationId.length);
  const workbook = biff12Record(156, bundlePayload);
  const row = new Uint8Array(4);
  const sheet = Uint8Array.from([
    ...biff12Record(145, new Uint8Array()),
    ...biff12Record(0, row),
    ...biff12Record(5, Uint8Array.from([...binaryCellHeader(0), ...binaryReal(42)])),
    ...biff12Record(6, Uint8Array.from([...binaryCellHeader(1), ...wideString('West')])),
    ...biff12Record(146, new Uint8Array()),
    ...biff12Record(132, new Uint8Array()),
  ]);
  const relationships = new TextEncoder().encode('<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.bin"/></Relationships>');
  const contentTypes = new TextEncoder().encode('<?xml version="1.0"?><Types><Override PartName="/xl/workbook.bin" ContentType="binary.macroEnabled.main"/><Override PartName="/xl/worksheets/sheet1.bin" ContentType="sheet.binary"/></Types>');
  return zipSync({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.bin': workbook,
    'xl/_rels/workbook.bin.rels': relationships,
    'xl/worksheets/sheet1.bin': sheet,
    'vendor/opaque.bin': Uint8Array.from([4, 5, 6]),
  }).buffer as ArrayBuffer;
}

describe('native document codec registry', () => {
  it('detects each supported text and native package family at the one boundary', async () => {
    const workbook = sourceSnapshot();
    const ods = await odsCodec.export({ snapshot: workbook, fileName: 'book.ods', options });
    const sjs = await sjsCodec.export({ snapshot: workbook, fileName: 'book.sjs', options });
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.csv', strToU8('a,b\n1,2').buffer as ArrayBuffer).family, 'text');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.xml', strToU8('<?xml version="1.0"?><ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"/>').buffer as ArrayBuffer).family, 'xmlss');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.ods', ods.buffer).family, 'ods');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.sjs', sjs.buffer).family, 'sjs');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.ssjson', strToU8('{"schema":"SSJSON","sheets":[]}').buffer as ArrayBuffer).family, 'ssjson');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.dbf', new ArrayBuffer(0)).family, 'dbf');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.pdf', new ArrayBuffer(0)).family, 'presentation');
    const ooxml = await nativeDocumentCodecRegistry.export({ snapshot: workbook, fileName: 'book.xlsx', options, execution: 'inline-test' });
    assert.equal(nativeDocumentCodecRegistry.detectFormat('book.csv', ooxml.buffer).family, 'ooxml');
  });

  it('keeps CSV dialect and returns the original bytes for an untouched native Save', async () => {
    const input = strToU8('\uFEFFRegion;Amount\r\nWest;42\r\n');
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'report.csv', buffer: input.buffer as ArrayBuffer, options, execution: 'inline-test' });
    assert.equal(imported.artifact.nativeGraph.kind, 'text');
    if (imported.artifact.nativeGraph.kind !== 'text') throw new Error('text graph missing');
    assert.deepEqual(imported.artifact.nativeGraph.dialect, { encoding: 'utf-8', bom: true, delimiter: ';', rowDelimiter: '\r\n', quote: 'double', variant: 'csv' });
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'report.csv', options, execution: 'inline-test' });
    assert.deepEqual([...bytesOf(saved.buffer)], [...input]);
    imported.snapshot.sheets[0]!.cells['1']!['1']!.value = 43;
    const edited = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'report.csv', options, execution: 'inline-test' });
    assert.equal(strFromU8(bytesOf(edited.buffer)).includes('West;43'), true);
    assert.equal(bytesOf(edited.buffer)[0], 0xef);
  });

  it('guards text formula injection while restoring the canonical string on re-import', async () => {
    const workbook = sourceSnapshot('Formula');
    workbook.sheets[0]!.cells['0']!['0']!.value = '=SUM(1,2)';
    const exported = await nativeDocumentCodecRegistry.export({ snapshot: workbook, fileName: 'formula.csv', options, execution: 'inline-test' });
    assert.match(strFromU8(bytesOf(exported.buffer)), /'=SUM/);
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'formula.csv', buffer: exported.buffer, options, execution: 'inline-test' });
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['0']!.value, '=SUM(1,2)');
  });

  it('returns untouched OOXML source bytes only after the canonical projection is unchanged', async () => {
    const workbook = sourceSnapshot('OOXML');
    const source = await nativeDocumentCodecRegistry.export({ snapshot: workbook, fileName: 'source.xlsx', options, execution: 'inline-test' });
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'source.xlsx', buffer: source.buffer, options, execution: 'inline-test' });
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'source.xlsx', options, execution: 'inline-test' });
    assert.deepEqual([...bytesOf(saved.buffer)], [...bytesOf(source.buffer)]);
    imported.snapshot.sheets[0]!.cells['1']!['1']!.value = 43;
    const edited = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'source.xlsx', options, execution: 'inline-test' });
    assert.notDeepEqual([...bytesOf(edited.buffer)], [...bytesOf(source.buffer)]);
  });


  it('writes XML Spreadsheet 2003 directly without an OOXML ZIP envelope', async () => {
    const workbook = sourceSnapshot('XMLSS');
    const exported = await nativeDocumentCodecRegistry.export({ snapshot: workbook, fileName: 'report.xml', options, execution: 'inline-test' });
    const bytes = bytesOf(exported.buffer);
    assert.notEqual(bytes[0], 0x50);
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'report.xml', buffer: exported.buffer, options, execution: 'inline-test' });
    assert.equal(imported.artifact.nativeGraph.kind, 'xml');
    assert.equal(imported.snapshot.sheets[0]!.cells['1']!['1']!.value, 42);
  });

  it('round-trips UTF-16 text and SpreadJS SSJSON dataTable documents', async () => {
    const text = 'Region\tAmount\r\nWest\t42';
    const utf16 = new Uint8Array(text.length * 2 + 2);
    utf16.set([0xff, 0xfe]);
    const view = new DataView(utf16.buffer);
    for (let index = 0; index < text.length; index += 1) view.setUint16(index * 2 + 2, text.charCodeAt(index), true);
    const importedText = await nativeDocumentCodecRegistry.import({ fileName: 'unicode.txt', buffer: utf16.buffer as ArrayBuffer, options, execution: 'inline-test' });
    assert.equal(importedText.artifact.nativeGraph.kind, 'text');
    const textExport = await nativeDocumentCodecRegistry.export({ snapshot: importedText.snapshot, artifact: importedText.artifact, fileName: 'unicode.txt', options, execution: 'inline-test' });
    assert.deepEqual([...bytesOf(textExport.buffer)], [...utf16]);

    const json = strToU8(JSON.stringify({ version: '17.0.0', vendorExtension: { future: true }, sheets: [{ name: 'Sheet1', data: { dataTable: { '0': { '0': { value: 'Region' }, '1': { value: 'Amount' } }, '1': { '0': { value: 'West' }, '1': { value: 42 } } } } }] }));
    const importedJson = await nativeDocumentCodecRegistry.import({ fileName: 'book.ssjson', buffer: json.buffer as ArrayBuffer, options, execution: 'inline-test' });
    assert.equal(importedJson.snapshot.sheets[0]!.cells['1']!['1']!.value, 42);
    importedJson.snapshot.sheets[0]!.cells['1']!['1']!.value = 43;
    const jsonExport = await nativeDocumentCodecRegistry.export({ snapshot: importedJson.snapshot, artifact: importedJson.artifact, fileName: 'book.ssjson', options, execution: 'inline-test' });
    const jsonRoot = JSON.parse(strFromU8(bytesOf(jsonExport.buffer))) as Record<string, unknown>;
    assert.deepEqual(jsonRoot.vendorExtension, { future: true });
    assert.equal((jsonRoot.sheets as Array<Record<string, unknown>>)[0]!.data instanceof Array, true);
  });

  it('preserves unknown ODS and SJS package parts across an edited native Save', async () => {
    const workbook = sourceSnapshot('Packages');
    const ods = await odsCodec.export({ snapshot: workbook, fileName: 'book.ods', options });
    const odsParts = unzipSync(bytesOf(ods.buffer));
    odsParts['content.xml'] = strToU8(strFromU8(odsParts['content.xml']!).replace('</office:document-content>', '<vendor:futureExtension xmlns:vendor="urn:vendor:test">keep</vendor:futureExtension></office:document-content>'));
    odsParts['custom/opaque.bin'] = Uint8Array.from([1, 2, 3, 4]);
    const odsInput = zipSync(odsParts);
    const importedOds = await nativeDocumentCodecRegistry.import({ fileName: 'book.ods', buffer: odsInput.buffer as ArrayBuffer, options, execution: 'inline-test' });
    importedOds.snapshot.sheets[0]!.cells['1']!['1']!.value = 99;
    const odsEdited = await nativeDocumentCodecRegistry.export({ snapshot: importedOds.snapshot, artifact: importedOds.artifact, fileName: 'book.ods', options, execution: 'inline-test' });
    assert.deepEqual([...unzipSync(bytesOf(odsEdited.buffer))['custom/opaque.bin']!], [1, 2, 3, 4]);
    assert.match(strFromU8(unzipSync(bytesOf(odsEdited.buffer))['content.xml']!), /futureExtension/);

    const sjs = await sjsCodec.export({ snapshot: workbook, fileName: 'book.sjs', options });
    const sjsParts = unzipSync(bytesOf(sjs.buffer));
    sjsParts['vendor/opaque.json'] = strToU8('{"future":true}');
    const sjsInput = zipSync(sjsParts);
    const importedSjs = await nativeDocumentCodecRegistry.import({ fileName: 'book.sjs', buffer: sjsInput.buffer as ArrayBuffer, options, execution: 'inline-test' });
    importedSjs.snapshot.sheets[0]!.cells['1']!['1']!.value = 100;
    const sjsEdited = await nativeDocumentCodecRegistry.export({ snapshot: importedSjs.snapshot, artifact: importedSjs.artifact, fileName: 'book.sjs', options, execution: 'inline-test' });
    assert.equal(strFromU8(unzipSync(bytesOf(sjsEdited.buffer))['vendor/opaque.json']!), '{"future":true}');
  });

  it('reads, edits, and rewrites BIFF/CFB cells while retaining unknown streams', async () => {
    const input = biffFixture();
    assert.equal(nativeDocumentCodecRegistry.detectFormat('fixture.xls', input).family, 'biff');
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'fixture.xls', buffer: input, options, execution: 'inline-test' });
    assert.equal(imported.artifact.nativeGraph.kind, 'biff');
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['0']!.value, 42);
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['1']!.value, 'West');
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['2']!.value, 7);
    imported.snapshot.sheets[0]!.cells['0']!['0']!.value = 99;
    imported.snapshot.sheets[0]!.cells['0']!['2']!.value = 9;
    imported.snapshot.sheets[0]!.cells['0']!['4'] = { value: 'New' };
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'fixture.xls', options, execution: 'inline-test' });
    const reparsed = await nativeDocumentCodecRegistry.import({ fileName: 'fixture.xls', buffer: saved.buffer, options, execution: 'inline-test' });
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['0']!.value, 99);
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['2']!.value, 9);
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['3']!.value, 8);
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['4']!.value, 'New');
    assert.equal(reparsed.artifact.nativeGraph.kind, 'biff');
    if (reparsed.artifact.nativeGraph.kind !== 'biff') throw new Error('BIFF graph missing');
    assert.deepEqual([...reparsed.artifact.nativeGraph.container.streams!['FutureStream']!], [9, 8, 7]);
  });

  it('keeps the BIFF5 record grammar for legacy single-byte strings', async () => {
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'legacy.xls', buffer: biffFixture(false, 0x0500), options, execution: 'inline-test' });
    assert.equal(imported.artifact.format.variant, 'biff5');
    imported.snapshot.sheets[0]!.cells['0']!['1']!.value = 'East';
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'legacy.xls', options, execution: 'inline-test' });
    const reparsed = await nativeDocumentCodecRegistry.import({ fileName: 'legacy.xls', buffer: saved.buffer, options, execution: 'inline-test' });
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['1']!.value, 'East');
  });

  it('reads BIFF8 shared strings and converts only the edited cell record', async () => {
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'shared.xls', buffer: biffFixture(false, 0x0600, true), options, execution: 'inline-test' });
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['1']!.value, 'West');
    imported.snapshot.sheets[0]!.cells['0']!['1']!.value = 'East';
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'shared.xls', options, execution: 'inline-test' });
    const reparsed = await nativeDocumentCodecRegistry.import({ fileName: 'shared.xls', buffer: saved.buffer, options, execution: 'inline-test' });
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['1']!.value, 'East');
    assert.equal(reparsed.report.issues.some((issue) => issue.feature === 'sharedStrings'), true);
  });

  it('reads, edits, and rewrites BIFF12/XLSB cells while retaining unknown package parts', async () => {
    const input = xlsbFixture();
    assert.equal(nativeDocumentCodecRegistry.detectFormat('fixture.xlsb', input).family, 'xlsb');
    assert.equal(nativeDocumentCodecRegistry.detectFormat('misnamed.xlsx', input).family, 'xlsb');
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'fixture.xlsb', buffer: input, options, execution: 'inline-test' });
    assert.equal(imported.artifact.nativeGraph.kind, 'xlsb');
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['0']!.value, 42);
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['1']!.value, 'West');
    imported.snapshot.sheets[0]!.cells['0']!['0']!.value = 99;
    const saved = await nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'fixture.xlsb', options, execution: 'inline-test' });
    const reparsed = await nativeDocumentCodecRegistry.import({ fileName: 'fixture.xlsb', buffer: saved.buffer, options, execution: 'inline-test' });
    assert.equal(reparsed.snapshot.sheets[0]!.cells['0']!['0']!.value, 99);
    assert.deepEqual([...unzipSync(bytesOf(saved.buffer))['vendor/opaque.bin']!], [4, 5, 6]);
  });

  it('rejects edits that would flatten a native formula instead of silently losing its expression', async () => {
    const input = biffFixture(true);
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'formula.xls', buffer: input, options, execution: 'inline-test' });
    imported.snapshot.sheets[0]!.cells['0']!['0']!.value = 2;
    await assert.rejects(
      nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'formula.xls', options, execution: 'inline-test' }),
      (error: unknown) => error instanceof NativeDocumentError && error.code === 'NATIVE_DOCUMENT_UNSUPPORTED',
    );
  });

  it('executes the binary import through the production worker protocol', async () => {
    const response = await consumeNativeDocumentWorkerTask(createNativeDocumentImportRequest('binary-worker', 3, { fileName: 'fixture.xlsb', buffer: xlsbFixture(), options }));
    assert.equal(response.status, 'completed');
    if (response.status !== 'completed' || !('snapshot' in response.result)) throw new Error('Binary worker import did not complete');
    assert.equal(response.result.artifact.nativeGraph.kind, 'xlsb');
  });

  it('fails closed for malformed binary containers', async () => {
    await assert.rejects(
      nativeDocumentCodecRegistry.import({ fileName: 'legacy.xls', buffer: Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).buffer as ArrayBuffer, options, execution: 'inline-test' }),
      (error: unknown) => error instanceof NativeDocumentError && error.code === 'NATIVE_DOCUMENT_INVALID',
    );
  });

  it('opens a DBF table through its native header/record layout and keeps it open-only', async () => {
    const headerLength = 97;
    const recordLength = 11;
    const bytes = new Uint8Array(headerLength + recordLength + 1);
    const view = new DataView(bytes.buffer);
    view.setUint8(0, 0x03);
    view.setUint32(4, 1, true);
    view.setUint16(8, headerLength, true);
    view.setUint16(10, recordLength, true);
    const name = strToU8('Name');
    bytes.set(name, 32);
    bytes[43] = 0x43;
    bytes[48] = 10;
    bytes[64] = 0x0d;
    bytes[headerLength] = 0x20;
    bytes.set(strToU8('Alice'), headerLength + 1);
    bytes[headerLength + recordLength] = 0x1a;
    const imported = await nativeDocumentCodecRegistry.import({ fileName: 'people.dbf', buffer: bytes.buffer as ArrayBuffer, options, execution: 'inline-test' });
    assert.equal(imported.snapshot.sheets[0]!.cells['0']!['0']!.value, 'Name');
    assert.equal(imported.snapshot.sheets[0]!.cells['1']!['0']!.value, 'Alice');
    assert.equal(imported.report.issues.find((issue) => issue.feature === 'dbf')?.status, 'preserved-only');
    await assert.rejects(nativeDocumentCodecRegistry.export({ snapshot: imported.snapshot, artifact: imported.artifact, fileName: 'people.dbf', options, execution: 'inline-test' }), /NATIVE_DOCUMENT_UNSUPPORTED/);
  });

  it('rejects native documents that exceed the explicit resource budget', async () => {
    await assert.rejects(
      nativeDocumentCodecRegistry.import({ fileName: 'large.csv', buffer: strToU8('a,b').buffer as ArrayBuffer, options: { compatibilityTarget: 'B', limits: { maxArchiveBytes: 1 } }, execution: 'inline-test' }),
      (error: unknown) => error instanceof NativeDocumentError && error.code === 'NATIVE_DOCUMENT_RESOURCE_LIMIT',
    );
  });
});
