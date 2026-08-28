import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { WorkbookModel } from '@react-sheets/core-model';
import { nativeDocumentCodecRegistry, odsCodec, sjsCodec } from './index';
import { NativeDocumentError } from './native-document-error';

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

  it('fails closed for binary formats instead of routing through OOXML', async () => {
    await assert.rejects(
      nativeDocumentCodecRegistry.import({ fileName: 'legacy.xls', buffer: Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0]).buffer as ArrayBuffer, options, execution: 'inline-test' }),
      (error: unknown) => error instanceof NativeDocumentError && error.code === 'NATIVE_BINARY_CODEC_BLOCKED',
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
