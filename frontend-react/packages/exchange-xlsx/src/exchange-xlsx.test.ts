import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { scanSnapshotFeatures } from './feature-scan';
import { exportSnapshotToXlsxBase64 } from './archive';
import { loadXlsxPackage, zipXlsxParts } from './archive';
import { strFromU8, strToU8 } from 'fflate';

describe('exchange-xlsx', () => {
  it('scans workbook features for compatibility reporting', () => {
    const workbook = new WorkbookModel('wb-xlsx', 'XLSX');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 10, formula: '=SUM(A2:A3)' });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
      anchor: { row: 1, column: 0 },
    });
    const features = scanSnapshotFeatures(workbook.snapshot());
    assert.ok(features.includes('cells'));
    assert.ok(features.includes('formulas'));
    assert.ok(features.includes('merges'));
  });

  it('scans canonical drawing payloads instead of removed per-kind collections', () => {
    const workbook = new WorkbookModel('wb-drawing-features', 'Drawing Features');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.drawings.push(
      {
        id: 'chart-drawing',
        sheetId: sheet.id,
        kind: 'chart',
        anchor: { kind: 'absolute' },
        transform: { x: 0, y: 0, width: 200, height: 120 },
        zIndex: 0,
        payloadId: 'chart-payload',
      },
      {
        id: 'image-drawing',
        sheetId: sheet.id,
        kind: 'image',
        anchor: { kind: 'one-cell', row: 0, column: 0 },
        transform: { x: 0, y: 0, width: 80, height: 40 },
        zIndex: 1,
        payloadId: 'image-payload',
      },
    );
    sheet.drawingPayloads.set('chart-payload', {
      kind: 'chart',
      chartId: 'chart-payload',
      chartType: 'line',
      sourceRanges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
    });
    sheet.drawingPayloads.set('image-payload', { kind: 'image', src: 'data:image/png;base64,AA==' });

    const features = scanSnapshotFeatures(workbook.snapshot());
    assert.ok(features.includes('charts'));
    assert.ok(features.includes('images'));
  });

  it('round-trips snapshot through xlsx archive import/export', async () => {
    const workbook = new WorkbookModel('wb-roundtrip', 'Roundtrip');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 'hello' });
    workbook.getSheet(workbook.primarySheetId).cells.set(1, 0, { value: 42, formula: '=A1&"!"' });
    const snapshot = workbook.snapshot();
    const base64 = exportSnapshotToXlsxBase64(snapshot);
    const imported = await importXlsx({
      fileName: 'roundtrip.xlsx',
      base64,
      options: { compatibilityTarget: 'B' },
    });
    assert.equal(imported.snapshot.sheets.length, snapshot.sheets.length);
    assert.equal(imported.report.schema, 'CompatibilityReport');
    const exported = await exportXlsx({
      snapshot: imported.snapshot,
      fileName: 'roundtrip.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    assert.ok(exported.base64.length > 0);
    assert.equal(exported.fileName, 'roundtrip.xlsx');
  });

  it('round-trips styles, merges, scoped names, freeze and the 1904 date system', async () => {
    const workbook = new WorkbookModel('wb-rich', 'Rich');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, {
      value: 45292,
      style: { bold: true, background: '#ff0000', horizontalAlignment: 'center' },
      numberFormat: 'm/d/yy',
    });
    sheet.cells.set(1, 1, { value: 'merged' });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 2 },
      anchor: { row: 1, column: 1 },
    });
    sheet.freeze = { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 };
    workbook.setDefinedName({ name: 'LocalValue', formula: `=${sheet.name}!$A$1`, scope: 'sheet', sheetId: sheet.id });
    const original = workbook.snapshot();
    const base64 = exportSnapshotToXlsxBase64(original, undefined, { dateSystem: '1904' });
    const imported = await importXlsx({ fileName: 'rich.xlsx', base64, options: { compatibilityTarget: 'A' } });
    const restored = imported.snapshot.sheets[0]!;
    assert.equal(imported.report.dateSystem, '1904');
    assert.equal(restored.merges.length, 1);
    assert.deepEqual(restored.freeze, sheet.freeze);
    assert.equal(restored.cells['0']?.['0']?.numberFormat, 'm/d/yy');
    assert.equal(restored.cells['0']?.['0']?.style?.bold, true);
    assert.equal(imported.snapshot.definedNameModels?.[0]?.scope, 'sheet');
  });

  it('preserves opaque chart/binary parts and relationships across an editable export', async () => {
    const workbook = new WorkbookModel('wb-preserve', 'Preserve');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 1 });
    const generated = loadXlsxPackage(exportSnapshotToXlsxBase64(workbook.snapshot()));
    generated.package.parts['xl/charts/chart1.xml'] = strToU8('<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><title>Keep</title></chartSpace>');
    generated.package.parts['xl/drawings/drawing1.xml'] = strToU8('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>');
    generated.package.parts['customXml/item1.bin'] = Uint8Array.from([0, 1, 2, 255]);
    generated.package.parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    generated.package.relationships['xl/worksheets/sheet1.xml'] = [{
      id: 'rIdChart',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
      target: '../drawings/drawing1.xml',
    }];
    generated.package.parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(generated.package.parts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<drawing r:id="rIdChart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></worksheet>'));
    // Rebuild through the public package writer so this test exercises the
    // same ZIP limits and relationship reader used by production imports.
    const imported = await importXlsx({ fileName: 'opaque.xlsx', base64: zipXlsxParts(generated.package.parts), options: { compatibilityTarget: 'B', preserveMacros: true } });
    const exported = await exportXlsx({ snapshot: imported.snapshot, package: imported.package, fileName: 'opaque.xlsx', options: { compatibilityTarget: 'B' } });
    const restored = loadXlsxPackage(exported.base64);
    assert.deepEqual([...restored.files['customXml/item1.bin']!], [0, 1, 2, 255]);
    assert.equal(strFromU8(restored.files['xl/charts/chart1.xml']!).includes('Keep'), true);
    assert.equal(strFromU8(restored.files['xl/worksheets/sheet1.xml']!).includes('rIdChart'), true);
    assert.equal(imported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
    assert.equal(exported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
  });

  it('rejects oversized and unsafe ZIP entries before inflation', () => {
    const workbook = new WorkbookModel('wb-limit', 'Limit');
    const base64 = exportSnapshotToXlsxBase64(workbook.snapshot());
    assert.throws(() => loadXlsxPackage(base64, { maxArchiveBytes: 10 }), /archive exceeds/);
    assert.throws(() => loadXlsxPackage(base64, { maxEntries: 1 }), /too many entries/);
    assert.throws(() => loadXlsxPackage(base64, { maxCompressionRatio: 1 }), /compression ratio/);
  });
});
