import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { scanSnapshotFeatures } from './feature-scan';
import { exportSnapshotToXlsxBuffer } from './archive';
import { loadXlsxPackage, parseLoadedXlsx, zipXlsxPartsBuffer } from './archive';
import { strFromU8, strToU8 } from 'fflate';
import { readFile } from 'node:fs/promises';

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
    const buffer = exportSnapshotToXlsxBuffer(snapshot);
    const imported = await importXlsx({
      fileName: 'roundtrip.xlsx',
      buffer,
      options: { compatibilityTarget: 'B' },
    });
    assert.equal(imported.snapshot.sheets.length, snapshot.sheets.length);
    assert.equal(imported.report.schema, 'CompatibilityReport');
    assert.equal(imported.sourceArtifact.schema, 'XlsxSourceArtifact');
    assert.equal(imported.sourceArtifact.checksum.length, 64);
    const exported = await exportXlsx({
      snapshot: imported.snapshot,
      fileName: 'roundtrip.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    assert.ok(exported.buffer.byteLength > 0);
    assert.equal(exported.fileName, 'roundtrip.xlsx');
  });

  it('round-trips worksheet tables and emits table relationship/content-type parts', async () => {
    const workbook = new WorkbookModel('wb-table', 'Table');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.sheetTables.push({
      id: 'table-1',
      sheetId: sheet.id,
      name: 'SalesTable',
      range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true,
      hasTotalRow: false,
      showBandedRows: true,
      showBandedColumns: false,
      showFilterButton: true,
      columns: [{ id: 'category', name: 'Category' }, { id: 'amount', name: 'Amount' }],
      styleName: 'TableStyleMedium2',
    });
    const imported = await importXlsx({ fileName: 'table.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.snapshot.sheets[0]?.sheetTables?.[0]?.name, 'SalesTable');
    assert.equal(imported.report.issues.find((issue) => issue.feature === 'tables')?.status, 'editable');
    const exported = await exportXlsx({ snapshot: imported.snapshot, package: imported.package, fileName: 'table.xlsx', options: { compatibilityTarget: 'B' } });
    const output = loadXlsxPackage(exported.buffer);
    assert.ok(output.files['xl/tables/table1.xml']);
    assert.match(strFromU8(output.files['[Content_Types].xml']!), /spreadsheetml\.table\+xml/);
    assert.match(strFromU8(output.files['xl/worksheets/_rels/sheet1.xml.rels']!), /\/table/);
  });

  it('writes hyperlinks from the canonical worksheet hyperlink collection', async () => {
    const workbook = new WorkbookModel('wb-links', 'Links');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'OpenAI' });
    sheet.hyperlinks.set('0:0', { id: 'link-1', target: { kind: 'url', url: 'https://openai.com/' }, tooltip: 'Open' });
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    const emitted = loadXlsxPackage(buffer);
    assert.match(strFromU8(emitted.files['xl/worksheets/sheet1.xml']!), /<hyperlink ref="A1"/);
    assert.match(strFromU8(emitted.files['xl/worksheets/_rels/sheet1.xml.rels']!), /https:\/\/openai\.com\//);
    const imported = await importXlsx({ fileName: 'links.xlsx', buffer, options: { compatibilityTarget: 'B' } });
    assert.equal(imported.snapshot.sheets[0]?.hyperlinks?.[0]?.hyperlink.target.kind, 'url');
  });

  it('writes a canonical table Pivot as a native cache and table graph', async () => {
    const workbook = new WorkbookModel('wb-new-pivot', 'New Pivot');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.cells.set(2, 0, { value: 'B' });
    sheet.cells.set(2, 1, { value: 20 });
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'SalesTable',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false, showFilterButton: true,
      columns: [{ id: 'category', name: 'Category' }, { id: 'amount', name: 'Amount' }], styleName: 'TableStyleMedium2',
    });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'sales-pivot', source: { kind: 'table', tableId: 'table-1' }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], values: [{ fieldId: 'amount', summarizeBy: 'sum' }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadXlsxPackage(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    assert.equal(output.package.nativePivotGraph?.caches.length, 1);
    assert.equal(output.package.nativePivotGraph?.tables.length, 1);
    assert.match(strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!), /worksheetSource name="SalesTable"/);
    assert.match(strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!), /<cacheSource type="worksheet">/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /pivotTableParts/);
    const imported = parseLoadedXlsx(output).snapshot;
    assert.equal(imported.sheets[0]?.pivots[0]?.source.kind, 'table');
  });

  it('writes and validates canonical Slicer and Timeline native parts', async () => {
    const workbook = new WorkbookModel('wb-native-controls', 'Native Controls');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Date' });
    sheet.cells.set(0, 2, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 45292 });
    sheet.cells.set(1, 2, { value: 10 });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'control-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 } }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'date', name: 'Date', dataType: 'date', ordinal: 1 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 2 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], values: [{ fieldId: 'amount', summarizeBy: 'sum' }], showSubtotals: true, showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push(
      { id: 'category-slicer', sheetId: sheet.id, kind: 'slicer', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 220, height: 180 }, zIndex: 1, payloadId: 'category-slicer' },
      { id: 'date-timeline', sheetId: sheet.id, kind: 'timeline', anchor: { kind: 'one-cell', row: 10, column: 0 }, transform: { x: 0, y: 0, width: 420, height: 120 }, zIndex: 1, payloadId: 'date-timeline' },
    );
    sheet.drawingPayloads.set('category-slicer', { kind: 'slicer', pivotId: 'control-pivot', fieldId: 'category', filter: { mode: 'all', memberKeys: [] }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
    sheet.drawingPayloads.set('date-timeline', { kind: 'timeline', pivotId: 'control-pivot', fieldId: 'date', period: { start: '2024-01-01T00:00:00Z', end: '2024-12-31T00:00:00Z' }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
    const output = loadXlsxPackage(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    assert.ok(output.files['xl/slicerCaches/slicerCache1.xml']);
    assert.ok(output.files['xl/slicers/slicer1.xml']);
    assert.ok(output.files['xl/timelineCaches/timelineCache1.xml']);
    assert.ok(output.files['xl/timelines/timeline1.xml']);
    const contentTypes = strFromU8(output.files['[Content_Types].xml']!);
    assert.match(contentTypes, /application\/vnd\.openxmlformats-officedocument\.drawing\+xml/);
    assert.match(contentTypes, /application\/vnd\.ms-excel\.slicerCache"/);
    assert.doesNotMatch(contentTypes, /application\/vnd\.ms-excel\.slicerCache\+xml/);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /slicerCaches/);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /timelineCacheRefs/);
    assert.match(strFromU8(output.files['xl/_rels/workbook.xml.rels']!), /relationships\/timelineCache/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /slicerList/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /timelineRefs/);
    const imported = parseLoadedXlsx(output).snapshot;
    assert.equal(Object.values(imported.sheets[0]?.drawingPayloads ?? {}).filter((payload) => payload.kind === 'slicer').length, 1);
    assert.equal(Object.values(imported.sheets[0]?.drawingPayloads ?? {}).filter((payload) => payload.kind === 'timeline').length, 1);
    const importedSlicerAnchor = imported.sheets[0]?.drawings.find((drawing) => drawing.kind === 'slicer')?.anchor;
    const importedTimelineAnchor = imported.sheets[0]?.drawings.find((drawing) => drawing.kind === 'timeline')?.anchor;
    assert.equal(importedSlicerAnchor?.kind, 'one-cell');
    assert.equal(importedSlicerAnchor?.kind === 'one-cell' ? importedSlicerAnchor.column : undefined, 4);
    assert.equal(importedTimelineAnchor?.kind === 'one-cell' ? importedTimelineAnchor.row : undefined, 10);

    const controlExport = await exportXlsx({ snapshot: workbook.snapshot(), fileName: 'native-controls.xlsx', options: { compatibilityTarget: 'B' } });
    assert.equal(controlExport.report.issues.some((issue) => issue.feature === 'images' && issue.status === 'unsupported'), false);

    const withoutControls = structuredClone(imported);
    const controlSheet = withoutControls.sheets[0]!;
    const controlDrawingIds = new Set(controlSheet.drawings.filter((drawing) => drawing.kind === 'slicer' || drawing.kind === 'timeline').map((drawing) => drawing.id));
    controlSheet.drawings = controlSheet.drawings.filter((drawing) => !controlDrawingIds.has(drawing.id));
    for (const id of controlDrawingIds) delete controlSheet.drawingPayloads[id];
    const removedControls = loadXlsxPackage(exportSnapshotToXlsxBuffer(withoutControls, output.package));
    const removedDrawing = removedControls.files['xl/drawings/drawing1.xml'];
    assert.ok(removedDrawing);
    assert.doesNotMatch(strFromU8(removedDrawing), /category[_-]slicer|date[_-]timeline/);
  });

  it('reads and rewrites native Pivot cache/table relationship graphs', async () => {
    const workbook = new WorkbookModel('wb-native-pivot', 'Native Pivot');
    const generated = loadXlsxPackage(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const parts = generated.package.parts;
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']!).replace('</workbook>', '<pivotCaches count="1"><pivotCache cacheId="1" r:id="rIdPivotCache"/></pivotCaches></workbook>'));
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>'));
    parts['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Sheet1"/></cacheSource><cacheFields count="2"><cacheField name="Category"><sharedItems containsString="1"><s v="A"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems containsNumber="1"><n v="10"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>');
    parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>');
    parts['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r><s v="0"/><n v="10"/></r></pivotCacheRecords>');
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<pivotTableParts count="1"><pivotTablePart r:id="rIdPivotTable"/></pivotTableParts></worksheet>'));
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'));
    parts['xl/pivotTables/pivotTable1.xml'] = strToU8('<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"><location ref="D1:E3"/><pivotFields count="2"><pivotField axis="axisRow"/><pivotField/></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1" name="Sum of Amount" subtotal="sum"/></dataFields></pivotTableDefinition>');
    const imported = await importXlsx({ fileName: 'native-pivot.xlsx', buffer: zipXlsxPartsBuffer(parts), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.package.nativePivotGraph?.caches[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.package.nativePivotGraph?.caches[0]?.fields[1]?.name, 'Amount');
    assert.equal(imported.package.nativePivotGraph?.tables[0]?.dataFields[0]?.field, 1);
    assert.equal(imported.snapshot.sheets[0]?.pivots.length, 1);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.fieldId, 'native:cache:1:field:0');
    assert.equal(imported.report.issues.find((issue) => issue.feature === 'pivot')?.status, 'editable');
    const exported = await exportXlsx({ snapshot: imported.snapshot, package: imported.package, fileName: 'native-pivot.xlsx', options: { compatibilityTarget: 'B' } });
    const output = loadXlsxPackage(exported.buffer);
    assert.ok(output.package.nativePivotGraph?.tables.length === 1);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /pivotCaches/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /pivotTableParts/);
    assert.match(strFromU8(output.files['[Content_Types].xml']!), /pivotTable/);
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
    sheet.pane = { kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 };
    workbook.setDefinedName({ name: 'LocalValue', formula: `=${sheet.name}!$A$1`, scope: 'sheet', sheetId: sheet.id });
    const original = workbook.snapshot();
    const buffer = exportSnapshotToXlsxBuffer(original, undefined, { dateSystem: '1904' });
    const imported = await importXlsx({ fileName: 'rich.xlsx', buffer, options: { compatibilityTarget: 'A' } });
    const restored = imported.snapshot.sheets[0]!;
    assert.equal(imported.report.dateSystem, '1904');
    assert.equal(restored.merges.length, 1);
    assert.deepEqual(restored.pane, original.sheets[0]!.pane);
    assert.equal(restored.cells['0']?.['0']?.numberFormat, 'm/d/yy');
    assert.equal(restored.cells['0']?.['0']?.style?.bold, true);
    assert.equal(imported.snapshot.definedNameModels?.[0]?.scope, 'sheet');
  });

  it('preserves opaque chart/binary parts and relationships across an editable export', async () => {
    const workbook = new WorkbookModel('wb-preserve', 'Preserve');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 1 });
    const generated = loadXlsxPackage(exportSnapshotToXlsxBuffer(workbook.snapshot()));
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
    const imported = await importXlsx({ fileName: 'opaque.xlsx', buffer: zipXlsxPartsBuffer(generated.package.parts), options: { compatibilityTarget: 'B', preserveMacros: true } });
    const exported = await exportXlsx({ snapshot: imported.snapshot, package: imported.package, fileName: 'opaque.xlsx', options: { compatibilityTarget: 'B' } });
    const restored = loadXlsxPackage(exported.buffer);
    assert.deepEqual([...restored.files['customXml/item1.bin']!], [0, 1, 2, 255]);
    assert.equal(strFromU8(restored.files['xl/charts/chart1.xml']!).includes('Keep'), true);
    assert.equal(strFromU8(restored.files['xl/worksheets/sheet1.xml']!).includes('rIdChart'), true);
    assert.equal(imported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
    assert.equal(exported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
  });

  it('rejects oversized and unsafe ZIP entries before inflation', () => {
    const workbook = new WorkbookModel('wb-limit', 'Limit');
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    assert.throws(() => loadXlsxPackage(buffer, { maxArchiveBytes: 10 }), /archive exceeds/);
    assert.throws(() => loadXlsxPackage(buffer, { maxEntries: 1 }), /too many entries/);
    assert.throws(() => loadXlsxPackage(buffer, { maxCompressionRatio: 1 }), /compression ratio/);
  });

  it('imports native OOXML geometry, shared formulas, split panes and worksheet capabilities', async () => {
    const workbook = new WorkbookModel('wb-native-geometry', 'Native Geometry');
    const generated = loadXlsxPackage(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    generated.package.parts['xl/styles.xml'] = strToU8('<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="18"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
    generated.package.parts['xl/worksheets/sheet1.xml'] = strToU8('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="3000" topLeftCell="B2" activePane="bottomRight" state="split"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="8.7109375" defaultRowHeight="15"/><cols><col min="1" max="1" width="9.625" customWidth="1"/></cols><sheetData><row r="1" ht="15" customHeight="1"><c r="A1"><v>2</v></c><c r="B1" s="1"><f t="shared" si="0" ref="B1:B2">A1*2</f><v>4</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><f t="shared" si="0"/><v>6</v></c></row></sheetData><autoFilter ref="A1:B2"/><conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" operator="greaterThan" priority="1"><formula>1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1:A2"><formula1>0</formula1></dataValidation></dataValidations></worksheet>');
    const imported = await importXlsx({ fileName: 'CAN配置.xlsx', buffer: zipXlsxPartsBuffer(generated.package.parts), options: { compatibilityTarget: 'B', compatibilityMode: 'balanced' } });
    const sheet = imported.snapshot.sheets[0]!;
    assert.equal(imported.snapshot.name, 'CAN配置');
    assert.ok((sheet.columnWidthsPx?.[0] ?? 0) >= 67);
    assert.equal(sheet.defaultColumnWidthPx, 61);
    assert.equal(sheet.rowHeightsPx?.[0], 20);
    assert.equal(sheet.cells['0']?.['1']?.style?.fontSizePx, 24);
    assert.equal(sheet.cells['1']?.['1']?.formula, '=A2*2');
    assert.equal(sheet.pane.kind, 'split');
    assert.equal(sheet.conditionalFormats?.length, 1);
    assert.equal(sheet.dataValidations?.length, 1);
    assert.equal(sheet.filter?.range.endColumn, 1);
  });

  it('resolves Strict relationship kinds and a non-standard workbook part path', () => {
    const parts: Record<string, Uint8Array> = {
      '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
      '_rels/.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="custom/book.xml"/></Relationships>'),
      'custom/book.xml': strToU8('<?xml version="1.0"?><workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Strict" sheetId="1" r:id="rIdSheet"/></sheets></workbook>'),
      'custom/_rels/book.xml.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/main.xml"/></Relationships>'),
      'custom/worksheets/main.xml': strToU8('<?xml version="1.0"?><worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetFormatPr defaultRowHeight="15" defaultColWidth="8.7109375"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row></sheetData></worksheet>'),
    };
    const loaded = loadXlsxPackage(zipXlsxPartsBuffer(parts));
    assert.equal(loaded.package.workbookPart, 'custom/book.xml');
    const parsed = parseLoadedXlsx(loaded, { workbookName: 'Strict' });
    assert.equal(parsed.snapshot.sheets[0]?.cells['0']?.['0']?.value, 'ok');
  });

  it('imports the repository House cleaning sample with Excel widths converted to CSS pixels', async () => {
    const bytes = await readFile(new URL('../../../../luckyexcel-node/House cleaning checklist.xlsx', import.meta.url));
    const imported = await importXlsx({ fileName: 'House cleaning checklist.xlsx', buffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), options: { compatibilityTarget: 'B', compatibilityMode: 'balanced' } });
    const widths = imported.snapshot.sheets.flatMap((sheet) => Object.values(sheet.columnWidthsPx ?? {}));
    assert.ok(widths.some((width) => width >= 67 && width <= 68), '9.625 Excel characters should render near 68 CSS pixels');
    assert.equal(widths.some((width) => Math.abs(width - 9.625) < 0.001), false);
  });
});
